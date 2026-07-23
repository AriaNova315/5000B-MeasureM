import io
import tempfile
import unittest
from pathlib import Path

import aiosqlite
from openpyxl import load_workbook

import app.database as database
from app.models import NarrativeBatchItem, NarrativeBatchUpdate
from app.routers.projects import (
    api_batch_update_narratives,
    api_get_narratives,
    api_get_report_types,
)
from app.services.excel_builder.builder import build_excel
from app.services.importer import import_project


ROOT = Path(__file__).resolve().parents[1]
SAMPLE_INPUT = ROOT / "项目相关信息表.xlsx"


class NarrativeEffectiveContentTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.old_db_path = database.DB_PATH
        database.DB_PATH = str(Path(self.temp_dir.name) / "measurem.db")
        await database.init_db()
        self.db = await aiosqlite.connect(database.DB_PATH)
        result = await import_project(
            self.db,
            SAMPLE_INPUT.read_bytes(),
            source_filename=SAMPLE_INPUT.name,
        )
        self.project_id = result["project_id"]

        async with self.db.execute("""
            SELECT ph.id, ph.wbs, tn.id, tn.date
            FROM phases ph JOIN tracking_nodes tn ON tn.phase_id=ph.id
            WHERE ph.project_id=? ORDER BY CAST(ph.wbs AS INTEGER), tn.date
        """, (self.project_id,)) as cur:
            rows = await cur.fetchall()
        self.phase1_id, _, self.phase1_node_id, _ = rows[0]
        phase2_rows = [row for row in rows if row[1] == "2"]
        self.phase2_id = phase2_rows[0][0]
        self.phase2_first_node_id = phase2_rows[0][2]
        self.phase2_last_node_id = phase2_rows[-1][2]

    async def asyncTearDown(self):
        await self.db.close()
        database.DB_PATH = self.old_db_path
        self.temp_dir.cleanup()

    async def test_report_instances_match_business_meeting_names(self):
        phase1 = await api_get_report_types(self.phase1_node_id, self.db)
        phase1_meetings = [(item.id, item.label) for item in phase1 if item.id.startswith("meeting_")]
        self.assertEqual([
            ("meeting_lh001", "LH001 项目启动会"),
            ("meeting_lh002", "LH002 软件研制任务书研讨会"),
            ("meeting_lh003", "LH003 阶段例会"),
        ], phase1_meetings)

        phase2_first = await api_get_report_types(self.phase2_first_node_id, self.db)
        phase2_meetings = [(item.id, item.label) for item in phase2_first if item.id.startswith("meeting_")]
        self.assertEqual([
            ("meeting_planning", "LH004 项目策划会"),
            ("meeting_regular", "LH005 双周例会"),
        ], phase2_meetings)

        phase2_last = await api_get_report_types(self.phase2_last_node_id, self.db)
        self.assertIn(
            ("meeting_regular", "LH006 阶段例会"),
            [(item.id, item.label) for item in phase2_last],
        )

    async def test_generated_meeting_content_is_visible_without_becoming_override(self):
        sections = await api_get_narratives(
            self.phase2_first_node_id, "meeting_regular", self.db,
        )
        by_section = {section.section: section for section in sections}

        self.assertEqual("610会议室", by_section["location"].content)
        self.assertTrue(by_section["agenda"].content.startswith("1、项目组长通报"))
        self.assertIn("项目组成员报告个人工作情况", by_section["agenda"].content)
        self.assertFalse(any(section.is_overridden for section in sections))

    async def test_lh003_is_regular_phase_meeting_with_one_fewer_agenda_item(self):
        lh003 = await api_get_narratives(self.phase1_node_id, "meeting_lh003", self.db)
        agenda = next(section.content for section in lh003 if section.section == "lh003_agenda")

        self.assertEqual(5, len(agenda.splitlines()))
        self.assertNotIn("项目组成员报告个人工作情况", agenda)
        self.assertIn("质量保证人员通报", agenda)

    async def test_manual_override_controls_excel_and_empty_override_restores_default(self):
        override = NarrativeBatchUpdate(items=[
            NarrativeBatchItem(section="location", content="第一会议室"),
            NarrativeBatchItem(section="agenda", content="1、人工确认的例会议题。"),
        ])
        await api_batch_update_narratives(
            self.phase2_first_node_id, "meeting_regular", override, self.db,
        )

        edited = await api_get_narratives(
            self.phase2_first_node_id, "meeting_regular", self.db,
        )
        edited_by_section = {section.section: section for section in edited}
        self.assertTrue(edited_by_section["location"].is_overridden)
        self.assertEqual("第一会议室", edited_by_section["location"].content)

        workbook = load_workbook(io.BytesIO(
            await build_excel(self.db, self.project_id, self.phase2_id)
        ), data_only=False)
        self.assertEqual("第一会议室", workbook["LH005项目例会单"]["G3"].value)
        self.assertEqual("1、人工确认的例会议题。", workbook["LH005项目例会单"]["C6"].value)

        reset = NarrativeBatchUpdate(items=[
            NarrativeBatchItem(section="location", content=""),
            NarrativeBatchItem(section="agenda", content=""),
        ])
        await api_batch_update_narratives(
            self.phase2_first_node_id, "meeting_regular", reset, self.db,
        )
        restored = await api_get_narratives(
            self.phase2_first_node_id, "meeting_regular", self.db,
        )
        restored_by_section = {section.section: section for section in restored}
        self.assertFalse(restored_by_section["location"].is_overridden)
        self.assertEqual("610会议室", restored_by_section["location"].content)

    async def test_planning_meeting_override_is_persisted_and_exported(self):
        await api_batch_update_narratives(
            self.phase2_first_node_id,
            "meeting_planning",
            NarrativeBatchUpdate(items=[
                NarrativeBatchItem(section="planning_host", content="策划主持人"),
            ]),
            self.db,
        )

        workbook = load_workbook(io.BytesIO(
            await build_excel(self.db, self.project_id, self.phase2_id)
        ), data_only=False)
        self.assertEqual("策划主持人", workbook["LH004项目例会单"]["C4"].value)

    async def test_phase1_prefixed_override_controls_the_matching_meeting_only(self):
        await api_batch_update_narratives(
            self.phase1_node_id,
            "meeting_lh002",
            NarrativeBatchUpdate(items=[
                NarrativeBatchItem(section="lh002_location", content="任务书研讨室"),
            ]),
            self.db,
        )

        workbook = load_workbook(io.BytesIO(
            await build_excel(self.db, self.project_id, self.phase1_id)
        ), data_only=False)
        self.assertEqual("610会议室", workbook["LH001项目例会单"]["G3"].value)
        self.assertEqual("任务书研讨室", workbook["LH002项目例会单"]["G3"].value)
        self.assertEqual("610会议室", workbook["LH003项目例会单"]["G3"].value)


if __name__ == "__main__":
    unittest.main()
