import json
import logging
from datetime import date

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
import aiosqlite

from app.database import get_db
from app.models import (
    ImportResponse, NarrativeUpdate, NarrativeBatchUpdate,
    PhaseDetail, NodeDetail, ProjectItem, NarrativeSection, NarrativeReportType,
)
from app.services.importer import import_project, _clear_project_data
from app.services.meeting_narratives import (
    MEETING_SECTIONS,
    meeting_defaults,
    meeting_instance_storage,
    meeting_sequence_start,
)
from app.services.progress_narratives import (
    PHASE1_PROGRESS_SECTIONS,
    phase1_progress_defaults,
)
from app.services.report_templates import get_section_label

router = APIRouter(prefix="/api/projects", tags=["projects"])
logger = logging.getLogger("uvicorn.error")


async def _db():
    db = await get_db()
    try:
        yield db
    finally:
        await db.close()


# ── 项目管理 ──────────────────────────────────────────────────────────────────

@router.post("/import", response_model=ImportResponse)
async def api_import(
    file: UploadFile = File(...),
    db: aiosqlite.Connection = Depends(_db),
):
    if not (file.filename or "").endswith(".xlsx"):
        raise HTTPException(400, "仅支持 .xlsx 文件")
    content = await file.read()
    logger.info("项目导入开始 filename=%r bytes=%d", file.filename, len(content))
    try:
        result = await import_project(db, content, source_filename=file.filename)
    except ValueError as exc:
        await db.rollback()
        logger.warning(
            "项目导入校验失败 filename=%r bytes=%d error=%s",
            file.filename, len(content), exc,
        )
        raise HTTPException(400, str(exc)) from exc
    except Exception:
        await db.rollback()
        logger.exception(
            "项目导入异常 filename=%r bytes=%d", file.filename, len(content),
        )
        raise
    logger.info(
        "项目导入成功 filename=%r bytes=%d project_id=%d phases=%d nodes=%d tasks=%d",
        file.filename,
        len(content),
        result["project_id"],
        result["phases_count"],
        result["nodes_count"],
        result["tasks_count"],
    )
    return result


@router.get("", response_model=list[ProjectItem])
async def api_list_projects(db: aiosqlite.Connection = Depends(_db)):
    async with db.execute("""
        SELECT
            p.id,
            p.name,
            p.code,
            p.source_filename,
            p.created_at,
            (SELECT COUNT(*) FROM phases ph WHERE ph.project_id = p.id),
            (SELECT COUNT(*)
             FROM tracking_nodes tn
             JOIN phases ph ON ph.id = tn.phase_id
             WHERE ph.project_id = p.id),
            (SELECT COUNT(*)
             FROM wbs_tasks wt
             JOIN phases ph ON ph.id = wt.phase_id
             WHERE ph.project_id = p.id)
        FROM projects p
        ORDER BY p.id DESC
    """) as cur:
        rows = await cur.fetchall()
    return [
        ProjectItem(
            id=r[0], name=r[1], code=r[2] or "", source_filename=r[3],
            created_at=r[4] or "", phases_count=r[5], nodes_count=r[6], tasks_count=r[7],
        )
        for r in rows
    ]


@router.delete("/{project_id}")
async def api_delete_project(project_id: int, db: aiosqlite.Connection = Depends(_db)):
    row = await _require_project(db, project_id)
    await _clear_project_data(db, project_id)
    await db.execute("DELETE FROM projects WHERE id=?", (project_id,))
    await db.commit()
    return {"ok": True}


# ── 阶段 & 节点 ───────────────────────────────────────────────────────────────

@router.get("/{project_id}/phases", response_model=list[PhaseDetail])
async def api_get_phases(project_id: int, db: aiosqlite.Connection = Depends(_db)):
    await _require_project(db, project_id)
    async with db.execute("""
        SELECT id, wbs, name, start_date, end_date FROM phases
        WHERE project_id=? ORDER BY CAST(wbs AS INTEGER)
    """, (project_id,)) as cur:
        phases = await cur.fetchall()

    result = []
    for ph in phases:
        async with db.execute("""
            SELECT id, date, node_type, node_name FROM tracking_nodes
            WHERE phase_id=? ORDER BY date,id
        """, (ph[0],)) as cur:
            nodes = await cur.fetchall()
        result.append(PhaseDetail(
            id=ph[0], wbs=ph[1], name=ph[2],
            start_date=ph[3], end_date=ph[4],
            nodes=[
                NodeDetail(id=n[0], date=n[1], node_type=n[2], node_name=n[3])
                for n in nodes
            ],
        ))
    return result


# ── 叙述（Narratives）────────────────────────────────────────────────────────

_REPORT_TYPE_LABELS = {
    "measure_phase1": "测量分析报告",
    "measure_biweekly": "测量分析（双周）",
    "measure_phase": "测量分析（阶段）",
    "measure_milestone": "测量分析（里程碑）",
    "progress_phase1": "进展报告",
    "progress_biweekly": "进展报告（双周）",
    "progress_phase": "进展报告（阶段）",
    "progress_milestone": "进展报告（里程碑）",
}
_REPORT_TYPE_ORDER = {report_type: index for index, report_type in enumerate([
    "measure_phase1", "measure_biweekly", "measure_phase", "measure_milestone",
    "progress_phase1", "progress_biweekly", "progress_phase", "progress_milestone",
])}


@router.get("/nodes/{node_id}/report-types", response_model=list[NarrativeReportType])
async def api_get_report_types(node_id: int, db: aiosqlite.Connection = Depends(_db)):
    """返回节点下的报告和可独立编辑的例会实例。"""
    context = await _load_narrative_context(db, node_id)
    async with db.execute("""
        SELECT DISTINCT report_type FROM narratives
        WHERE tracking_node_id=? ORDER BY report_type
    """, (node_id,)) as cur:
        rows = await cur.fetchall()

    report_types = sorted(
        (r[0] for r in rows if r[0] != "meeting"),
        key=lambda report_type: _REPORT_TYPE_ORDER.get(report_type, 999),
    )
    result = [
        NarrativeReportType(id=report_type, label=_REPORT_TYPE_LABELS.get(report_type, report_type))
        for report_type in report_types
    ]

    meeting_items = _meeting_report_items(context, node_id)
    meeting_reports = [NarrativeReportType(id=item_id, label=label) for item_id, label in meeting_items]
    insert_at = sum(1 for item in result if item.id.startswith("measure_"))
    result[insert_at:insert_at] = meeting_reports
    return result


@router.get("/nodes/{node_id}/narratives/{report_type}",
            response_model=list[NarrativeSection])
async def api_get_narratives(
    node_id: int,
    report_type: str,
    db: aiosqlite.Connection = Depends(_db),
):
    if report_type.startswith("meeting_"):
        context = await _load_narrative_context(db, node_id)
        valid_instances = {item_id for item_id, _ in _meeting_report_items(context, node_id)}
        if report_type not in valid_instances:
            raise HTTPException(404, f"节点 {node_id} 不包含例会实例 {report_type}")
        try:
            storage_type, prefix = meeting_instance_storage(report_type)
        except ValueError as exc:
            raise HTTPException(404, str(exc)) from exc
        defaults = meeting_defaults(
            report_type,
            context["project"],
            context["phase"],
            context["node"],
            context["tasks"],
        )
        async with db.execute("""
            SELECT section, content FROM narratives
            WHERE tracking_node_id=? AND report_type=?
        """, (node_id, storage_type)) as cur:
            overrides = {r[0]: r[1] or "" for r in await cur.fetchall()}
        return [
            NarrativeSection(
                section=f"{prefix}{section}",
                label=get_section_label("meeting", section),
                content=overrides.get(f"{prefix}{section}", "") or defaults[section],
                default_content=defaults[section],
                override_content=overrides.get(f"{prefix}{section}", ""),
                is_overridden=bool(overrides.get(f"{prefix}{section}", "")),
            )
            for section in MEETING_SECTIONS
        ]

    if report_type == "progress_phase1":
        context = await _load_narrative_context(db, node_id)
        if context["phase"]["wbs"] != "1":
            raise HTTPException(404, f"节点 {node_id} 不包含阶段1进展报告")
        config_items = await _phase1_config_items(db, context)
        defaults = phase1_progress_defaults(context["project"], config_items)
        async with db.execute("""
            SELECT section, content FROM narratives
            WHERE tracking_node_id=? AND report_type=?
        """, (node_id, report_type)) as cur:
            overrides = {r[0]: r[1] or "" for r in await cur.fetchall()}
        return [
            NarrativeSection(
                section=section,
                label=get_section_label(report_type, section),
                content=overrides.get(section, "") or defaults[section],
                default_content=defaults[section],
                override_content=overrides.get(section, ""),
                is_overridden=bool(overrides.get(section, "")),
            )
            for section in PHASE1_PROGRESS_SECTIONS
        ]

    async with db.execute("""
        SELECT section, content FROM narratives
        WHERE tracking_node_id=? AND report_type=? ORDER BY id
    """, (node_id, report_type)) as cur:
        rows = await cur.fetchall()
    if not rows:
        raise HTTPException(404, f"节点 {node_id} 的 {report_type} 报告不存在")
    return [
        NarrativeSection(
            section=r[0],
            label=get_section_label(report_type, r[0]),
            content=r[1] or "",
            override_content=r[1] or "",
            is_overridden=bool(r[1]),
        )
        for r in rows
    ]


@router.put("/nodes/{node_id}/narratives/{report_type}/{section}")
async def api_update_narrative(
    node_id: int,
    report_type: str,
    section: str,
    body: NarrativeUpdate,
    db: aiosqlite.Connection = Depends(_db),
):
    storage_type, allowed = await _narrative_storage(db, node_id, report_type)
    if section not in allowed:
        raise HTTPException(400, f"章节 {section} 不属于报告 {report_type}")
    await db.execute("""
        INSERT INTO narratives (tracking_node_id, report_type, section, content)
        VALUES (?,?,?,?)
        ON CONFLICT(tracking_node_id, report_type, section)
        DO UPDATE SET content=excluded.content
    """, (node_id, storage_type, section, body.content))
    await db.commit()
    return {"ok": True}


@router.put("/nodes/{node_id}/narratives/{report_type}")
async def api_batch_update_narratives(
    node_id: int,
    report_type: str,
    body: NarrativeBatchUpdate,
    db: aiosqlite.Connection = Depends(_db),
):
    """批量更新同一报告类型的多个 section（减少前端请求次数）。"""
    storage_type, allowed = await _narrative_storage(db, node_id, report_type)
    invalid = [item.section for item in body.items if item.section not in allowed]
    if invalid:
        raise HTTPException(400, f"章节 {invalid[0]} 不属于报告 {report_type}")
    for item in body.items:
        await db.execute("""
            INSERT INTO narratives (tracking_node_id, report_type, section, content)
            VALUES (?,?,?,?)
            ON CONFLICT(tracking_node_id, report_type, section)
            DO UPDATE SET content=excluded.content
        """, (node_id, storage_type, item.section, item.content))
    await db.commit()
    return {"ok": True, "updated": len(body.items)}


# ── 内部工具 ──────────────────────────────────────────────────────────────────

async def _require_project(db: aiosqlite.Connection, project_id: int):
    async with db.execute("SELECT id FROM projects WHERE id=?", (project_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, f"项目 {project_id} 不存在")
    return row


def _as_date(value) -> date | None:
    return date.fromisoformat(value) if value else None


def _meeting_report_items(context: dict, node_id: int) -> list[tuple[str, str]]:
    phase = context["phase"]
    node = context["node"]
    if node["node_type"] == "里程碑":
        return []
    if phase["wbs"] == "1":
        return [
            ("meeting_lh001", "LH001 项目启动会"),
            ("meeting_lh002", "LH002 软件研制任务书研讨会"),
            ("meeting_lh003", "LH003 阶段例会"),
        ]

    start = meeting_sequence_start(phase["wbs"], context["all_node_phase_wbs"])
    node_index = next(
        i for i, item in enumerate(context["phase_nodes"]) if item["id"] == node_id
    )
    items = []
    if phase["wbs"] == "2" and node_index == 0:
        items.append(("meeting_planning", f"LH{start:03d} 项目策划会"))
    regular_seq = start + node_index + (1 if phase["wbs"] == "2" else 0)
    meeting_name = "双周例会" if node["node_type"] == "双周" else "阶段例会"
    items.append(("meeting_regular", f"LH{regular_seq:03d} {meeting_name}"))
    return items


async def _load_narrative_context(db: aiosqlite.Connection, node_id: int) -> dict:
    async with db.execute("""
        SELECT
            tn.id, tn.date, tn.node_type,
            ph.id, ph.wbs, ph.name, ph.start_date, ph.end_date,
            p.id, p.name, p.code, p.manager, p.cm_person, p.qa_person, p.team_members
        FROM tracking_nodes tn
        JOIN phases ph ON ph.id=tn.phase_id
        JOIN projects p ON p.id=ph.project_id
        WHERE tn.id=?
    """, (node_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, f"节点 {node_id} 不存在")

    node = {"id": row[0], "date": _as_date(row[1]), "node_type": row[2]}
    phase = {
        "id": row[3], "wbs": row[4], "name": row[5],
        "start_date": _as_date(row[6]), "end_date": _as_date(row[7]),
    }
    project = {
        "id": row[8], "name": row[9], "code": row[10], "manager": row[11],
        "cm_person": row[12], "qa_person": row[13], "team_members": row[14],
    }

    async with db.execute("""
        SELECT id, wbs, name, start_date, end_date, resources
        FROM wbs_tasks WHERE phase_id=? ORDER BY wbs
    """, (phase["id"],)) as cur:
        task_rows = await cur.fetchall()
    tasks = [
        {
            "id": r[0], "wbs": r[1], "name": r[2],
            "start_date": _as_date(r[3]), "end_date": _as_date(r[4]),
            "resources": json.loads(r[5] or "[]"),
        }
        for r in task_rows
    ]

    async with db.execute("""
        SELECT tn.id, tn.date, tn.node_type
        FROM tracking_nodes tn
        WHERE tn.phase_id=? AND tn.node_type <> '里程碑'
        ORDER BY tn.date, tn.id
    """, (phase["id"],)) as cur:
        phase_node_rows = await cur.fetchall()
    phase_nodes = [
        {"id": r[0], "date": _as_date(r[1]), "node_type": r[2]}
        for r in phase_node_rows
    ]

    async with db.execute("""
        SELECT ph.wbs
        FROM tracking_nodes tn JOIN phases ph ON ph.id=tn.phase_id
        WHERE ph.project_id=? AND tn.node_type <> '里程碑'
        ORDER BY CAST(ph.wbs AS INTEGER), tn.date, tn.id
    """, (project["id"],)) as cur:
        all_node_phase_wbs = [r[0] for r in await cur.fetchall()]

    return {
        "project": project, "phase": phase, "node": node, "tasks": tasks,
        "phase_nodes": phase_nodes, "all_node_phase_wbs": all_node_phase_wbs,
    }


async def _phase1_config_items(db: aiosqlite.Connection, context: dict) -> list[dict]:
    """按阶段1跟踪日期读取与Excel生成器一致的有效配置项。"""
    async with db.execute("""
        SELECT name, item_type, actual_pages, review_date
        FROM config_items
        WHERE project_id=? AND phase_wbs=?
        ORDER BY seq
    """, (context["project"]["id"], context["phase"]["wbs"])) as cur:
        rows = await cur.fetchall()
    node_date = context["node"]["date"]
    return [
        {
            "name": row[0],
            "item_type": row[1],
            "actual_pages": row[2],
            "review_date": _as_date(row[3]),
        }
        for row in rows
        if not row[3] or _as_date(row[3]) <= node_date
    ]


async def _narrative_storage(
    db: aiosqlite.Connection,
    node_id: int,
    report_type: str,
) -> tuple[str, set[str]]:
    if report_type.startswith("meeting_"):
        context = await _load_narrative_context(db, node_id)
        valid_instances = {item_id for item_id, _ in _meeting_report_items(context, node_id)}
        if report_type not in valid_instances:
            raise HTTPException(404, f"节点 {node_id} 不包含例会实例 {report_type}")
        try:
            storage_type, prefix = meeting_instance_storage(report_type)
        except ValueError as exc:
            raise HTTPException(404, str(exc)) from exc
        return storage_type, {f"{prefix}{section}" for section in MEETING_SECTIONS}

    if report_type == "progress_phase1":
        context = await _load_narrative_context(db, node_id)
        if context["phase"]["wbs"] != "1":
            raise HTTPException(404, f"节点 {node_id} 不包含阶段1进展报告")
        return report_type, set(PHASE1_PROGRESS_SECTIONS)

    async with db.execute("""
        SELECT section FROM narratives
        WHERE tracking_node_id=? AND report_type=?
    """, (node_id, report_type)) as cur:
        sections = {r[0] for r in await cur.fetchall()}
    if not sections:
        raise HTTPException(404, f"节点 {node_id} 的 {report_type} 报告不存在")
    return report_type, sections
