import tempfile
import unittest
from pathlib import Path

import aiosqlite

import app.database as database
from app.main import index
from app.routers.projects import api_delete_project, api_list_projects
from app.services.importer import import_project


ROOT = Path(__file__).resolve().parents[1]
SAMPLE_INPUT = ROOT / "项目相关信息表.xlsx"


class ProjectImportListingTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.old_db_path = database.DB_PATH
        database.DB_PATH = str(Path(self.temp_dir.name) / "measurem.db")

    async def asyncTearDown(self):
        database.DB_PATH = self.old_db_path
        self.temp_dir.cleanup()

    async def test_init_db_adds_source_filename_to_legacy_projects_table(self):
        async with aiosqlite.connect(database.DB_PATH) as db:
            await db.execute("""
                CREATE TABLE projects (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    code TEXT,
                    created_at TEXT
                )
            """)
            await db.commit()

        await database.init_db()

        async with aiosqlite.connect(database.DB_PATH) as db:
            async with db.execute("PRAGMA table_info(projects)") as cur:
                columns = {row[1] for row in await cur.fetchall()}
        self.assertIn("source_filename", columns)

    async def test_index_disables_html_caching(self):
        response = await index()

        self.assertEqual("no-store", response.headers["cache-control"])

    async def test_import_persists_filename_and_list_returns_data_scale(self):
        await database.init_db()
        async with aiosqlite.connect(database.DB_PATH) as db:
            result = await import_project(
                db,
                SAMPLE_INPUT.read_bytes(),
                source_filename=SAMPLE_INPUT.name,
            )
            projects = await api_list_projects(db)

        self.assertEqual(1, len(projects))
        project = projects[0]
        self.assertEqual(result["project_id"], project.id)
        self.assertEqual(SAMPLE_INPUT.name, project.source_filename)
        self.assertEqual(result["phases_count"], project.phases_count)
        self.assertEqual(result["nodes_count"], project.nodes_count)
        self.assertEqual(result["tasks_count"], project.tasks_count)
        self.assertEqual(7, project.phases_count)
        self.assertEqual(10, project.nodes_count)
        self.assertEqual(193, project.tasks_count)

    async def test_delete_then_reimport_restores_project_listing(self):
        await database.init_db()
        content = SAMPLE_INPUT.read_bytes()
        async with aiosqlite.connect(database.DB_PATH) as db:
            first = await import_project(db, content, SAMPLE_INPUT.name)
            self.assertEqual(1, len(await api_list_projects(db)))

            await api_delete_project(first["project_id"], db)
            self.assertEqual([], await api_list_projects(db))

            second = await import_project(db, content, SAMPLE_INPUT.name)
            projects = await api_list_projects(db)

        self.assertGreater(second["project_id"], first["project_id"])
        self.assertEqual(1, len(projects))
        self.assertEqual(second["project_id"], projects[0].id)
        self.assertEqual(SAMPLE_INPUT.name, projects[0].source_filename)
        self.assertEqual(7, projects[0].phases_count)
        self.assertEqual(10, projects[0].nodes_count)
        self.assertEqual(193, projects[0].tasks_count)


if __name__ == "__main__":
    unittest.main()
