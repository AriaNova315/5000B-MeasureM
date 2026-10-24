import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import AsyncMock, patch

import aiosqlite
from fastapi import HTTPException

import app.database as database
from app.main import app
from app.routers import export
from app.services.importer import import_project
import run


ROOT = Path(__file__).resolve().parents[1]


class StartupTests(unittest.TestCase):
    def test_entrypoint_anchors_paths_and_binds_loopback(self):
        original = Path.cwd()
        try:
            with tempfile.TemporaryDirectory() as other, patch.object(run.uvicorn, "run") as server:
                os.chdir(other)
                run.main()
                self.assertEqual(ROOT, Path.cwd())
                server.assert_called_once_with("app.main:app", host="127.0.0.1", port=8002, reload=False)
        finally:
            os.chdir(original)

    def test_readonly_directory_reports_actionable_error(self):
        with patch.object(run.os, "chdir"), patch.object(run.tempfile, "TemporaryFile", side_effect=PermissionError):
            with self.assertRaisesRegex(SystemExit, "writable local directory"):
                run.main()

    def test_filename_keeps_legal_names_and_rejects_windows_hazards(self):
        self.assertEqual("02.需求分析阶段.xlsx", export._export_filename("2", "需求分析阶段"))
        for name in ['../escape', r'..\escape', 'bad:name', 'bad?name', 'bad\x00name', 'x' * 201, '', 'name.', 'name ']:
            with self.subTest(name=repr(name)), self.assertRaises(HTTPException) as error:
                export._export_filename("2", name)
            self.assertEqual(400, error.exception.status_code)
        with self.assertRaises(HTTPException):
            export._export_filename("CON", "name")

    def test_no_external_ui_resources_or_documentation_routes(self):
        html = (ROOT / "app/static/index.html").read_text(encoding="utf-8")
        self.assertNotIn("https://", html)
        self.assertIn('/static/vendor/alpine-3.14.9.min.js', html)
        self.assertTrue((ROOT / "app/static/vendor/alpine-3.14.9.min.js").is_file())
        self.assertIsNone(app.docs_url)
        self.assertIsNone(app.redoc_url)
        self.assertEqual("/openapi.json", app.openapi_url)


class ExportFailureTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.original_db, self.original_exports = database.DB_PATH, export.GENERATED_DIR
        database.DB_PATH = str(Path(self.temporary.name) / "test.db")
        export.GENERATED_DIR = str(Path(self.temporary.name) / "generated")
        await database.init_db()
        self.db = await aiosqlite.connect(database.DB_PATH)
        data = await import_project(self.db, (ROOT / "examples/demo-project.xlsx").read_bytes())
        self.project_id = data["project_id"]
        self.phase_id = (await (await self.db.execute("SELECT id FROM phases LIMIT 1")).fetchone())[0]

    async def asyncTearDown(self):
        await self.db.close()
        database.DB_PATH, export.GENERATED_DIR = self.original_db, self.original_exports
        self.temporary.cleanup()

    async def test_invalid_name_rejected_before_workbook_build(self):
        await self.db.execute("UPDATE phases SET name='bad/name' WHERE id=?", (self.phase_id,))
        with patch.object(export, "build_excel", new_callable=AsyncMock) as builder:
            with self.assertRaises(HTTPException) as error:
                await export.api_export_phase(self.project_id, self.phase_id, self.db)
            self.assertEqual(400, error.exception.status_code)
            builder.assert_not_awaited()

    async def test_file_lock_or_permission_and_disk_error_are_explained(self):
        for failure, status in [(PermissionError("locked"), 409), (OSError("disk full"), 500)]:
            with self.subTest(status=status), patch.object(export, "build_excel", new_callable=AsyncMock, return_value=b"test"), patch("app.routers.export.open", side_effect=failure, create=True):
                with self.assertRaises(HTTPException) as error:
                    await export.api_export_phase(self.project_id, self.phase_id, self.db)
                self.assertEqual(status, error.exception.status_code)


if __name__ == "__main__":
    unittest.main()
