import os
import re
from urllib.parse import quote
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import Response
import aiosqlite

from app.database import get_db
from app.services.excel_builder import build_excel

router = APIRouter(prefix="/api/export", tags=["export"])

GENERATED_DIR = "generated"


def _export_filename(phase_wbs: str, phase_name: str) -> str:
    filename = f"{phase_wbs.zfill(2)}.{phase_name}.xlsx"
    # Use the same rules on macOS and Windows; never silently rename a report.
    if (
        re.search(r'[<>:"/\\|?*\x00-\x1f]', filename)
        or filename.split(".", 1)[0].upper() in {
            "CON", "PRN", "AUX", "NUL",
            *(f"COM{i}" for i in range(1, 10)),
            *(f"LPT{i}" for i in range(1, 10)),
        }
        or len(filename.encode("utf-16-le")) // 2 > 200
        or not phase_name.strip()
        or phase_name.endswith((" ", "."))
    ):
        raise HTTPException(400, "阶段名称无法用作 Windows 文件名，请去掉特殊字符、末尾空格/句点或缩短名称后重新导入")
    return filename


async def _db():
    db = await get_db()
    try:
        yield db
    finally:
        await db.close()


@router.get("/{project_id}/{phase_id}")
async def api_export_phase(
    project_id: int,
    phase_id: int,
    db: aiosqlite.Connection = Depends(_db),
):
    # 验证阶段属于该项目
    async with db.execute(
        "SELECT wbs, name FROM phases WHERE id=? AND project_id=?",
        (phase_id, project_id),
    ) as cur:
        phase_row = await cur.fetchone()
    if not phase_row:
        raise HTTPException(404, f"阶段 {phase_id} 不存在或不属于项目 {project_id}")

    # 获取项目编号，用于文件命名
    async with db.execute("SELECT code FROM projects WHERE id=?", (project_id,)) as cur:
        proj_row = await cur.fetchone()
    if not proj_row:
        raise HTTPException(404, f"项目 {project_id} 不存在")

    filename = _export_filename(phase_row[0], phase_row[1])
    content = await build_excel(db, project_id, phase_id)

    # 保存到本地（可选，方便调试）
    out_dir = os.path.join(GENERATED_DIR, str(project_id))
    out_path = os.path.join(out_dir, filename)
    try:
        os.makedirs(out_dir, exist_ok=True)
        with open(out_path, "wb") as f:
            f.write(content)
    except PermissionError as exc:
        raise HTTPException(409, "无法保存导出文件，请关闭 Excel 中的同名文件，并检查 generated 目录写入权限后重试") from exc
    except OSError as exc:
        raise HTTPException(500, "无法保存导出文件，请检查磁盘空间、目录权限及路径长度后重试") from exc

    encoded_name = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}",
    }
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
