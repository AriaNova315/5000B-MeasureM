import os
from urllib.parse import quote
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import Response
import aiosqlite

from app.database import get_db
from app.services.excel_builder import build_excel

router = APIRouter(prefix="/api/export", tags=["export"])

GENERATED_DIR = "generated"


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

    content = await build_excel(db, project_id, phase_id)

    # 保存到本地（可选，方便调试）
    out_dir = os.path.join(GENERATED_DIR, str(project_id))
    os.makedirs(out_dir, exist_ok=True)
    phase_wbs = phase_row[0]
    phase_name = phase_row[1]
    filename = f"{phase_wbs.zfill(2)}.{phase_name}.xlsx"
    out_path = os.path.join(out_dir, filename)
    with open(out_path, "wb") as f:
        f.write(content)

    encoded_name = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}",
    }
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
