from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
import aiosqlite

from app.database import get_db
from app.models import (
    ImportResponse, NarrativeUpdate, NarrativeBatchUpdate,
    PhaseDetail, NodeDetail, ProjectItem, NarrativeSection,
)
from app.services.importer import import_project, _clear_project_data
from app.services.report_templates import get_section_label

router = APIRouter(prefix="/api/projects", tags=["projects"])


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
    result = await import_project(db, content)
    return result


@router.get("", response_model=list[ProjectItem])
async def api_list_projects(db: aiosqlite.Connection = Depends(_db)):
    async with db.execute("""
        SELECT p.id, p.name, p.code, p.created_at, COUNT(ph.id)
        FROM projects p
        LEFT JOIN phases ph ON ph.project_id = p.id
        GROUP BY p.id ORDER BY p.id DESC
    """) as cur:
        rows = await cur.fetchall()
    return [
        ProjectItem(id=r[0], name=r[1], code=r[2], created_at=r[3] or "", phases_count=r[4])
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
            SELECT id, date, node_type FROM tracking_nodes
            WHERE phase_id=? ORDER BY date
        """, (ph[0],)) as cur:
            nodes = await cur.fetchall()
        result.append(PhaseDetail(
            id=ph[0], wbs=ph[1], name=ph[2],
            start_date=ph[3], end_date=ph[4],
            nodes=[NodeDetail(id=n[0], date=n[1], node_type=n[2]) for n in nodes],
        ))
    return result


# ── 叙述（Narratives）────────────────────────────────────────────────────────

@router.get("/nodes/{node_id}/report-types")
async def api_get_report_types(node_id: int, db: aiosqlite.Connection = Depends(_db)):
    """返回节点下的所有报告类型（用于前端构建 Tab）。"""
    async with db.execute("""
        SELECT DISTINCT report_type FROM narratives
        WHERE tracking_node_id=? ORDER BY report_type
    """, (node_id,)) as cur:
        rows = await cur.fetchall()
    return [r[0] for r in rows]


@router.get("/nodes/{node_id}/narratives/{report_type}",
            response_model=list[NarrativeSection])
async def api_get_narratives(
    node_id: int,
    report_type: str,
    db: aiosqlite.Connection = Depends(_db),
):
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
    await db.execute("""
        UPDATE narratives SET content=?
        WHERE tracking_node_id=? AND report_type=? AND section=?
    """, (body.content, node_id, report_type, section))
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
    for item in body.items:
        await db.execute("""
            UPDATE narratives SET content=?
            WHERE tracking_node_id=? AND report_type=? AND section=?
        """, (item.content, node_id, report_type, item.section))
    await db.commit()
    return {"ok": True, "updated": len(body.items)}


# ── 内部工具 ──────────────────────────────────────────────────────────────────

async def _require_project(db: aiosqlite.Connection, project_id: int):
    async with db.execute("SELECT id FROM projects WHERE id=?", (project_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, f"项目 {project_id} 不存在")
    return row
