"""
Excel 生成主入口。
build_excel(db, project_id, phase_id) -> bytes
"""
import io
import json
from datetime import date

import aiosqlite
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from app.services.report_templates import format_report_number, format_meeting_number
from .workreport       import build_workreport
from .progress_tracking import build_progress_tracking
from .measure_report    import build_measure_report
from .progress_report   import build_progress_report
from .meeting_sheet     import build_meeting_sheet


# ── 辅助：DB 查询 ────────────────────────────────────────────────────────────

async def _fetch_one(db, sql, params=()):
    async with db.execute(sql, params) as cur:
        return await cur.fetchone()


async def _fetch_all(db, sql, params=()):
    async with db.execute(sql, params) as cur:
        return await cur.fetchall()


def _to_date(v) -> date | None:
    if v is None:
        return None
    if isinstance(v, date):
        return v
    return date.fromisoformat(str(v))


# ── 辅助：偏差统计 ────────────────────────────────────────────────────────────

def _count_exceed(tnd: dict, node_id: int, tasks: list, threshold: float = 0.2):
    """统计超出阈值的工作量偏差任务数 / 进度偏差任务数。"""
    wl_exceed = 0
    sc_exceed  = 0
    for task in tasks:
        if not task.get("resources") or not task.get("plan_hours"):
            continue
        td = tnd.get((task["id"], node_id), {})
        actual = td.get("actual", 0)
        plan_should = td.get("plan_should", 0)
        plan_h = task["plan_hours"]
        if plan_h > 0 and abs(actual - plan_h) / plan_h > threshold:
            wl_exceed += 1
        if plan_should > 0 and abs(actual - plan_should) / plan_should > threshold:
            sc_exceed += 1
    return wl_exceed, sc_exceed


# ── 主入口 ───────────────────────────────────────────────────────────────────

async def build_excel(db: aiosqlite.Connection, project_id: int, phase_id: int) -> bytes:

    # 1. 项目信息
    row = await _fetch_one(
        db, "SELECT id,name,code,manager,cm_person,qa_person,team_members FROM projects WHERE id=?",
        (project_id,))
    project = dict(zip(["id","name","code","manager","cm_person","qa_person","team_members"], row))

    # 2. 阶段信息
    row = await _fetch_one(
        db, "SELECT id,wbs,name,start_date,end_date FROM phases WHERE id=?", (phase_id,))
    phase = {
        "id": row[0], "wbs": row[1], "name": row[2],
        "start_date": _to_date(row[3]), "end_date": _to_date(row[4]),
    }

    # 2b. 全项目起止时间（用于报告"项目周期"字段）
    rows = await _fetch_all(
        db, "SELECT start_date, end_date FROM phases WHERE project_id=?", (project_id,))
    all_starts = [_to_date(r[0]) for r in rows if r[0]]
    all_ends   = [_to_date(r[1]) for r in rows if r[1]]
    project["proj_start"] = min(all_starts) if all_starts else phase["start_date"]
    project["proj_end"]   = max(all_ends)   if all_ends   else phase["end_date"]

    # 3. 本阶段跟踪节点（按日期升序）
    rows = await _fetch_all(
        db, "SELECT id,date,node_type FROM tracking_nodes WHERE phase_id=? ORDER BY date",
        (phase_id,))
    phase_nodes = [{"id": r[0], "date": _to_date(r[1]), "type": r[2]} for r in rows]

    # 4. 全项目所有节点（按阶段 WBS 整数升序 → 节点日期升序），用于全局序号
    rows = await _fetch_all(db, """
        SELECT n.id
        FROM tracking_nodes n
        JOIN phases p ON p.id = n.phase_id
        WHERE p.project_id = ?
        ORDER BY CAST(p.wbs AS INTEGER), n.date
    """, (project_id,))
    all_node_ids = [r[0] for r in rows]
    node_positions = {nid: i + 1 for i, nid in enumerate(all_node_ids)}

    # 5. 所有任务
    rows = await _fetch_all(db, """
        SELECT id,wbs,name,start_date,end_date,plan_hours,resources
        FROM wbs_tasks WHERE phase_id=? ORDER BY wbs
    """, (phase_id,))
    tasks = [{
        "id": r[0], "wbs": r[1], "name": r[2],
        "start_date": _to_date(r[3]), "end_date": _to_date(r[4]),
        "plan_hours": r[5],
        "resources": json.loads(r[6] or "[]"),
    } for r in rows]

    # 6. task_node_data
    rows = await _fetch_all(db, """
        SELECT task_id,node_id,plan_should_complete,pct_complete,actual_complete
        FROM task_node_data
        WHERE task_id IN (SELECT id FROM wbs_tasks WHERE phase_id=?)
    """, (phase_id,))
    tnd = {(r[0], r[1]): {"plan_should": r[2], "pct": r[3], "actual": r[4]} for r in rows}

    # 7. Narratives
    rows = await _fetch_all(db, """
        SELECT tracking_node_id,report_type,section,content
        FROM narratives
        WHERE tracking_node_id IN (SELECT id FROM tracking_nodes WHERE phase_id=?)
    """, (phase_id,))
    narratives = {(r[0], r[1], r[2]): r[3] for r in rows}

    # 8. 配置项（当前阶段，含评审日期，按节点日期过滤）
    rows = await _fetch_all(db, """
        SELECT seq,name,identifier,version,item_type,actual_pages,plan_pages,review_date
        FROM config_items WHERE project_id=? AND phase_wbs=? ORDER BY seq
    """, (project_id, phase["wbs"]))
    config_items = [{"seq": r[0], "name": r[1], "identifier": r[2], "version": r[3],
                     "item_type": r[4], "actual_pages": r[5], "plan_pages": r[6],
                     "review_date": _to_date(r[7])} for r in rows]

    # 9. 代码模块（含阶段 WBS 和完成日期，按节点日期过滤）
    rows = await _fetch_all(db, """
        SELECT seq,name,actual_loc,plan_loc,phase_wbs,completion_date
        FROM code_modules WHERE project_id=? ORDER BY seq
    """, (project_id,))
    code_modules = [{"seq": r[0], "name": r[1], "actual_loc": r[2], "plan_loc": r[3],
                     "phase_wbs": r[4], "completion_date": _to_date(r[5])} for r in rows]

    # ── 工作量聚合 ───────────────────────────────────────────────────────────
    leaf_tasks = [t for t in tasks if t.get("resources") and t["plan_hours"] > 0]
    plan_hours_total = sum(t["plan_hours"] for t in leaf_tasks)

    def _phase_metrics(node_id):
        """返回 (plan_should, actual_complete) for the given node."""
        ps = sum(tnd.get((t["id"], node_id), {}).get("plan_should", 0) for t in leaf_tasks)
        ac = sum(tnd.get((t["id"], node_id), {}).get("actual",      0) for t in leaf_tasks)
        return round(ps, 2), round(ac, 2)

    # ── 累计工作量（里程碑版：含所有先前阶段最终节点）────────────────────────
    async def _cumulative_actual(current_phase_wbs: str, current_node_id: int) -> float:
        rows = await _fetch_all(db, """
            SELECT COALESCE(SUM(d.actual_complete), 0)
            FROM task_node_data d
            JOIN wbs_tasks t ON t.id = d.task_id
            JOIN phases p ON p.id = t.phase_id
            WHERE p.project_id = ?
              AND CAST(p.wbs AS INTEGER) < CAST(? AS INTEGER)
              AND d.node_id = (
                  SELECT id FROM tracking_nodes WHERE phase_id = p.id ORDER BY date DESC LIMIT 1
              )
        """, (project_id, current_phase_wbs))
        prev_actual = rows[0][0] if rows else 0.0
        _, cur_actual = _phase_metrics(current_node_id)
        return round(prev_actual + cur_actual, 2)

    # ── 构建工作簿 ────────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    is_phase1 = (phase["wbs"] == "1")

    # 阶段短名（去掉末尾"阶段"，用于里程碑 Sheet 标签）
    phase_short = phase["name"]
    if phase_short.endswith("阶段"):
        phase_short = phase_short[:-2]

    def _nd_str(d) -> str:
        return f"{d.year}.{d.month}.{d.day}"

    def _sn(name: str) -> str:
        """Sheet 名称不能超过 31 字符，截断时保留末尾括号。"""
        if len(name) <= 31:
            return name
        suffix = ""
        for end in ["）", "）", ")"]:
            if name.endswith(end):
                suffix = name[name.rfind("（"):]
                break
        core = name[: 31 - len(suffix)]
        return core + suffix

    # ── 预计算工时分布（所有任务×所有成员，一次生成，跨 WorkReport 复用）────
    from app.services.workhour_engine import distribute_task_hours
    hour_dists: dict = {}
    for task in tasks:
        if not task.get("resources") or not task.get("start_date") or not task.get("end_date"):
            continue
        for res in task["resources"]:
            hour_dists[(task["id"], res["name"])] = distribute_task_hours(
                task["wbs"], res["name"], res["hours"],
                task["start_date"], task["end_date"],
            )

    # ── Sheet 1..N: 每个跟踪节点一张 WorkReport ───────────────────────────────
    last_wr_name: str | None = None
    last_wr_rows: dict | None = None
    for node in phase_nodes:
        nd      = node["date"]
        wr_name = f"WorkReport{nd.strftime('%Y%m%d')}"
        wr_rows = build_workreport(
            wb, wr_name, tasks, phase, phase_nodes, tnd,
            node_date=nd, hour_dists=hour_dists,
        )
        last_wr_name = wr_name
        last_wr_rows = wr_rows

    # ── ProgressTracking（引用最后一个 WorkReport 的公式）────────────────────
    build_progress_tracking(
        wb, "ProgressTracking-ok", tasks, phase_nodes, tnd, phase,
        wr_sheet_name=last_wr_name, task_wr_rows=last_wr_rows,
    )

    # ProgressTracking 动态列游标（8 固定列 A-H 之后，从列 I=9 起）
    pt_col_cursor = 9

    # 全局例会序号起点
    first_node_id   = phase_nodes[0]["id"] if phase_nodes else None
    first_node_pos  = node_positions.get(first_node_id, 1)
    if first_node_pos == 1:          # 阶段1
        meeting_seq = 1
    else:
        meeting_seq = first_node_pos + 2

    # 前置节点（上一个跟踪节点）用于计算报告时间段起点
    all_phase_node_ids = [n["id"] for n in phase_nodes]
    prev_node_map: dict[int, date | None] = {}
    phase_start = phase["start_date"]
    for i, node in enumerate(phase_nodes):
        if i == 0:
            prev_node_map[node["id"]] = None   # 第一个节点无前驱
        else:
            prev_node_map[node["id"]] = phase_nodes[i - 1]["date"]

    # ── 逐节点生成报告 Sheet ──────────────────────────────────────────────────
    for node in phase_nodes:
        nid  = node["id"]
        nd   = node["date"]
        pos  = node_positions[nid]
        is_bw = (node["type"] == "双周")

        # ProgressTracking 公式引用（传给测量分析报告）
        if is_phase1:
            pt_ref_node = None
        elif is_bw:
            pt_ref_node = {
                "is_biweekly": True,
                "sheet": "ProgressTracking-ok",
                "plan_should_col": get_column_letter(pt_col_cursor),
                "actual_complete_col": get_column_letter(pt_col_cursor + 2),
                "row": 3,
            }
        else:
            pt_ref_node = {
                "is_biweekly": False,
                "sheet": "ProgressTracking-ok",
                "plan_wl_col": "H",
                "actual_hrs_col": get_column_letter(pt_col_cursor + 3),
                "plan_should_col": get_column_letter(pt_col_cursor),
                "actual_complete_col": get_column_letter(pt_col_cursor + 2),
                "row": 3,
            }

        # 按节点日期过滤：仅显示截止 nd 已评审/完成的文档和代码模块
        nd_config = [
            ci for ci in config_items
            if not ci.get("review_date") or ci["review_date"] <= nd
        ]
        nd_modules = [
            cm for cm in code_modules
            if cm.get("phase_wbs") == phase["wbs"]
            and (not cm.get("completion_date") or cm["completion_date"] <= nd)
        ]

        # 报告时间段
        prev_date = prev_node_map[nid]
        period_start = phase_start if prev_date is None else (
            date(prev_date.year, prev_date.month, prev_date.day)
        )
        period_end = nd

        # 阶段指标
        plan_should, actual_complete = _phase_metrics(nid)
        actual_hours = actual_complete  # 双周视为当期工时（简化）
        from app.services.report_templates import deviation_rate
        wl_dev   = deviation_rate(plan_hours_total, actual_complete) or 0.0
        sched_dev = round(actual_complete - plan_should, 2)
        wl_ex, sc_ex = _count_exceed(tnd, nid, leaf_tasks)

        def nar(report_type: str, section: str) -> str:
            return narratives.get((nid, report_type, section), "")

        # 当前节点处理完后推进 PT 列游标
        _pt_advance = 4 if is_bw else 6

        # ── 阶段1 ─────────────────────────────────────────────────────────────
        if is_phase1:
            rn_m = format_report_number("measure_phase1", project["code"], pos)
            rn_p = format_report_number("progress_phase1", project["code"], pos)

            build_measure_report(
                wb, f"{str(pos).zfill(3)}测量与分析报告",
                project, phase, nd, rn_m, "measure_phase1",
                period_start, period_end,
                nd_config, nd_modules,
                plan_hours_total, actual_complete, plan_should, actual_complete,
                0, 0,
                lambda s, _rt="measure_phase1": nar(_rt, s),
                pt_ref=pt_ref_node,
            )
            build_progress_report(
                wb, f"{str(pos).zfill(3)}进展报告",
                project, phase, nd, rn_p, "progress_phase1",
                period_start, period_end,
                wl_dev, sched_dev, wl_ex, sc_ex,
                lambda s, _rt="progress_phase1": nar(_rt, s),
            )

            # 3 个例会
            for i, prefix in enumerate(["lh001", "lh002", "lh003"]):
                seq = meeting_seq + i
                rn_mtg = format_meeting_number(project["code"], seq)
                def _get_nar(section, _p=prefix):
                    return narratives.get((nid, "meeting", f"{_p}_{section}"), "")
                build_meeting_sheet(
                    wb, f"LH{str(seq).zfill(3)}项目例会单",
                    project, phase, seq, rn_mtg, _get_nar,
                )
            meeting_seq += 3

        # ── 双周节点 ──────────────────────────────────────────────────────────
        elif is_bw:
            rn_m = format_report_number("measure_biweekly", project["code"], pos)
            rn_p = format_report_number("progress_biweekly", project["code"], pos)

            build_measure_report(
                wb, _sn(f"{str(pos).zfill(3)}测量与分析报告{_nd_str(nd)}（双周）"),
                project, phase, nd, rn_m, "measure_biweekly",
                period_start, period_end,
                nd_config, nd_modules,
                plan_hours_total, actual_complete, plan_should, actual_complete,
                0, 0,
                lambda s, _rt="measure_biweekly": nar(_rt, s),
                pt_ref=pt_ref_node,
            )
            build_progress_report(
                wb, _sn(f"{str(pos).zfill(3)}进展报告{_nd_str(nd)}（双周）"),
                project, phase, nd, rn_p, "progress_biweekly",
                period_start, period_end,
                wl_dev, sched_dev, wl_ex, sc_ex,
                lambda s, _rt="progress_biweekly": nar(_rt, s),
            )
            rn_mtg = format_meeting_number(project["code"], meeting_seq)
            def _nar_mtg(section):
                return narratives.get((nid, "meeting", section), "")
            build_meeting_sheet(
                wb, f"LH{str(meeting_seq).zfill(3)}项目例会单",
                project, phase, meeting_seq, rn_mtg, _nar_mtg,
            )
            meeting_seq += 1

        # ── 阶段节点 ──────────────────────────────────────────────────────────
        else:
            rn_mp = format_report_number("measure_phase",     project["code"], pos)
            rn_mm = format_report_number("measure_milestone",  project["code"], pos)
            rn_pp = format_report_number("progress_phase",     project["code"], pos)
            rn_pm = format_report_number("progress_milestone", project["code"], pos)

            # 里程碑版累计工时
            ms_actual = await _cumulative_actual(phase["wbs"], nid)

            build_measure_report(
                wb, _sn(f"{str(pos).zfill(3)}测量与分析报告{_nd_str(nd)}（阶段）"),
                project, phase, nd, rn_mp, "measure_phase",
                period_start, period_end,
                nd_config, nd_modules,
                plan_hours_total, actual_complete, plan_should, actual_complete,
                0, 0,
                lambda s, _rt="measure_phase": nar(_rt, s),
                pt_ref=pt_ref_node,
            )
            build_measure_report(
                wb, _sn(f"{str(pos).zfill(3)}测量与分析报告{_nd_str(nd)}（{phase_short}里程碑）"),
                project, phase, nd, rn_mm, "measure_milestone",
                phase_start, nd,
                nd_config, nd_modules,
                plan_hours_total, ms_actual, plan_should, ms_actual,
                0, 0,
                lambda s, _rt="measure_milestone": nar(_rt, s),
                pt_ref=pt_ref_node,
            )
            build_progress_report(
                wb, _sn(f"{str(pos).zfill(3)}进展报告{_nd_str(nd)}（阶段）"),
                project, phase, nd, rn_pp, "progress_phase",
                period_start, period_end,
                wl_dev, sched_dev, wl_ex, sc_ex,
                lambda s, _rt="progress_phase": nar(_rt, s),
            )
            build_progress_report(
                wb, _sn(f"{str(pos).zfill(3)}进展报告{_nd_str(nd)}（{phase_short}里程碑）"),
                project, phase, nd, rn_pm, "progress_milestone",
                phase_start, nd,
                wl_dev, sched_dev, wl_ex, sc_ex,
                lambda s, _rt="progress_milestone": nar(_rt, s),
            )
            rn_mtg = format_meeting_number(project["code"], meeting_seq)
            def _nar_mtg2(section):
                return narratives.get((nid, "meeting", section), "")
            build_meeting_sheet(
                wb, f"LH{str(meeting_seq).zfill(3)}项目例会单",
                project, phase, meeting_seq, rn_mtg, _nar_mtg2,
            )
            meeting_seq += 1

        pt_col_cursor += _pt_advance

    # ── 序列化 ───────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
