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
from app.services.meeting_narratives import (
    meeting_defaults,
    meeting_progress_agenda as _meeting_progress_agenda,
    meeting_sequence_start as _meeting_sequence_start,
    task_progress_summary as _task_progress_summary,
)
from app.services.progress_narratives import phase1_progress_defaults
from app.services.workhour_engine import distribute_task_hours
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


def _wbs_index(v) -> int | None:
    try:
        return int(str(v))
    except (TypeError, ValueError):
        return None


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


def _progress_metric_refs(pt_ref: dict | None) -> dict:
    """返回进展报告应直接引用的 ProgressTracking 汇总指标。"""
    if not pt_ref:
        return {}
    sheet = pt_ref["sheet"].replace("'", "''")
    row = pt_ref["row"]
    progress_ref = f"='{sheet}'!{pt_ref['progress_dev_col']}{row}"
    if pt_ref.get("is_biweekly"):
        workload_ref = "——"
    else:
        workload_ref = f"='{sheet}'!{pt_ref['workload_dev_col']}{row}"
    return {
        "workload_deviation": workload_ref,
        "schedule_deviation": progress_ref,
    }


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
        db, "SELECT id,date,node_type,node_name FROM tracking_nodes WHERE phase_id=? ORDER BY date,id",
        (phase_id,))
    phase_nodes = [
        {"id": r[0], "date": _to_date(r[1]), "type": r[2], "name": r[3]}
        for r in rows
    ]
    tracking_nodes = [node for node in phase_nodes if node["type"] != "里程碑"]

    # 4. 全项目所有节点（按阶段 WBS 整数升序 → 节点日期升序），用于全局序号
    rows = await _fetch_all(db, """
        SELECT n.id, p.wbs, n.node_type
        FROM tracking_nodes n
        JOIN phases p ON p.id = n.phase_id
        WHERE p.project_id = ?
        ORDER BY CAST(p.wbs AS INTEGER), n.date, n.id
    """, (project_id,))
    all_node_ids = [r[0] for r in rows]
    all_node_phase_wbs = [r[1] for r in rows if r[2] != "里程碑"]
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

    # 8. 配置项（加载全项目，阶段/里程碑报告再分别按范围过滤）
    rows = await _fetch_all(db, """
        SELECT seq,name,identifier,version,item_type,actual_pages,plan_pages,review_date,phase_wbs
        FROM config_items WHERE project_id=?
        ORDER BY CAST(phase_wbs AS INTEGER), seq
    """, (project_id,))
    config_items = [{"seq": r[0], "name": r[1], "identifier": r[2], "version": r[3],
                     "item_type": r[4], "actual_pages": r[5], "plan_pages": r[6],
                     "review_date": _to_date(r[7]), "phase_wbs": r[8]} for r in rows]

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

    # ── 里程碑累计偏移量（所有前序阶段的最终节点）────────────────────────────
    async def _prior_milestone_metrics(current_phase_wbs: str) -> dict[str, float]:
        rows = await _fetch_all(db, """
            SELECT
                p.wbs, t.wbs, t.start_date, t.end_date, t.plan_hours, t.resources,
                d.plan_should_complete, d.actual_complete
            FROM phases p
            JOIN wbs_tasks t ON t.phase_id = p.id
            JOIN tracking_nodes n ON n.id = (
                SELECT id FROM tracking_nodes
                WHERE phase_id = p.id AND node_type <> '里程碑'
                ORDER BY date DESC, id DESC LIMIT 1
            )
            LEFT JOIN task_node_data d ON d.task_id = t.id AND d.node_id = n.id
            WHERE p.project_id = ?
              AND CAST(p.wbs AS INTEGER) < CAST(? AS INTEGER)
              AND t.resources <> '[]'
              AND t.plan_hours > 0
        """, (project_id, current_phase_wbs))

        totals = {
            "plan_workload": 0.0,
            "actual_hours": 0.0,
            "plan_should": 0.0,
            "actual_complete": 0.0,
        }
        for phase_wbs, task_wbs, start, end, plan_hours, resources, plan_should, actual in rows:
            actual = float(actual or 0)
            if str(phase_wbs) == "1":
                # 阶段1没有计划工作量，按业务规则以实际工作量补齐计划口径。
                totals["plan_workload"] += actual
                totals["plan_should"] += actual
            else:
                totals["plan_workload"] += float(plan_hours or 0)
                totals["plan_should"] += float(plan_should or 0)
            totals["actual_complete"] += actual

            task_start = _to_date(start)
            task_end = _to_date(end)
            for resource in json.loads(resources or "[]"):
                distribution = distribute_task_hours(
                    task_wbs,
                    resource["name"],
                    float(resource.get("hours") or 0),
                    task_start,
                    task_end,
                )
                totals["actual_hours"] += sum(distribution.values())

        return {key: round(value, 2) for key, value in totals.items()}

    # ── 构建工作簿 ────────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)
    # 报告分析区使用动态数组公式；确保 Excel 打开文件时自动重算。
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

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
    workreport_refs: list[dict] = []
    for node in tracking_nodes:
        nd      = node["date"]
        wr_name = f"WorkReport{nd.strftime('%Y%m%d')}"
        wr_rows = build_workreport(
            wb, wr_name, tasks, phase, tracking_nodes, tnd,
            node_date=nd, hour_dists=hour_dists,
        )
        last_wr_name = wr_name
        last_wr_rows = wr_rows
        workreport_refs.append({"sheet": wr_name, "task_wr_rows": wr_rows})

    # ── ProgressTracking（引用最后一个 WorkReport 的公式）────────────────────
    build_progress_tracking(
        wb, "ProgressTracking-ok", tasks, tracking_nodes, tnd, phase,
        wr_sheet_name=last_wr_name, task_wr_rows=last_wr_rows,
        workreport_refs=workreport_refs,
    )

    # 里程碑不在 ProgressTracking 中增加列组，它引用关联阶段最后一个跟踪节点。
    pt_refs: dict[int, dict] = {}
    pt_col_cursor = 9
    for tracking_node in tracking_nodes:
        if is_phase1:
            pt_ref = {
                "is_biweekly": False,
                "sheet": "ProgressTracking-ok",
                "actual_complete_col": "E",
                "row": 3,
            }
        elif tracking_node["type"] == "双周":
            pt_ref = {
                "is_biweekly": True,
                "sheet": "ProgressTracking-ok",
                "plan_should_col": get_column_letter(pt_col_cursor),
                "actual_complete_col": get_column_letter(pt_col_cursor + 2),
                "workload_dev_col": get_column_letter(pt_col_cursor + 3),
                "progress_dev_col": get_column_letter(pt_col_cursor + 3),
                "data_start": 4,
                "data_end": len(tasks) + 3,
                "row": 3,
            }
        else:
            pt_ref = {
                "is_biweekly": False,
                "sheet": "ProgressTracking-ok",
                "plan_wl_col": "H",
                "actual_hrs_col": get_column_letter(pt_col_cursor + 3),
                "plan_should_col": get_column_letter(pt_col_cursor),
                "actual_complete_col": get_column_letter(pt_col_cursor + 2),
                "progress_dev_col": get_column_letter(pt_col_cursor + 4),
                "workload_dev_col": get_column_letter(pt_col_cursor + 5),
                "data_start": 4,
                "data_end": len(tasks) + 3,
                "row": 3,
            }
        pt_refs[tracking_node["id"]] = pt_ref
        pt_col_cursor += 4 if tracking_node["type"] == "双周" else 6

    # 全局例会序号起点：阶段1固定3张，阶段2另有1张项目策划会。
    meeting_seq = _meeting_sequence_start(phase["wbs"], all_node_phase_wbs)

    # 前置节点（上一个跟踪节点）仅用于计算双周报告时间段起点
    prev_node_map: dict[int, date | None] = {}
    phase_start = phase["start_date"]
    project_start = project.get("proj_start") or phase_start
    for i, node in enumerate(tracking_nodes):
        if i == 0:
            prev_node_map[node["id"]] = None   # 第一个节点无前驱
        else:
            prev_node_map[node["id"]] = tracking_nodes[i - 1]["date"]

    # 阶段2在普通跟踪节点例会前增加一张项目策划例会单，并占用连续流水号。
    if phase["wbs"] == "2":
        planning_seq = meeting_seq
        planning_report_number = format_meeting_number(project["code"], planning_seq)
        planning_node = tracking_nodes[0]
        planning_defaults = meeting_defaults(
            "meeting_planning", project, phase, planning_node, tasks,
        )

        def _planning_nar(section):
            override = narratives.get(
                (planning_node["id"], "meeting", f"planning_{section}"), ""
            )
            return override or planning_defaults.get(section, "")

        build_meeting_sheet(
            wb, f"LH{str(planning_seq).zfill(3)}项目例会单",
            project, phase, planning_seq, planning_report_number, _planning_nar,
        )
        meeting_seq += 1

    # ── 逐节点生成报告 Sheet ──────────────────────────────────────────────────
    for node in phase_nodes:
        nid  = node["id"]
        nd   = node["date"]
        pos  = node_positions[nid]
        is_milestone = (node["type"] == "里程碑")
        is_bw = (node["type"] == "双周")

        # 里程碑引用关联阶段中不晚于里程碑日期的最后一个普通跟踪节点。
        metric_node = node
        if is_milestone:
            candidates = [item for item in tracking_nodes if item["date"] <= nd]
            if not candidates:
                raise ValueError(f"里程碑“{node.get('name') or nd}”前没有可用的阶段跟踪节点")
            metric_node = candidates[-1]
        pt_ref_node = pt_refs[metric_node["id"]]

        # 按节点日期过滤：仅显示截止 nd 已评审/完成的文档和代码模块
        nd_config = [
            ci for ci in config_items
            if ci.get("phase_wbs") == phase["wbs"]
            and (not ci.get("review_date") or ci["review_date"] <= nd)
        ]
        nd_modules = [
            cm for cm in code_modules
            if cm.get("phase_wbs") == phase["wbs"]
            and (not cm.get("completion_date") or cm["completion_date"] <= nd)
        ]

        # 报告时间段：双周按相邻节点区间，阶段按整个阶段区间。
        prev_date = prev_node_map.get(nid)
        if is_bw:
            period_start = phase_start if prev_date is None else prev_date
        else:
            period_start = phase_start
        period_end = nd

        # 阶段指标
        plan_should, actual_complete = _phase_metrics(metric_node["id"])
        actual_hours = actual_complete  # 双周视为当期工时（简化）
        from app.services.report_templates import deviation_rate
        wl_dev   = deviation_rate(plan_hours_total, actual_complete) or 0.0
        sched_dev = round(actual_complete - plan_should, 2)
        wl_ex, sc_ex = _count_exceed(tnd, metric_node["id"], leaf_tasks)

        def nar(report_type: str, section: str) -> str:
            return narratives.get((nid, report_type, section), "")

        # ── 阶段1 ─────────────────────────────────────────────────────────────
        if is_phase1 and not is_milestone:
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
            progress_defaults = phase1_progress_defaults(project, nd_config)

            def _get_phase1_progress_nar(section, _defaults=progress_defaults):
                override = nar("progress_phase1", section)
                return override or _defaults.get(section, "")

            build_progress_report(
                wb, f"{str(pos).zfill(3)}进展报告",
                project, phase, nd, rn_p, "progress_phase1",
                period_start, period_end,
                wl_dev, sched_dev, wl_ex, sc_ex,
                _get_phase1_progress_nar,
                config_items=nd_config,
            )

            # 3 个例会
            for i, prefix in enumerate(["lh001", "lh002", "lh003"]):
                seq = meeting_seq + i
                rn_mtg = format_meeting_number(project["code"], seq)
                defaults = meeting_defaults(
                    f"meeting_{prefix}", project, phase, node, tasks,
                )
                def _get_nar(section, _p=prefix, _defaults=defaults):
                    existing = narratives.get((nid, "meeting", f"{_p}_{section}"), "")
                    return existing or _defaults.get(section, "")
                build_meeting_sheet(
                    wb, f"LH{str(seq).zfill(3)}项目例会单",
                    project, phase, seq, rn_mtg, _get_nar,
                )
            meeting_seq += 3

        # ── 双周节点 ──────────────────────────────────────────────────────────
        elif is_bw:
            rn_m = format_report_number("measure_biweekly", project["code"], pos)
            rn_p = format_report_number("progress_biweekly", project["code"], pos)

            measure_sheet = _sn(f"{str(pos).zfill(3)}测量与分析报告{_nd_str(nd)}（双周）")
            measure_metrics = build_measure_report(
                wb, measure_sheet,
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
                metric_refs=_progress_metric_refs(pt_ref_node),
                pt_ref=pt_ref_node,
            )
            rn_mtg = format_meeting_number(project["code"], meeting_seq)
            regular_defaults = meeting_defaults(
                "meeting_regular", project, phase, node, tasks,
            )
            def _nar_mtg(section):
                existing = narratives.get((nid, "meeting", section), "")
                return existing or regular_defaults.get(section, "")
            build_meeting_sheet(
                wb, f"LH{str(meeting_seq).zfill(3)}项目例会单",
                project, phase, meeting_seq, rn_mtg, _nar_mtg,
            )
            meeting_seq += 1

        # ── 阶段节点 ──────────────────────────────────────────────────────────
        elif is_milestone:
            rn_mm = format_report_number("measure_milestone", project["code"], pos)
            rn_pm = format_report_number("progress_milestone", project["code"], pos)

            prior_metrics = await _prior_milestone_metrics(phase["wbs"])
            current_phase_index = _wbs_index(phase["wbs"])
            milestone_config = [
                ci for ci in config_items
                if _wbs_index(ci.get("phase_wbs")) is not None
                and _wbs_index(ci.get("phase_wbs")) <= current_phase_index
                and (not ci.get("review_date") or ci["review_date"] <= nd)
            ]
            milestone_modules = [
                cm for cm in code_modules
                if _wbs_index(cm.get("phase_wbs")) is not None
                and _wbs_index(cm.get("phase_wbs")) <= current_phase_index
                and (not cm.get("completion_date") or cm["completion_date"] <= nd)
            ]
            milestone_pt_ref = (
                None
                if is_phase1
                else {**pt_ref_node, "cumulative_offsets": prior_metrics}
            )
            milestone_label = node.get("name") or f"{phase_short}里程碑"
            milestone_plan_hours = actual_complete if is_phase1 else plan_hours_total
            milestone_plan_should = actual_complete if is_phase1 else plan_should

            build_measure_report(
                wb, _sn(f"{str(pos).zfill(3)}测量与分析报告{_nd_str(nd)}（{milestone_label}）"),
                project, phase, nd, rn_mm, "measure_milestone",
                project_start, nd,
                milestone_config, milestone_modules,
                milestone_plan_hours, actual_complete,
                milestone_plan_should, actual_complete,
                0, 0,
                lambda s, _rt="measure_milestone": nar(_rt, s),
                pt_ref=milestone_pt_ref,
            )
            build_progress_report(
                wb, _sn(f"{str(pos).zfill(3)}进展报告{_nd_str(nd)}（{milestone_label}）"),
                project, phase, nd, rn_pm, "progress_milestone",
                project_start, nd,
                wl_dev, sched_dev, wl_ex, sc_ex,
                lambda s, _rt="progress_milestone": nar(_rt, s),
                metric_refs={} if is_phase1 else _progress_metric_refs(pt_ref_node),
                pt_ref=pt_ref_node,
            )

        else:
            rn_mp = format_report_number("measure_phase",     project["code"], pos)
            rn_pp = format_report_number("progress_phase",     project["code"], pos)

            phase_measure_sheet = _sn(f"{str(pos).zfill(3)}测量与分析报告{_nd_str(nd)}（阶段）")
            phase_measure_metrics = build_measure_report(
                wb, phase_measure_sheet,
                project, phase, nd, rn_mp, "measure_phase",
                period_start, period_end,
                nd_config, nd_modules,
                plan_hours_total, actual_complete, plan_should, actual_complete,
                0, 0,
                lambda s, _rt="measure_phase": nar(_rt, s),
                pt_ref=pt_ref_node,
            )
            build_progress_report(
                wb, _sn(f"{str(pos).zfill(3)}进展报告{_nd_str(nd)}（阶段）"),
                project, phase, nd, rn_pp, "progress_phase",
                period_start, period_end,
                wl_dev, sched_dev, wl_ex, sc_ex,
                lambda s, _rt="progress_phase": nar(_rt, s),
                metric_refs=_progress_metric_refs(pt_ref_node),
                pt_ref=pt_ref_node,
            )
            rn_mtg = format_meeting_number(project["code"], meeting_seq)
            regular_defaults = meeting_defaults(
                "meeting_regular", project, phase, node, tasks,
            )
            def _nar_mtg2(section):
                existing = narratives.get((nid, "meeting", section), "")
                return existing or regular_defaults.get(section, "")
            build_meeting_sheet(
                wb, f"LH{str(meeting_seq).zfill(3)}项目例会单",
                project, phase, meeting_seq, rn_mtg, _nar_mtg2,
            )
            meeting_seq += 1

    # ── 序列化 ───────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
