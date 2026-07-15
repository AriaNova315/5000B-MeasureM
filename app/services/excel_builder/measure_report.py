"""
测量与分析报告 Sheet 构建器。
支持 measure_biweekly / measure_phase / measure_milestone / measure_phase1 四种类型。
行数随文档数量动态变化，约 31 行 × 10 列。
"""
from datetime import date

from app.services.report_templates import deviation_rate, get_report_spec
from .utils import (
    _w, _mc, _wmc, _s, border_range,
    F_TITLE, F_HDR, F_BODY, A_CTR, A_CTR_W, A_LEFT, A_LEFT_W, B_THIN,
    fmt_date, fmt_date_range, set_col_width, set_row_height,
)

_PCT_FMT = "0.00%"
_NUM_FMT = "0.00"
_DEV_FMT = "0.000000"


def build_measure_report(
    wb,
    sheet_name: str,
    project: dict,
    phase: dict,
    node_date: date,
    report_number: str,
    report_type: str,
    period_start: date,
    period_end: date,
    config_items: list,     # [{seq, name, plan_pages, actual_pages, item_type}]
    code_modules: list,     # [{seq, name, plan_loc, actual_loc}]
    plan_hours: float,      # 阶段计划总工时
    actual_hours: float,    # 实际总工时（到节点日期）
    plan_should: float,     # 计划应完成工作量
    actual_complete: float, # 实际已完成工作量
    req_total: int,         # 需求总数（暂由叙述提供）
    req_changes: int,       # 需求变更数
    get_nar,                # callable(section) -> str
    pt_ref: dict | None = None,
    # pt_ref = None → phase1, use numeric values
    # pt_ref = {"is_biweekly": True} → biweekly: workload/req → "——"
    # pt_ref = {"is_biweekly": False, "sheet":..., "plan_wl_col":...,
    #            "actual_hrs_col":..., "plan_should_col":...,
    #            "actual_complete_col":..., "row":...}
    #         → phase/milestone: formula refs to ProgressTracking row 3
):
    spec = get_report_spec(report_type)
    is_phase1 = (report_type == "measure_phase1")
    ws = wb.create_sheet(sheet_name)

    # 列宽
    col_widths = [8, 6, 22, 6, 6, 6, 10, 10, 14, 6]
    for i, w in enumerate(col_widths, 1):
        set_col_width(ws, i, w)

    r = 1  # 当前行游标

    # ── R1: 标题 ──────────────────────────────────────────────────────────────
    _wmc(ws, r, 1, r, 10, "测量与分析报告", font=F_TITLE, alignment=A_CTR)
    set_row_height(ws, r, 36)
    r += 1

    # ── R2: 报告编号 ──────────────────────────────────────────────────────────
    _wmc(ws, r, 5, r, 10, f"编号：{report_number}", font=F_BODY, alignment=A_LEFT)
    r += 1

    # ── R3-R5: 基本信息 ───────────────────────────────────────────────────────
    _wmc(ws, r, 1, r, 2, "项目名称", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 3, r, 5, project["name"], font=F_BODY, alignment=A_LEFT)
    _wmc(ws, r, 6, r, 7, "项目编号", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 8, r, 10, f"ZDPX{project['code']}", font=F_BODY, alignment=A_LEFT)
    r += 1

    phase_start = phase["start_date"]
    phase_end   = phase["end_date"]
    proj_period = fmt_date_range(
        project.get("proj_start") or phase_start,
        project.get("proj_end")   or phase_end,
    )

    _wmc(ws, r, 1, r, 2, "项目周期", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 3, r, 5, proj_period, font=F_BODY, alignment=A_LEFT)
    _wmc(ws, r, 6, r, 7, "项目经理", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 8, r, 10, project.get("manager", ""), font=F_BODY, alignment=A_LEFT)
    r += 1

    # 阶段名（去掉"阶段"后缀以简化）
    phase_short = phase["name"].replace("阶段", "").strip() or phase["name"]
    _wmc(ws, r, 1, r, 2, "所处阶段", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 3, r, 5, phase_short, font=F_BODY, alignment=A_LEFT)
    _wmc(ws, r, 6, r, 7, "报告时间段", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 8, r, 10, fmt_date_range(period_start, period_end), font=F_BODY, alignment=A_LEFT)
    r += 1

    # ── 测量项统计 ────────────────────────────────────────────────────────────
    _wmc(ws, r, 1, r, 10, "测量项统计", font=F_HDR, alignment=A_CTR)
    r += 1

    # ── 规模区 ────────────────────────────────────────────────────────────────
    if spec.has_scale:
        doc_items = [ci for ci in config_items if ci["item_type"] == "文档"]
        # 纵向合并：文档表头 + 文档行 + 代码表头 + 代码行
        scale_span = len(doc_items) + 1 + len(code_modules)
        _wmc(ws, r, 1, r + scale_span, 1, "规模", font=F_HDR, alignment=A_CTR)

        # 文档表头
        _w(ws, r, 2, "序号"); _wmc(ws, r, 3, r, 6, "文档名称", font=F_HDR, alignment=A_CTR)
        _w(ws, r, 7, "计划规模"); _w(ws, r, 8, "实际规模"); _wmc(ws, r, 9, r, 10, "文档规模偏差率")
        _s(ws, r, 2, font=F_HDR, alignment=A_CTR)
        _s(ws, r, 7, font=F_HDR, alignment=A_CTR)
        _s(ws, r, 8, font=F_HDR, alignment=A_CTR)
        _s(ws, r, 9, font=F_HDR, alignment=A_CTR)
        r += 1

        for doc in doc_items:
            plan_p  = doc.get("plan_pages") or 0
            act_p   = doc.get("actual_pages") or 0
            dev     = deviation_rate(plan_p, act_p)
            _w(ws, r, 2, doc["seq"])
            _wmc(ws, r, 3, r, 6, doc["name"], font=F_BODY, alignment=A_LEFT)
            _w(ws, r, 7, plan_p); _w(ws, r, 8, act_p)
            _wmc(ws, r, 9, r, 10, dev if dev is not None else "——")
            if dev is not None:
                _s(ws, r, 9, number_format=_PCT_FMT)
            r += 1

        # 代码模块表头（去掉"（LOC）"）
        _w(ws, r, 2, "序号"); _wmc(ws, r, 3, r, 6, "代码模块名称", font=F_HDR, alignment=A_CTR)
        _w(ws, r, 7, "计划规模"); _w(ws, r, 8, "实际规模")
        _wmc(ws, r, 9, r, 10, "代码规模偏差率")
        _s(ws, r, 2, font=F_HDR, alignment=A_CTR)
        _s(ws, r, 7, font=F_HDR, alignment=A_CTR)
        _s(ws, r, 8, font=F_HDR, alignment=A_CTR)
        _s(ws, r, 9, font=F_HDR, alignment=A_CTR)
        r += 1

        for mod in code_modules:
            plan_l  = mod.get("plan_loc") or 0
            act_l   = mod.get("actual_loc") or 0
            dev     = deviation_rate(plan_l, act_l)
            _w(ws, r, 2, mod["seq"])
            _wmc(ws, r, 3, r, 6, mod["name"], font=F_BODY, alignment=A_LEFT)
            _w(ws, r, 7, plan_l); _w(ws, r, 8, act_l)
            _wmc(ws, r, 9, r, 10, dev if dev is not None else "——")
            if dev is not None:
                _s(ws, r, 9, number_format=_PCT_FMT)
            r += 1

    # ── 工作量区 ──────────────────────────────────────────────────────────────
    r0 = r
    _wmc(ws, r0, 1, r0 + 1, 1, "工作量", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r0, 2, r0, 4, "计划工作量", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r0, 7, r0, 8, "实际工时", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r0 + 1, 2, r0 + 1, 4, "工作量偏差", font=F_HDR, alignment=A_CTR)

    if pt_ref is not None and pt_ref.get("is_biweekly"):
        _wmc(ws, r0, 5, r0, 6, "——", alignment=A_CTR)
        _wmc(ws, r0, 9, r0, 10, "——", alignment=A_CTR)
        _wmc(ws, r0 + 1, 5, r0 + 1, 10, "——", alignment=A_CTR)
    elif pt_ref is not None:
        sn, pr = pt_ref["sheet"], pt_ref["row"]
        pc, ac = pt_ref["plan_wl_col"], pt_ref["actual_hrs_col"]
        _wmc(ws, r0, 5, r0, 6, f"='{sn}'!{pc}{pr}", alignment=A_CTR)
        _wmc(ws, r0, 9, r0, 10, f"='{sn}'!{ac}{pr}", alignment=A_CTR)
        _wmc(ws, r0 + 1, 5, r0 + 1, 10,
             f"=IF('{sn}'!{pc}{pr}=0,\"——\",('{sn}'!{ac}{pr}-'{sn}'!{pc}{pr})/'{sn}'!{pc}{pr})",
             alignment=A_CTR)
        _s(ws, r0 + 1, 5, number_format=_PCT_FMT)
    else:
        wl_dev = deviation_rate(plan_hours, actual_hours)
        _wmc(ws, r0, 5, r0, 6, plan_hours, alignment=A_CTR)
        _wmc(ws, r0, 9, r0, 10, actual_hours, alignment=A_CTR)
        _wmc(ws, r0 + 1, 5, r0 + 1, 10, wl_dev if wl_dev is not None else "——", alignment=A_CTR)
        if wl_dev is not None:
            _s(ws, r0 + 1, 5, number_format=_PCT_FMT)

    r = r0 + 2

    # ── 进度区 ────────────────────────────────────────────────────────────────
    if spec.has_progress:
        r0 = r
        _wmc(ws, r0, 1, r0 + 2, 1, "进度", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r0, 2, r0, 4, "计划应完成工作量", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r0, 7, r0, 8, "实际已完成工作量", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r0 + 1, 2, r0 + 1, 4, "完成工作量偏差", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r0 + 1, 7, r0 + 1, 8, "阶段进度偏差", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r0 + 2, 2, r0 + 2, 4, "各任务进度偏差", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r0 + 2, 5, r0 + 2, 10, "详见进度追踪表（ProgressTracking）", font=F_BODY, alignment=A_LEFT)

        if pt_ref is not None:
            sn, pr = pt_ref["sheet"], pt_ref["row"]
            psc = pt_ref["plan_should_col"]
            acc = pt_ref["actual_complete_col"]
            _wmc(ws, r0, 5, r0, 6, f"='{sn}'!{psc}{pr}", alignment=A_CTR)
            _wmc(ws, r0, 9, r0, 10, f"='{sn}'!{acc}{pr}", alignment=A_CTR)
            _wmc(ws, r0 + 1, 5, r0 + 1, 6,
                 f"=IF('{sn}'!{psc}{pr}=0,\"——\",('{sn}'!{acc}{pr}-'{sn}'!{psc}{pr})/'{sn}'!{psc}{pr})",
                 alignment=A_CTR)
            _s(ws, r0 + 1, 5, number_format=_PCT_FMT)
            _wmc(ws, r0 + 1, 9, r0 + 1, 10,
                 f"=IF('{sn}'!{psc}{pr}=0,\"——\",'{sn}'!{acc}{pr}-'{sn}'!{psc}{pr})",
                 alignment=A_CTR)
        else:
            sched_dev = round(actual_complete - plan_should, 2)
            _wmc(ws, r0, 5, r0, 6, plan_should, alignment=A_CTR)
            _wmc(ws, r0, 9, r0, 10, actual_complete, alignment=A_CTR)
            _wmc(ws, r0 + 1, 5, r0 + 1, 6, sched_dev, alignment=A_CTR)
            _wmc(ws, r0 + 1, 9, r0 + 1, 10, round(actual_complete - plan_should, 2), alignment=A_CTR)

        r = r0 + 3

    # 过程和产品符合性
    _wmc(ws, r, 1, r, 3, "过程和产品符合性", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 10, "测量数据见质量审核工作报告", font=F_BODY, alignment=A_LEFT)
    r += 1

    # ── 需求区 ────────────────────────────────────────────────────────────────
    if spec.has_req:
        _wmc(ws, r, 1, r + 1, 1, "需求", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r, 2, r, 4, "需求总数", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r, 7, r, 8, "需求变更数", font=F_HDR, alignment=A_CTR)
        if pt_ref is not None and pt_ref.get("is_biweekly"):
            _wmc(ws, r, 5, r, 6, "——", alignment=A_CTR)
            _wmc(ws, r, 9, r, 10, "——", alignment=A_CTR)
            r += 1
            _wmc(ws, r, 2, r, 4, "需求变更率", font=F_HDR, alignment=A_CTR)
            _wmc(ws, r, 5, r, 10, "——", alignment=A_CTR)
        else:
            req_rate = deviation_rate(req_total, req_total + req_changes) if req_total else None
            _wmc(ws, r, 5, r, 6, req_total, alignment=A_CTR)
            _wmc(ws, r, 9, r, 10, req_changes, alignment=A_CTR)
            r += 1
            _wmc(ws, r, 2, r, 4, "需求变更率", font=F_HDR, alignment=A_CTR)
            _wmc(ws, r, 5, r, 10, req_rate if req_rate is not None else 0, alignment=A_CTR)
        r += 1

    # ── 测量项分析 ────────────────────────────────────────────────────────────
    _wmc(ws, r, 1, r, 10, "测量项分析", font=F_HDR, alignment=A_CTR)
    r += 1

    analysis_rows = [
        ("存在工作量偏差任务",      get_nar("deviation_tasks") or "无"),
        ("工作量偏差率超阈值的任务", get_nar("workload_analysis") or "无"),
    ]
    if spec.has_progress:
        analysis_rows.extend([
            ("存在进度偏差任务",      get_nar("schedule_analysis") or "无"),
            ("进度偏差率超阈值的任务", "无"),
        ])
    if spec.has_scale:
        analysis_rows.append(("规模偏差率超阈值的工作产品", get_nar("scale_analysis") or "无"))
    analysis_rows.append(("异常状态原因分析", "无"))

    for label, content in analysis_rows:
        _wmc(ws, r, 1, r, 3, label, font=F_HDR, alignment=A_CTR_W)
        _wmc(ws, r, 4, r, 10, content, font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 30)
        r += 1

    # ── 拟制人员 ──────────────────────────────────────────────────────────────
    _wmc(ws, r, 1, r, 3, "拟制人员", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 10, get_nar("author") or project.get("manager", ""), font=F_BODY, alignment=A_LEFT)
    last_row = r

    # ── 边框 ─────────────────────────────────────────────────────────────────
    border_range(ws, 3, 1, last_row, 10)
