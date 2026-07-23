"""
进展报告 Sheet 构建器。
支持 progress_biweekly / progress_phase / progress_milestone / progress_phase1 四种类型。
约 36 行 × 9 列（数据从 B 列起）。
"""
from datetime import date

from app.services.report_templates import get_report_spec
from .utils import (
    _w, _mc, _wmc, _s, border_range,
    F_TITLE, F_HDR, F_BODY, A_CTR, A_LEFT, A_LEFT_W, B_THIN,
    fmt_date, fmt_date_range, set_col_width, set_row_height,
)

# 文档列表：阶段版（结尾文档）
_PHASE_DOCS = [
    "测量与分析报告", "质量审核工作报告", "配置管理报告",
    "项目风险跟踪表", "利益相关方参与计划跟踪表",
    "项目资源计划跟踪表", "人员培训计划跟踪表",
]
_MILESTONE_DOCS = [
    "测量与分析报告", "质量审核工作报告", "配置管理报告",
    "项目风险跟踪表", "利益相关方参与计划跟踪表",
    "项目资源计划跟踪表", "人员培训计划跟踪表",
]


def _count_threshold_formula(pt_ref: dict, deviation_key: str) -> str | None:
    """按叶节点偏差列统计绝对值超过 20% 的任务数量。"""
    if not pt_ref or not pt_ref.get(deviation_key):
        return None
    sheet = pt_ref["sheet"].replace("'", "''")
    start, end = pt_ref["data_start"], pt_ref["data_end"]
    leaf_range = f"'{sheet}'!$E${start}:$E${end}"
    dev_col = pt_ref[deviation_key]
    dev_range = f"'{sheet}'!${dev_col}${start}:${dev_col}${end}"
    return (
        f'=COUNTIFS({leaf_range},"<>",{dev_range},">0.2")+'
        f'COUNTIFS({leaf_range},"<>",{dev_range},"<-0.2")'
    )


def build_progress_report(
    wb,
    sheet_name: str,
    project: dict,
    phase: dict,
    node_date: date,
    report_number: str,
    report_type: str,
    period_start: date,
    period_end: date,
    wl_deviation: float,    # 工作量偏差率
    sched_deviation: float, # 进度偏差
    wl_exceed_tasks: int,   # 工作量偏差>20%任务数
    sched_exceed_tasks: int,# 进度偏差>20%任务数
    get_nar,                # callable(section) -> str
    metric_refs: dict | None = None,
    pt_ref: dict | None = None,
):
    spec = get_report_spec(report_type)
    is_phase1 = (report_type == "progress_phase1")
    is_milestone = (report_type == "progress_milestone")
    ws = wb.create_sheet(sheet_name)

    # 列宽（A 留空，B-I 为内容区）
    col_widths = [2, 16, 14, 8, 8, 8, 14, 8, 16]
    for i, w in enumerate(col_widths, 1):
        set_col_width(ws, i, w)

    r = 1

    # ── R1: 标题 ──────────────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 9, "项目进展报告", font=F_TITLE, alignment=A_CTR)
    set_row_height(ws, r, 36)
    r += 1

    # ── R2: 编号 ──────────────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 9, f"报告顺序号：{report_number}", font=F_BODY, alignment=A_LEFT)
    r += 1

    # ── 项目基本信息 ──────────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 9, "项目基本信息", font=F_HDR, alignment=A_CTR)
    r += 1

    phase_start = phase["start_date"]
    phase_end   = phase["end_date"]
    phase_short = phase["name"].replace("阶段", "").strip() or phase["name"]
    proj_period = fmt_date_range(
        project.get("proj_start") or phase_start,
        project.get("proj_end")   or phase_end,
    )

    _wmc(ws, r, 2, r, 3, "项目名称", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 6, project["name"], font=F_BODY, alignment=A_LEFT)
    _wmc(ws, r, 7, r, 7, "项目编号", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 8, r, 9, f"ZDPX{project['code']}", font=F_BODY, alignment=A_LEFT)
    r += 1

    _wmc(ws, r, 2, r, 3, "项目周期", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 6, proj_period, font=F_BODY, alignment=A_LEFT)
    _wmc(ws, r, 7, r, 7, "项目负责人", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 8, r, 9, project.get("manager", ""), font=F_BODY, alignment=A_LEFT)
    r += 1

    _wmc(ws, r, 2, r, 3, "所处阶段", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 6, phase_short, font=F_BODY, alignment=A_LEFT)
    _wmc(ws, r, 7, r, 7, "报告时间段", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 8, r, 9, fmt_date_range(period_start, period_end), font=F_BODY, alignment=A_LEFT)
    r += 1

    # ── 项目进展情况 ──────────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 9, "项目进展情况", font=F_HDR, alignment=A_CTR)
    r += 1

    _wmc(ws, r, 2, r, 3, "项目变更情况", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 9, get_nar("change_summary") or "——", font=F_BODY, alignment=A_LEFT_W)
    r += 1

    # 进展叙述。左侧标签在本节结束后统一纵向合并，覆盖统计行。
    progress_section_start = r
    _wmc(ws, r, 4, r, 9, "本进展期内任务完成情况：", font=F_BODY, alignment=A_LEFT)
    r += 1
    _wmc(ws, r, 4, r, 9, get_nar("progress_narrative") or "（待填写）",
         font=F_BODY, alignment=A_LEFT_W)
    set_row_height(ws, r, 60)
    r += 1

    if not is_milestone:
        # 偏差统计（非里程碑版）
        _wmc(ws, r, 4, r, 5, "工作量偏差超20%任务数", font=F_BODY, alignment=A_LEFT)
        workload_count = _count_threshold_formula(pt_ref, "workload_dev_col") or wl_exceed_tasks
        _wmc(ws, r, 6, r, 9, workload_count, alignment=A_CTR)
        r += 1
        _wmc(ws, r, 4, r, 5, "进度偏差超20%任务数", font=F_BODY, alignment=A_LEFT)
        progress_count = _count_threshold_formula(pt_ref, "progress_dev_col") or sched_exceed_tasks
        _wmc(ws, r, 6, r, 9, progress_count, alignment=A_CTR)
        r += 1

        # 综合情况
        _wmc(ws, r, 4, r, 9, "本进展期内项目综合情况：", font=F_BODY, alignment=A_LEFT)
        r += 1
        _wmc(ws, r, 4, r, 5, "工作量偏差：", font=F_BODY, alignment=A_LEFT)
        workload_metric = (metric_refs or {}).get("workload_deviation", wl_deviation)
        _wmc(ws, r, 6, r, 9, workload_metric, alignment=A_CTR)
        _s(ws, r, 6, number_format="0.00%")
        r += 1
        _wmc(ws, r, 4, r, 5, "进度偏差", font=F_BODY, alignment=A_LEFT)
        schedule_metric = (metric_refs or {}).get("schedule_deviation", sched_deviation)
        _wmc(ws, r, 6, r, 9, schedule_metric, alignment=A_CTR)
        _s(ws, r, 6, number_format="0.00%")
        r += 1
        _wmc(ws, r, 4, r, 5, "各工作产品规模偏差", font=F_BODY, alignment=A_LEFT)
        _wmc(ws, r, 6, r, 9, "见《测量分析报告》", font=F_BODY, alignment=A_LEFT)
        r += 1
    else:
        # 里程碑版同样展示偏差统计，指标直接引用 ProgressTracking。
        _wmc(ws, r, 4, r, 5, "工作量偏差超20%任务数", font=F_BODY, alignment=A_LEFT)
        workload_count = _count_threshold_formula(pt_ref, "workload_dev_col") or wl_exceed_tasks
        _wmc(ws, r, 6, r, 9, workload_count, alignment=A_CTR)
        r += 1
        _wmc(ws, r, 4, r, 5, "进度偏差超20%任务数", font=F_BODY, alignment=A_LEFT)
        progress_count = _count_threshold_formula(pt_ref, "progress_dev_col") or sched_exceed_tasks
        _wmc(ws, r, 6, r, 9, progress_count, alignment=A_CTR)
        r += 1
        _wmc(ws, r, 4, r, 9, "本进展期内项目综合情况：", font=F_BODY, alignment=A_LEFT)
        r += 1
        _wmc(ws, r, 4, r, 5, "工作量偏差：", font=F_BODY, alignment=A_LEFT)
        workload_metric = (metric_refs or {}).get("workload_deviation", wl_deviation)
        _wmc(ws, r, 6, r, 9, workload_metric, alignment=A_CTR)
        _s(ws, r, 6, number_format="0.00%")
        r += 1
        _wmc(ws, r, 4, r, 5, "进度偏差", font=F_BODY, alignment=A_LEFT)
        schedule_metric = (metric_refs or {}).get("schedule_deviation", sched_deviation)
        _wmc(ws, r, 6, r, 9, schedule_metric, alignment=A_CTR)
        _s(ws, r, 6, number_format="0.00%")
        r += 1
        _wmc(ws, r, 4, r, 5, "各工作产品规模偏差", font=F_BODY, alignment=A_LEFT)
        _wmc(ws, r, 6, r, 9, "见《测量分析报告》", font=F_BODY, alignment=A_LEFT)
        r += 1

    progress_section_end = r - 1
    _wmc(ws, progress_section_start, 2, progress_section_end, 3,
         "项目进展情况", font=F_HDR, alignment=A_CTR)

    # ── 风险问题 ──────────────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 3, "风险问题说明", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 9, get_nar("risk_issues") or "（待填写）", font=F_BODY, alignment=A_LEFT_W)
    set_row_height(ws, r, 40)
    r += 1

    # ── 仅 biweekly / phase 版：利益相关方、资源、培训 ─────────────────────────
    if spec.has_stakeholder:
        _wmc(ws, r, 2, r, 3, "利益相关方参与计划跟踪情况", font=F_HDR, alignment=A_CTR_W)
        _wmc(ws, r, 4, r, 9, get_nar("stakeholder") or "按计划进行", font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 40)
        r += 1

    if spec.has_resources:
        _wmc(ws, r, 2, r, 3, "资源计划\n跟踪情况", font=F_HDR, alignment=A_CTR_W)
        _wmc(ws, r, 4, r, 9, get_nar("resources_plan") or "（待填写）", font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 40)
        r += 1

    if spec.has_training:
        _wmc(ws, r, 2, r, 3, "人员培训计划\n跟踪情况", font=F_HDR, alignment=A_CTR_W)
        _wmc(ws, r, 4, r, 9, get_nar("training") or "（待填写）", font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 40)
        r += 1

    # ── 仅里程碑版：QA 工作、配管工作 ────────────────────────────────────────
    if spec.has_qa:
        _wmc(ws, r, 2, r, 3, "质量保证工作", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r, 4, r, 9, get_nar("qa_work") or "（待填写）", font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 40)
        r += 1

    if spec.has_cm:
        _wmc(ws, r, 2, r, 3, "配置管理工作", font=F_HDR, alignment=A_CTR)
        _wmc(ws, r, 4, r, 9, get_nar("cm_work") or "（待填写）", font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 40)
        r += 1

    # ── 问题跟踪（非阶段1）────────────────────────────────────────────────────
    if not is_phase1:
        _wmc(ws, r, 2, r, 3, "项目问题\n跟踪情况", font=F_HDR, alignment=A_CTR_W)
        _wmc(ws, r, 4, r, 9, get_nar("issue_tracking") or "未发生新的项目问题", font=F_BODY, alignment=A_LEFT_W)
        set_row_height(ws, r, 30)
        r += 1

    # ── 本期总结、下期安排 ────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 3, "本期工作总结", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 9, get_nar("current_summary") or "本阶段工作按计划正常进行", font=F_BODY, alignment=A_LEFT_W)
    set_row_height(ws, r, 30)
    r += 1

    _wmc(ws, r, 2, r, 3, "下期工作安排", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 9, get_nar("next_plan") or "按开发计划实施", font=F_BODY, alignment=A_LEFT_W)
    set_row_height(ws, r, 30)
    r += 1

    # ── 拟制人员、领导意见 ────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 3, "拟制人员", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 9, get_nar("author") or project.get("manager", ""), font=F_BODY, alignment=A_LEFT)
    r += 1

    _wmc(ws, r, 2, r, 3, "领导意见", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 4, r, 9, "", font=F_BODY, alignment=A_LEFT)
    r += 1

    # ── 查看相关文档 ──────────────────────────────────────────────────────────
    _wmc(ws, r, 2, r, 9, "查看相关文档", font=F_HDR, alignment=A_CTR)
    r += 1
    _wmc(ws, r, 2, r, 2, "序号", font=F_HDR, alignment=A_CTR)
    _wmc(ws, r, 3, r, 8, "文档名称", font=F_HDR, alignment=A_CTR)
    _w(ws, r, 9, "上传日期")
    _s(ws, r, 9, font=F_HDR, alignment=A_CTR)
    r += 1

    doc_list = _MILESTONE_DOCS if is_milestone else _PHASE_DOCS
    date_str = fmt_date(node_date)
    for i, doc_name in enumerate(doc_list, 1):
        _w(ws, r, 2, i)
        _wmc(ws, r, 3, r, 8, doc_name, font=F_BODY, alignment=A_LEFT)
        _w(ws, r, 9, date_str)
        r += 1

    last_row = r - 1
    border_range(ws, 3, 2, last_row, 9)


# 解决 A_CTR_W 未定义问题
from .utils import A_CTR_W
