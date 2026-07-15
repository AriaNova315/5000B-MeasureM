"""
静态报告配置：编号格式、节区中文标签、报告结构定义。
excel_builder 的各子模块从这里取配置，不在构建代码中散落硬编码字符串。
"""

# ── 报告编号格式 ─────────────────────────────────────────────────────────────


def format_report_number(report_type: str, project_code: str, position: int) -> str:
    """
    生成标准报告编号。
    position = 该节点在全项目所有追踪节点按日期排序后的序号（从 1 起，含双周+阶段）。
    里程碑版与阶段版共享同一 position（同一跟踪节点）。
    """
    pos = str(position).zfill(3)
    base = f"ZDPX-{project_code}/0.0"
    if "measure" in report_type:
        return f"{base}-MA-测分报告-{pos}"
    if "progress" in report_type:
        return f"{base}-PMC-进展报告-{pos}"
    if report_type == "meeting":
        return f"{base}-PMC-工作例会-{pos}"
    return f"{base}-{pos}"


def format_meeting_number(project_code: str, global_meeting_seq: int) -> str:
    """例会单编号（全项目累计例会序号）。"""
    pos = str(global_meeting_seq).zfill(3)
    return f"ZDPX-{project_code}/0.0-PMC-工作例会-{pos}"


# ── Narrative 节区中文标签 ────────────────────────────────────────────────────

SECTION_LABELS: dict[str, dict[str, str]] = {
    "measure_biweekly": {
        "deviation_tasks":   "偏差任务描述",
        "scale_analysis":    "规模偏差原因分析",
        "workload_analysis": "工作量偏差原因分析",
        "schedule_analysis": "进度偏差原因分析",
        "req_analysis":      "需求偏差原因分析",
        "author":            "拟制人员",
    },
    "measure_phase": {
        "deviation_tasks":   "偏差任务描述",
        "scale_analysis":    "规模偏差原因分析",
        "workload_analysis": "工作量偏差原因分析",
        "schedule_analysis": "进度偏差原因分析",
        "req_analysis":      "需求偏差原因分析",
        "author":            "拟制人员",
    },
    "measure_milestone": {
        "deviation_tasks":   "偏差任务描述",
        "scale_analysis":    "规模偏差原因分析",
        "workload_analysis": "工作量偏差原因分析",
        "schedule_analysis": "进度偏差原因分析",
        "req_analysis":      "需求偏差原因分析",
        "author":            "拟制人员",
    },
    "measure_phase1": {
        "deviation_tasks":   "偏差任务描述",
        "workload_analysis": "工作量偏差原因分析",
        "author":            "拟制人员",
    },
    "progress_biweekly": {
        "change_summary":     "变更情况",
        "progress_narrative": "进展叙述",
        "deviation_stats":    "偏差统计",
        "overall":            "综合情况",
        "risk_issues":        "风险与问题说明",
        "stakeholder":        "利益相关方参与计划跟踪情况",
        "resources_plan":     "资源计划跟踪情况",
        "training":           "人员培训计划跟踪情况",
        "issue_tracking":     "项目问题跟踪情况",
        "current_summary":    "本期工作总结",
        "next_plan":          "下期工作安排",
        "author":             "拟制人员",
    },
    "progress_phase": {
        "change_summary":     "变更情况",
        "progress_narrative": "进展叙述",
        "deviation_stats":    "偏差统计",
        "overall":            "综合情况",
        "risk_issues":        "风险与问题说明",
        "stakeholder":        "利益相关方参与计划跟踪情况",
        "resources_plan":     "资源计划跟踪情况",
        "training":           "人员培训计划跟踪情况",
        "issue_tracking":     "项目问题跟踪情况",
        "current_summary":    "本期工作总结",
        "next_plan":          "下期工作安排",
        "author":             "拟制人员",
    },
    "progress_milestone": {
        "change_summary":     "变更情况",
        "progress_narrative": "进展叙述",
        "deviation_stats":    "偏差统计",
        "overall":            "综合情况",
        "risk_issues":        "风险与问题说明",
        "qa_work":            "质量保证工作",
        "cm_work":            "配置管理工作",
        "issue_tracking":     "项目问题跟踪情况",
        "current_summary":    "本期工作总结",
        "next_plan":          "下期工作安排",
        "author":             "拟制人员",
    },
    "progress_phase1": {
        "progress_narrative": "进展叙述",
        "risk_issues":        "风险与问题说明",
        "current_summary":    "本期工作总结",
        "next_plan":          "下期工作安排",
        "author":             "拟制人员",
    },
    "meeting": {
        "time":      "会议时间",
        "location":  "会议地点",
        "host":      "主持人",
        "attendees": "参加人员",
        "agenda":    "主要内容",
    },
}


def get_section_label(report_type: str, section: str) -> str:
    """返回节区中文标签；阶段1多会议键（如 lh001_time）自动剥离前缀。"""
    # 阶段1 meeting 节键格式：lh001_time → 查 "time"
    plain = section.split("_", 1)[-1] if section.startswith("lh0") else section
    labels = SECTION_LABELS.get(report_type, {})
    return labels.get(plain, plain)


# ── 报告结构定义 ─────────────────────────────────────────────────────────────

class ReportSpec:
    """描述一种报告类型的结构特征，供 excel_builder 分支判断。"""
    __slots__ = (
        "title", "code_type",
        "has_scale", "has_workload", "has_progress", "has_req",
        "has_deviation_analysis",
        "has_stakeholder", "has_resources", "has_training",
        "has_qa", "has_cm",
        "cumulative",
    )

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k, False))


REPORT_SPECS: dict[str, ReportSpec] = {
    "measure_biweekly": ReportSpec(
        title="测量与分析报告", code_type="MA",
        has_scale=True, has_workload=True, has_progress=True,
        has_req=True, has_deviation_analysis=True,
    ),
    "measure_phase": ReportSpec(
        title="测量与分析报告（阶段）", code_type="MA",
        has_scale=True, has_workload=True, has_progress=True,
        has_req=True, has_deviation_analysis=True,
    ),
    "measure_milestone": ReportSpec(
        title="测量与分析报告（里程碑）", code_type="MA",
        has_scale=True, has_workload=True, has_progress=True,
        has_req=True, has_deviation_analysis=True, cumulative=True,
    ),
    "measure_phase1": ReportSpec(
        title="测量与分析报告", code_type="MA",
        has_scale=False, has_workload=True, has_progress=False,
        has_req=False, has_deviation_analysis=True,
    ),
    "progress_biweekly": ReportSpec(
        title="项目进展报告", code_type="PMC",
        has_stakeholder=True, has_resources=True, has_training=True,
    ),
    "progress_phase": ReportSpec(
        title="项目进展报告（阶段）", code_type="PMC",
        has_stakeholder=True, has_resources=True, has_training=True,
    ),
    "progress_milestone": ReportSpec(
        title="项目进展报告（里程碑）", code_type="PMC",
        has_qa=True, has_cm=True, cumulative=True,
    ),
    "progress_phase1": ReportSpec(
        title="项目进展报告", code_type="PMC",
    ),
    "meeting": ReportSpec(
        title="项目例会单", code_type="PMC",
    ),
}


def get_report_spec(report_type: str) -> ReportSpec:
    return REPORT_SPECS.get(report_type, ReportSpec(title=report_type, code_type=""))


# ── 阶段名称标准化 ───────────────────────────────────────────────────────────

def normalize_phase_name(raw_name: str) -> str:
    """去掉常见冗余后缀，保留核心阶段名。"""
    for suffix in ("阶段", "及测试", "与交付"):
        if raw_name.endswith(suffix) and len(raw_name) > len(suffix):
            return raw_name
    return raw_name


# ── 偏差率计算 ───────────────────────────────────────────────────────────────

def deviation_rate(plan: float, actual: float) -> float | None:
    """(actual - plan) / plan，plan=0 时返回 None。"""
    if plan == 0:
        return None
    return round((actual - plan) / plan, 6)
