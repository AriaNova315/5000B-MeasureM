"""例会实例和自动叙述内容的统一规则。"""

from datetime import date


MEETING_SECTIONS = ("time", "location", "host", "attendees", "agenda")

MEETING_002_AGENDA = [
    "1、讨论并明确技术要求。项目还未正式签订合同，根据已定稿的作为合同附件的技术要求初稿为基础，讨论并明确甲方功能需求、技术指标及接口等要求。",
    "2、按甲方技术要求确定软件关键性等级为C级。",
    "3、讨论软件测试要求。软件测试拟实施单元测试、合格性测试（与出厂所测试合并）。",
    "4、讨论用户验收要求。用户将依据正式签订的技术协议书进行验收。",
    "5、讨论进度要求和关键节点设置。",
]

MEETING_003_TAIL = [
    "2、项目组长通报本周期测量与分析结果，传阅测量与分析报告、进展报告。",
    "3、质量保证人员通报本周期内工作情况。",
    "4、配置管理人员通报本周期内配置活动。",
    "5、问题讨论与决策：无。",
]

MEETING_004_AGENDA = [
    "1、讨论并确定项目生命周期模型和软件过程。经过讨论，并经EPG组批准，本软件开发项目拟采用改进的瀑布模型。该模型将标准瀑布模型的软件设计阶段和编码与单元测试阶段合并，并将合格性测试和出厂检验阶段合并。改进后的瀑布模型过程阶段包括：系统分析与设计、软件需求分析、软件设计与编码测试阶段、软件和合格性测试与验收阶段。",
    "2、讨论软件规模估算方法。确定采用专家估计法，由公司相关人员结合前期工作经验，对文档、代码、测试等工作量进行估计。",
    "3、讨论工作分解结构的构造方法。确定采用基于产品和活动的混合方法。",
    "4、讨论并确认需要纳入数据管理的文档、记录。",
    "5、讨论并确认需要作为配置项管理的文档、记录和源代码。",
    "6、讨论并识别利益相关方及其参与的活动。",
    "7、讨论并识别项目风险。",
    "8、讨论本项目的测量目标、要使用的测量项和测量分析频度。",
]

MEETING_REGULAR_TAIL = [
    "2、项目组长通报本周期测量与分析结果，传阅测量与分析报告、进展报告。",
    "3、项目组成员报告个人工作情况。",
    "4、质量保证人员通报本周期内工作情况。",
    "5、配置管理人员通报本周期内配置活动。",
    "6、问题讨论与决策：无。",
]

MEETING_PROGRESS_PARENT_NAMES = {"软件工程活动", "评审活动"}


def meeting_date_text(value: date | None) -> str:
    return value.strftime("%Y/%m/%d") if value else ""


def meeting_attendees(project: dict) -> str:
    members = str(project.get("team_members") or "")
    return ",".join(name.strip() for name in members.split(",") if name.strip())


def task_resource_names(
    tasks: list[dict],
    task_name: str,
    task_date: date | None = None,
) -> str:
    """按任务活动（及可选节点日期）汇总资源名称，保持原顺序并去重。"""
    matching_tasks = [
        task for task in tasks
        if str(task.get("name") or "").strip() == task_name
        and (
            task_date is None
            or task.get("start_date") == task_date
            or task.get("end_date") == task_date
        )
    ]
    names: list[str] = []
    seen: set[str] = set()
    for task in matching_tasks:
        for resource in task.get("resources") or []:
            name = str(resource.get("name") or "").strip()
            if name and name not in seen:
                seen.add(name)
                names.append(name)
    return ",".join(names)


def meeting_001_agenda(project: dict) -> list[str]:
    start, end = project.get("proj_start"), project.get("proj_end")
    if start and end:
        total_days = (end - start).days + 1
        duration_months = max(1, round(total_days / 30))
    else:
        duration_months = 1
    return [
        f"1、确定软件研制任务。项目正式启动，项目属于软件研制任务，将由项目管理办公室以软件研制任务书的形式下达给项目组，并纳入公司5000体系管理。项目周期为{duration_months}个月左右。",
        "2、讨论并明确项目组关键岗位人选。经讨论，确定以下关键岗位人选：项目总监、项目负责人、质量保证组组长、测试主管，软件设计开发人员暂定。",
        "3、协调项目资源，成立项目策划组。",
    ]


def task_progress_summary(tasks: list[dict], as_of: date) -> tuple[list[str], list[str]]:
    parent_wbs = {
        str(task.get("wbs") or "").strip()
        for task in tasks
        if str(task.get("name") or "").strip() in MEETING_PROGRESS_PARENT_NAMES
    }
    completed: list[str] = []
    in_progress: list[str] = []
    for task in tasks:
        if not task.get("resources"):
            continue
        task_wbs = str(task.get("wbs") or "").strip()
        if not any(task_wbs.startswith(f"{parent}.") for parent in parent_wbs):
            continue
        start, end = task.get("start_date"), task.get("end_date")
        if not start or not end:
            continue
        if end <= as_of:
            completed.append(task["name"])
        elif start <= as_of:
            in_progress.append(task["name"])
    return completed, in_progress


def meeting_progress_agenda(tasks: list[dict], as_of: date, tail: list[str]) -> list[str]:
    completed, in_progress = task_progress_summary(tasks, as_of)
    completed_text = "、".join(completed) if completed else "无"
    first = (
        f"1、项目组长通报本周期项目组完成工作情况：截至{meeting_date_text(as_of)}，"
        f"已完成工作：{completed_text}"
    )
    if in_progress:
        first += f"；正在进行的工作：{'、'.join(in_progress)}"
    first += "。"
    return [first, *tail]


def meeting_sequence_start(current_phase_wbs: str, node_phase_wbs: list[str]) -> int:
    current_phase_number = int(current_phase_wbs)
    counts: dict[str, int] = {}
    for phase_wbs in node_phase_wbs:
        counts[phase_wbs] = counts.get(phase_wbs, 0) + 1

    preceding_meetings = 0
    for phase_wbs, node_count in counts.items():
        if int(phase_wbs) >= current_phase_number:
            continue
        if phase_wbs == "1":
            preceding_meetings += 3 if node_count else 0
        else:
            preceding_meetings += node_count
            if phase_wbs == "2":
                preceding_meetings += 1
    return preceding_meetings + 1


def meeting_instance_storage(instance_id: str) -> tuple[str, str]:
    """返回数据库 report_type 和 section 前缀。"""
    prefixes = {
        "meeting_lh001": "lh001_",
        "meeting_lh002": "lh002_",
        "meeting_lh003": "lh003_",
        "meeting_planning": "planning_",
        "meeting_regular": "",
    }
    if instance_id not in prefixes:
        raise ValueError(f"未知例会实例：{instance_id}")
    return "meeting", prefixes[instance_id]


def meeting_defaults(
    instance_id: str,
    project: dict,
    phase: dict,
    node: dict,
    tasks: list[dict],
) -> dict[str, str]:
    task_start = {
        str(task.get("name") or "").strip(): task.get("start_date")
        for task in tasks
        if str(task.get("name") or "").strip()
    }
    defaults = {
        "location": "610会议室",
        "host": project.get("manager", "") or "",
        "attendees": meeting_attendees(project),
    }

    if instance_id == "meeting_lh001":
        defaults.update(
            time=meeting_date_text(task_start.get("项目启动会") or phase.get("start_date")),
            attendees=task_resource_names(tasks, "项目启动会") or defaults["attendees"],
            agenda="\n".join(meeting_001_agenda(project)),
        )
    elif instance_id == "meeting_lh002":
        defaults.update(
            time=meeting_date_text(task_start.get("软件研制任务书研讨会") or phase.get("start_date")),
            attendees=task_resource_names(tasks, "软件研制任务书研讨会") or defaults["attendees"],
            agenda="\n".join(MEETING_002_AGENDA),
        )
    elif instance_id == "meeting_lh003":
        meeting_date = task_start.get("阶段例会") or phase.get("start_date")
        defaults.update(
            time=meeting_date_text(meeting_date),
            attendees=task_resource_names(tasks, "阶段例会") or defaults["attendees"],
            agenda="\n".join(meeting_progress_agenda(tasks, meeting_date, MEETING_003_TAIL)),
        )
    elif instance_id == "meeting_planning":
        defaults.update(
            time=meeting_date_text(task_start.get("项目策划会") or phase.get("start_date")),
            attendees=task_resource_names(tasks, "项目策划会") or defaults["attendees"],
            agenda="\n".join(MEETING_004_AGENDA),
        )
    elif instance_id == "meeting_regular":
        meeting_date = node.get("date")
        meeting_task_name = (
            "双周例会" if node.get("type") == "双周" or node.get("node_type") == "双周"
            else "阶段例会"
        )
        defaults.update(
            time=meeting_date_text(meeting_date),
            attendees=(
                task_resource_names(tasks, meeting_task_name, meeting_date)
                or defaults["attendees"]
            ),
            agenda="\n".join(
                meeting_progress_agenda(tasks, meeting_date, MEETING_REGULAR_TAIL)
            ),
        )
    else:
        raise ValueError(f"未知例会实例：{instance_id}")
    return {section: str(defaults.get(section) or "") for section in MEETING_SECTIONS}
