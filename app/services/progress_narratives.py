"""进展报告自动叙述的共享规则。"""


PHASE1_PROGRESS_SECTIONS = (
    "progress_narrative",
    "current_summary",
    "next_plan",
    "author",
)


def phase1_documents(config_items: list[dict]) -> list[dict]:
    """返回阶段1进展报告使用的有效文档配置项。"""
    return [
        item for item in config_items
        if item.get("item_type") == "文档"
    ]


def phase1_progress_defaults(
    project: dict,
    config_items: list[dict],
) -> dict[str, str]:
    """生成阶段1默认叙述；数据库中的非空内容可覆盖这些默认值。"""
    document_names = [
        str(item.get("name") or "").strip()
        for item in phase1_documents(config_items)
        if str(item.get("name") or "").strip()
    ]
    if document_names:
        progress_narrative = (
            "本阶段对用户需求进行分析和确认，生成的工作产品包括："
            f"{'、'.join(document_names)}，上述工作产品已评审后入库。"
        )
    else:
        progress_narrative = "本阶段对用户需求进行分析和确认，工作产品待补充。"

    return {
        "progress_narrative": progress_narrative,
        "current_summary": "本进展期内所有任务正常推进。",
        "next_plan": (
            "进入软件需求分析阶段，制定软件开发计划，"
            "完成软件需求分析和合格性测试需求分析。"
        ),
        "author": str(project.get("manager") or ""),
    }


def phase1_document_scale(config_items: list[dict]) -> int | float:
    """阶段1已生成文档规模，缺失实际页数的文档按0页计。"""
    return sum(
        item.get("actual_pages") or 0
        for item in phase1_documents(config_items)
    )
