from datetime import date

from app.services.excel_builder.builder import (
    _meeting_progress_agenda,
    _meeting_sequence_start,
    _task_progress_summary,
)


def _task(wbs, name, start, end, *, leaf=False):
    return {
        "wbs": wbs,
        "name": name,
        "start_date": start,
        "end_date": end,
        "resources": [{"name": "测试人员"}] if leaf else [],
    }


def test_meeting_sequence_accounts_for_phase1_and_phase2_extra_meetings():
    node_phase_wbs = ["1", "2", "2", "3", "4", "4"]

    assert _meeting_sequence_start("1", node_phase_wbs) == 1
    assert _meeting_sequence_start("2", node_phase_wbs) == 4
    assert _meeting_sequence_start("3", node_phase_wbs) == 7
    assert _meeting_sequence_start("4", node_phase_wbs) == 8


def test_task_progress_summary_only_uses_selected_activity_leaf_tasks():
    tasks = [
        _task("2.2", "软件工程活动", date(2026, 1, 1), date(2026, 1, 10)),
        _task("2.2.1", "编写需求", date(2026, 1, 1), date(2026, 1, 5), leaf=True),
        _task("2.3", "评审活动", date(2026, 1, 6), date(2026, 1, 12)),
        _task("2.3.1", "评审需求", date(2026, 1, 6), date(2026, 1, 12), leaf=True),
        _task("2.5", "配置管理活动", date(2026, 1, 1), date(2026, 1, 3)),
        _task("2.5.1", "需求入库", date(2026, 1, 2), date(2026, 1, 3), leaf=True),
    ]

    completed, in_progress = _task_progress_summary(tasks, date(2026, 1, 9))

    assert completed == ["编写需求"]
    assert in_progress == ["评审需求"]


def test_meeting_agenda_omits_in_progress_clause_when_empty():
    tasks = [
        _task("3.2", "软件工程活动", date(2026, 2, 1), date(2026, 2, 5)),
        _task("3.2.1", "编写设计", date(2026, 2, 1), date(2026, 2, 5), leaf=True),
        _task("3.3", "评审活动", date(2026, 2, 6), date(2026, 2, 8)),
        _task("3.3.1", "评审设计", date(2026, 2, 6), date(2026, 2, 8), leaf=True),
    ]

    first = _meeting_progress_agenda(tasks, date(2026, 2, 9), ["2、固定议题"])[0]

    assert "已完成工作：编写设计、评审设计。" in first
    assert "正在进行的工作" not in first


def test_meeting_agenda_keeps_in_progress_clause_when_present():
    tasks = [
        _task("3.2", "软件工程活动", date(2026, 2, 1), date(2026, 2, 10)),
        _task("3.2.1", "编写设计", date(2026, 2, 1), date(2026, 2, 5), leaf=True),
        _task("3.2.2", "编写接口", date(2026, 2, 6), date(2026, 2, 10), leaf=True),
        _task("3.3", "评审活动", date(2026, 2, 11), date(2026, 2, 12)),
    ]

    first = _meeting_progress_agenda(tasks, date(2026, 2, 8), [])[0]

    assert "已完成工作：编写设计；正在进行的工作：编写接口。" in first
