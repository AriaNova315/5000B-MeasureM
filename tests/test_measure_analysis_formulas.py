from app.services.excel_builder.measure_report import (
    _helper_task_formula,
    _helper_scale_formula,
    _helper_accumulate_formula,
    _helper_result_formula,
)


def test_progress_tracking_formula_filters_leaf_tasks_and_threshold():
    pt_ref = {
        "sheet": "ProgressTracking-ok",
        "data_start": 4,
        "data_end": 45,
        "workload_dev_col": "N",
        "progress_dev_col": "M",
    }

    formula = _helper_task_formula(pt_ref, 4, "workload_dev_col", threshold=True)

    assert "'ProgressTracking-ok'!$B$4" in formula
    assert "'ProgressTracking-ok'!$E$4<>\"\"" in formula
    assert "ISNUMBER('ProgressTracking-ok'!$N$4)" in formula
    assert "ABS('ProgressTracking-ok'!$N$4)>0.2" in formula

    first = _helper_accumulate_formula("A", "F", 2, True)
    next_row = _helper_accumulate_formula("A", "F", 3, False)
    result = _helper_result_formula("F", 43)
    assert first == '=IF($A$2="","",$A$2)'
    assert "CHAR(10)" in next_row
    assert "'AnalysisHelper'!$F$43" in result


def test_scale_formula_references_document_and_code_ranges():
    formula = _helper_scale_formula("002测量与分析报告", 12)

    assert "'002测量与分析报告'!$C$12" in formula
    assert "'002测量与分析报告'!$I$12" in formula
    assert "ABS('002测量与分析报告'!$I$12)>0.2" in formula
