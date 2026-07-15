"""
工时分配引擎：将任务计划工时按确定性随机分布摊铺到工作日。

两个独立概念：
  - 实际已完成工作量（actual_complete）：任务进度指标，由三态逻辑生成
  - 实际工时（WorkReport 日工时）：人员实际花费时间，由 distribute_task_hours 生成

两者单位相同（人时）但含义不同，近似关系：|工时 - 工作量| ≤ 20%。

确定性随机种子 = hash(wbs + person_name)，保证同一任务+人员组合
在不同测量节点间生成完全一致的逐日工时序列。
"""
import random
from datetime import date, timedelta
from dataclasses import dataclass

from app.services.cn_holidays import is_workday


def working_days_in_range(start: date, end: date) -> list[date]:
    """返回 [start, end] 内的所有工作日（含两端）。"""
    days: list[date] = []
    d = start
    while d <= end:
        if is_workday(d):
            days.append(d)
        d += timedelta(days=1)
    return days


def distribute_task_hours(
    wbs: str,
    person_name: str,
    plan_hours: float,
    task_start: date,
    task_end: date,
    exact: bool = False,
) -> dict[date, float]:
    """
    将 plan_hours 按确定性随机权重分配到 [task_start, task_end] 内的工作日。
    用于生成 WorkReport 日工时列（实际工时，非实际已完成工作量）。

    exact=True（阶段1）：总和严格等于 plan_hours，精度 0.1h。
    exact=False：四舍五入到 0.5h，总和在 plan_hours ±20% 范围内。

    返回 {日期: 工时}，工时为 0 的日期不包含在内。
    """
    if plan_hours <= 0:
        return {}

    work_days = working_days_in_range(task_start, task_end)
    if not work_days:
        return {}

    rng = random.Random(hash(wbs + person_name))
    n = len(work_days)

    # 随机权重：稍微前置加权，模拟实际工作节奏
    weights = [rng.uniform(0.4, 1.0) for _ in range(n)]
    total_w = sum(weights)
    raw = [plan_hours * w / total_w for w in weights]

    if exact:
        rounded = [round(h, 1) for h in raw]
        diff = round(plan_hours - sum(rounded), 1)
        if rounded:
            rounded[-1] = round(rounded[-1] + diff, 1)
        return {d: h for d, h in zip(work_days, rounded) if h > 0}
    else:
        if plan_hours >= 1:
            snapped = [round(h) for h in raw]
        else:
            snapped = [round(h * 2) / 2 for h in raw]
        return {d: h for d, h in zip(work_days, snapped) if h > 0}


def compute_plan_should(
    plan_hours: float,
    task_start: date,
    task_end: date,
    node_date: date,
) -> float:
    """
    计划应完成工作量：按任务计划工作日线性进度计算。

    = plan_hours × (节点日期前的计划工作日数 / 任务总计划工作日数)
    """
    all_task_days = working_days_in_range(task_start, task_end)
    total_wd = len(all_task_days)
    if total_wd == 0:
        return 0.0
    days_done = sum(1 for d in all_task_days if d <= node_date)
    return round(plan_hours * days_done / total_wd, 2)


def compute_actual_complete(
    plan_hours: float,
    plan_should: float,
    task_start: date,
    task_end: date,
    node_date: date,
    is_phase1: bool = False,
) -> tuple[float, float]:
    """
    三态逻辑计算实际已完成工作量和完成百分比。

    Returns: (actual_complete, pct_complete)

    三种状态：
    - 已完成（end_date <= node_date）：actual = plan_hours，pct = 1.0
    - 未开始（start_date > node_date）：actual = 0，pct = 0.0
    - 进行中：取最近 5% 倍数，满足 |actual - plan_should| ≤ plan_should × 20%
    """
    if plan_hours <= 0:
        return 0.0, 0.0

    if task_end <= node_date:
        return round(plan_hours, 2), 1.0

    if task_start > node_date:
        return 0.0, 0.0

    # 进行中
    if plan_should <= 0:
        return 0.0, 0.0

    if is_phase1:
        pct = round(plan_should / plan_hours, 4)
        return round(plan_should, 2), pct

    raw_pct = plan_should / plan_hours

    # 取最近 5% 倍数，限制在 [0.05, 0.95]
    pct_snapped = round(raw_pct / 0.05) * 0.05
    pct_snapped = max(0.05, min(0.95, pct_snapped))

    # 候选值：pct_snapped ± 0.05（去重，保留合法范围）
    candidates = sorted({
        max(0.05, pct_snapped - 0.05),
        pct_snapped,
        min(0.95, pct_snapped + 0.05),
    })

    # 优先选满足 ±20% 约束的，多个则选最接近 raw_pct 的
    threshold = plan_should * 0.20
    valid = [c for c in candidates if abs(c * plan_hours - plan_should) <= threshold]
    chosen = min(valid if valid else candidates, key=lambda c: abs(c - raw_pct))

    actual = round(chosen * plan_hours, 2)
    return actual, round(chosen, 4)


@dataclass
class NodeMetrics:
    plan_should_complete: float
    actual_complete: float
    pct_complete: float


def compute_node_metrics(
    plan_hours: float,
    task_start: date,
    task_end: date,
    node_date: date,
    is_phase1: bool = False,
) -> NodeMetrics:
    """
    计算单个（任务, 节点）的三项指标，使用三态逻辑（正确生成顺序）。
    """
    plan_should = compute_plan_should(plan_hours, task_start, task_end, node_date)
    actual, pct = compute_actual_complete(
        plan_hours, plan_should, task_start, task_end, node_date, is_phase1=is_phase1,
    )
    return NodeMetrics(
        plan_should_complete=plan_should,
        actual_complete=actual,
        pct_complete=pct,
    )
