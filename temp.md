# MeasureM 当前分析临时记录

更新时间：2026-07-13

## 当前任务

修改“测量与分析报告”的“测量项分析”区，使其根据 `ProgressTracking-ok` 中对应跟踪节点列组的数据自动填写任务/工作产品名称。

当前仅完成阅读和分析，尚未修改业务代码。

## 已阅读内容

- `~/.gstack/projects/MeasureM/gaoyiman-MeasureM-design-20260703-134845.md`
- `app/services/excel_builder/builder.py`
- `app/services/excel_builder/measure_report.py`
- `app/services/excel_builder/progress_tracking.py`
- `app/services/excel_builder/workreport.py`
- `app/services/importer.py`
- `app/services/report_templates.py`

## 实际数据生成链路

```text
importer.py
  → task_node_data
  → builder.py 读取节点数据
  → progress_tracking.py 生成 ProgressTracking-ok
  → measure_report.py 生成测量与分析报告
```

## ProgressTracking 中实际偏差逻辑

双周节点列组为：

```text
计划应完成工作量 | 完成百分比 | 实际已完成工作量 | 完成工作量偏差
```

其中：

```text
完成工作量偏差
= (实际已完成工作量 - 计划应完成工作量) / 计划应完成工作量
```

阶段节点列组为：

```text
计划应完成工作量 | 实际完成百分比 | 实际已完成工作量 | 实际工时 | 进度偏差（阶段统计） | 工作量偏差（阶段统计）
```

其中：

```text
进度偏差（阶段统计）
= (实际完成时间 - 计划完成时间) / (计划完成时间 - 计划开始时间 + 1)

工作量偏差（阶段统计）
= (实际工时 - 实际已完成工作量) / 实际已完成工作量
```

阶段列的实际工时来自最后一个 WorkReport 的人员工时汇总；ProgressTracking 的父节点和阶段汇总行由叶节点聚合。

## 当前报告生成逻辑

`builder.py` 中：

- 双周节点生成 `measure_biweekly`；
- 阶段节点同时生成 `measure_phase` 和 `measure_milestone`；
- 阶段一生成 `measure_phase1`；
- `measure_phase` 和 `measure_milestone` 都传入当前阶段节点的 ProgressTracking 汇总行引用；
- 里程碑报告另计算了累计实际工作量，但当前 `pt_ref` 仍引用当前阶段节点数据，部分公式并非累计值。

`measure_report.py` 当前“测量项分析”全部读取 narrative：

```python
get_nar("deviation_tasks")
get_nar("workload_analysis")
get_nar("schedule_analysis")
get_nar("scale_analysis")
```

导入时这些 narrative 默认是空字符串，所以报告最终显示“无”。本次需求需要新增自动生成分析内容的逻辑。

## 规模数据实际来源

当前报告中显示的规模数据：

- 文档：当前阶段、截至当前节点已评审的配置项；
- 代码模块：当前阶段、截至当前节点已完成的代码模块；
- 规模偏差率由 `deviation_rate(plan, actual)` 计算：

```text
(实际规模 - 计划规模) / 计划规模
```

里程碑报告目前也只显示当前阶段的文档和代码模块，不会合并前序阶段产品。

## 用户原始需求的当前理解

1. 存在工作量偏差任务
   - 只有阶段测量与分析报告填写；
   - 读取对应阶段跟踪节点列组中的“工作量偏差（阶段统计）”；
   - 只筛选叶节点任务；
   - 偏差不为 0 时填写任务名称；无匹配填“无”。

2. 工作量偏差率超阈值的任务
   - 使用同一个阶段“工作量偏差（阶段统计）”列；
   - 筛选绝对值超过 20% 的叶节点任务；
   - 暂理解为 `abs(value) > 0.2`。

3. 存在进度偏差任务
   - 双周报告读取对应双周列组的“完成工作量偏差”；
   - 阶段报告读取对应阶段列组的“进度偏差（阶段统计）”；
   - 偏差不为 0 时填写叶节点任务名称。

4. 进度偏差率超阈值的任务
   - 使用上述对应列；
   - 筛选绝对值超过 20% 的叶节点任务。

5. 规模偏差率超阈值的工作产品
   - 扫描当前报告中显示的文档和代码模块；
   - 使用“文档规模偏差率”和“代码规模偏差率”；
   - 筛选绝对值超过 20% 的产品名称。

## 待用户确认的问题

1. “只有阶段测量与分析报告填写工作量偏差任务”是否表示：
   - `measure_phase` 填写；
   - `measure_biweekly` 不填写；
   - `measure_milestone` 也不填写；
   - `measure_phase1` 继续遵循阶段一特殊结构。

2. 进度偏差分析是否始终只读取当前报告对应的节点列组，而不是多个节点的并集？

3. 阈值是否严格使用：

```python
abs(value) > 0.2
```

4. 多个匹配项是否只填写名称并用换行分隔；没有匹配时填写“无”，不附带偏差率数值？

## 重启后继续工作

1. 先读取本文件和 `AGENTS.md`、`CLAUDE.md`；
2. 如果用户已确认上述边界，使用 gstack 对应 skill（当前任务属于需求/工程规则确认，优先 `/office-hours` 或 `/plan-eng-review`）；
3. 再检查是否需要用户确认的代码实现位置；
4. 编码前不要修改与本需求无关的逻辑；
5. 编码后生成 Excel，检查测量项分析区和 ProgressTracking 对应列组。

