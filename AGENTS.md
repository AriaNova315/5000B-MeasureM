# MeasureM 开发指南

本文件是 Codex 在本项目中的长期开发约定。为兼容 Claude Code，现有的 `CLAUDE.md` 保留不删除；开始具体任务前应同时参考本文件和 `CLAUDE.md`，若两者冲突，以较新的、明确的项目状态和用户要求为准。

## 项目概述

MeasureM 是项目测量与监控记录生成系统：根据项目实际数据生成逻辑自洽的测量、进度和例会 Excel 记录。

技术栈：

- Python
- FastAPI
- SQLite / aiosqlite
- Alpine.js
- openpyxl

当前核心功能已完成。后续修改应尽量保持现有行为和生成文件格式兼容，避免没有必要的大范围重构。

## 启动与验证

启动服务：

```bash
python3 run.py
```

服务端口固定为 `8002`；不要擅自改为 8000 或 8001，因为这些端口留给同族系统使用。

完成代码修改后，根据改动范围进行验证：

```bash
python3 -m pytest -q
```

如果没有覆盖对应行为的测试，应至少执行模块导入检查、服务启动检查，或实际生成 Excel 并检查输出文件。

## 关键文件和权威资料

- `run.py`：服务启动入口
- `app/main.py`：FastAPI 应用和路由注册
- `app/database.py`：SQLite 数据库连接与初始化
- `app/services/importer.py`：项目 Excel 数据导入
- `app/services/workhour_engine.py`：工作日和工时分配算法
- `app/services/excel_builder/`：各类 Excel Sheet 生成器
- `app/static/index.html`：前端页面
- `02.需求分析阶段.xlsx`：生成格式的权威样例
- `项目相关信息表.xlsx`：输入字段来源样例
- `~/.gstack/projects/MeasureM/gaoyiman-MeasureM-design-20260703-134845.md`：架构和业务设计文档；涉及整体架构、数据库结构或复杂报告规则时优先查阅
- `DEV_PLAN.md`：开发计划，仅作历史参考；如与实际代码或 `CLAUDE.md` 冲突，不要直接按其中的未完成清单判断项目状态

## Excel 兼容性规则

- 修改 Excel 生成逻辑前，先检查权威样例文件和现有 builder 实现。
- 保持 Sheet 名称、报告编号、列顺序、合并单元格、字体、公式和日期范围兼容。
- 日期应使用 `datetime.date` 或 `datetime.datetime` 写入，不要手写 Excel 序列号。
- 不要为了简化实现而删除已有的阶段、双周、里程碑或阶段一特殊逻辑。
- 生成文件后，优先用 openpyxl 重新读取，检查 Sheet 名称、尺寸、关键单元格和公式。

## 数据与业务边界

- MeasureM 不读取 ConfigM 或 EstimateM 的数据库。
- MeasureM 从 Excel 自行解析所需的项目和进度数据。
- 工时分布和实际已完成工作量是两个独立概念，不要混用。
- 工时算法需要保持确定性；不要随意替换现有随机种子或三态完成逻辑。
- 非项目组成员不能写入 WorkReport。
- 偏差率按现有样例保留小数，不要未经确认改成百分比数值。

## 修改流程

1. 先阅读相关代码、`CLAUDE.md` 和样例文件，再提出实现方案。
2. 修改范围保持最小，只触碰完成任务所需文件。
3. 不覆盖或回退用户已有的未提交修改。
4. 完成后运行与改动匹配的验证。
5. 在交付说明中列出修改文件、验证命令和仍存在的不确定性。
6. 如果业务规则、样例格式或现有文档互相矛盾，暂停相关修改并向用户确认。

## 任务路由

当用户请求匹配某个可用 skill 时，必须通过 Skill 工具调用对应 skill；不确定是否匹配时，优先调用 skill。

关键路由规则：

- 产品想法、需求讨论、头脑风暴 → 调用 `/office-hours`
- 产品策略、范围和优先级 → 调用 `/plan-ceo-review`
- 架构、数据库或跨模块设计 → 调用 `/plan-eng-review`
- 设计系统或设计方案评审 → 调用 `/design-consultation` 或 `/plan-design-review`
- 完整需求评审流程 → 调用 `/autoplan`
- Bug、错误和异常定位 → 调用 `/investigate`
- QA、网站行为或交互测试 → 调用 `/qa` 或 `/qa-only`
- 代码审查、diff 检查 → 调用 `/review`
- 视觉质量和界面打磨 → 调用 `/design-review`
- 发布、部署和 PR → 调用 `/ship` 或 `/land-and-deploy`
- 保存工作进度 → 调用 `/context-save`
- 恢复历史上下文 → 调用 `/context-restore`

Excel 格式或生成结果问题仍需先对比权威样例和实际输出；涉及真实页面交互时，除调用 `/qa` 或 `/qa-only` 外，还要检查 API、数据模型和页面状态流转。

### MeasureM 默认 gstack 流程

对于新的、非纯机械的小型或中型需求，默认遵循：

```text
/office-hours → /plan-eng-review → 实现 → /review → 测试与实际验证
```

如果需求涉及明显的产品范围取舍，先加入 `/plan-ceo-review`；如果涉及多个模块、较大重构或完整功能链路，优先使用 `/autoplan`。纯文字、单行修复或用户明确要求的简单操作可以跳过规划 skill，但仍应进行必要验证。

## 工具兼容性

- 本文件供 Codex 使用，`CLAUDE.md` 继续供 Claude Code 使用；两者不要互相删除或重命名。
- 本机已安装 Codex 专用 gstack，skills 位于 `~/.codex/skills/`，由 `~/gstack` 的官方仓库生成。
- Claude Code 的 gstack 位于 `~/.claude/skills/gstack/`；两套 skill 安装位置独立，可以并存。
- 当前已确认可用的核心 gstack skill 包括：`gstack-office-hours`、`gstack-plan-ceo-review`、`gstack-plan-eng-review`、`gstack-autoplan`、`gstack-review`、`gstack-investigate`、`gstack-qa`、`gstack-qa-only`、`gstack-context-save` 和 `gstack-context-restore`。
- gstack 通过 Codex skill 集成使用，不依赖独立的 `gstack` shell 命令。安装或升级后需要重启 Codex，才能让新 skills 在会话中被发现。
- gstack 的项目状态、设计文档和决策记录继续使用 `~/.gstack/projects/MeasureM/`；不要把这些状态文件复制进项目仓库。
- Claude Code 的 `.claude/settings.local.json` 是本地权限配置，不迁移为项目业务规则。
- 除非确实需要配置 MCP、hooks、沙箱或模型默认值，否则不要新增 `.codex/config.toml`。
