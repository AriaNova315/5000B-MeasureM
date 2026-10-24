# 离线交付验证记录

日期：2026-10-24。部署实现已完成，Windows 原生安装与桌面 Excel 尚待验收。

## 已通过

| 环境/范围 | 证据 |
|---|---|
| Mac原有环境 | Python3.14.3；未重装依赖。unittest发现测试类并补跑6个函数测试，共42项通过 |
| 新依赖隔离环境 | `/private/tmp`中新建venv，按requirements.txt安装；`python -B -m pytest -q`结果：42 passed、11 subtests passed；pip check通过 |
| Windows wheel可获得性 | CPython3.14/win_amd64交叉下载18个完整wheels；SHA-256写入lock；随后使用`--no-index --no-cache-dir --find-links`本地读取并校验通过 |
| 空白模板/虚构输入 | 从零生成，不复制真实项目；openpyxl可读，演示项目覆盖阶段1、双周、阶段、里程碑 |
| 干净交付包 | 明确清单打包39个文件并逐项核对MANIFEST哈希；解压至临时中文/空格路径，无数据库、历史导出、venv、wheels或安装程序 |
| 核心流程 | 在临时数据库中首次初始化为空、重复初始化、导入、项目/节点列表、叙述覆盖及恢复、两个阶段导出成功；阶段1共7个Sheet、阶段2共13个Sheet |
| 实际HTTP流程 | 从临时交付副本在8002启动；HTML与本地Alpine资源200；docs/redoc为404；OpenAPI200；multipart上传、叙述PUT、两个Excel下载和恢复成功 |
| 输出结构 | openpyxl重读成功，阶段1 ProgressTracking为5列；存在WorkReport、阶段/里程碑报告与自动重算标志；HTTP中文文件名使用RFC5987编码 |
| 本机监听 | lsof确认仅`127.0.0.1:8002`监听；从项目外目录启动仍使用临时副本根目录 |
| 错误处理 | 测试覆盖非法名称在构建前拒绝、权限/占用409、磁盘错误500，以及启动目录不可写 |
| 代码/资源 | git diff --check、Python AST、本地Alpine哈希、Windows bat CRLF检查通过；手工按/review检查清单核对部署diff |
| 数据保护 | 真实measurem.db及既有generated文件的SHA-256在测试前后保持不变 |

锁定清单对应标准CPython3.14 x64，Windows原生部署人员应使用说明指定的完整安装程序与同一lock重建wheelhouse。交叉下载不等于Windows安装认证，也不提供Windows二进制包在Mac可运行的保证。

依赖选择：直接依赖明确固定版本，移除当前无代码引用的sse-starlette及未使用的uvicorn[standard]扩展；不添加业务功能。FastAPI的传递依赖包含OpenTelemetry API，但本项目未配置遥测提供者/导出器，不新增在线调用。Mac当前已安装环境保持原样。

## 尚未完成与签收限制

- Windows11 x64上的完整Python安装、原生无缓存离线pip安装、start_windows.bat和prepare_windows.ps1实际执行。
- Windows文件占用/权限行为、从中文/空格路径双击启动、局域网另一台机器的不可访问性。Mac已验证回环绑定和对应错误分支，但不是Windows实测。
- 实际桌面Excel版本下的公式重算、缓存结果和打印版式。openpyxl只验证结构，不能代替公式求值。
- 严格禁外网、无缓存浏览器的完整交互链。gstack浏览器曾显示本地主页及Alpine驱动的“开始导入”状态，但后续浏览器守护进程连接/启动锁不稳定；备用计算机浏览器接口没有可用浏览器。未把HTTP验收标为全链浏览器验收。
- 浏览器工具不允许CDP网络拦截，本次未改变其权限或系统网络配置。页面源码已不含外网资源URL，静态脚本已通过本机HTTP实际读取，但仍需目标机禁网验收。

没有Windows实机，不写“Windows全部验收通过”。部署人员按WINDOWS_DEPLOYMENT.md的6项精简清单签收即可，不要求额外CI或全量人工穷举。

## 保留的现有业务边界

工时算法与随机种子未修改，跨重启不保证完全复现；内置法定日历覆盖2024—2026年。输入公式需要事先在Excel重算保存，输出需要Excel重算；本轮没有迁移旧数据库、携带真实项目数据或改变报告公式/格式。
