# MeasureM Windows 单机离线部署说明

适用：Windows 11 x64、标准 CPython 3.14 x64，每台电脑独立运行。电脑可以连接内部局域网；MeasureM 只允许本机浏览器访问，不需要互联网、共享数据库或入站防火墙规则。

**验收状态：**源码和 Mac 隔离环境验证已完成的项目见 `DEPLOYMENT_VALIDATION.md`；Windows 原生安装、批处理/PowerShell 和桌面 Excel 验收仍须由部署人员完成。wheel 可下载不等于已通过 Windows 实机验收。

## 1. 交付物和分工

开发方提供源码 ZIP、Windows 依赖锁定文件、本地 Alpine.js、空白模板、虚构演示项目和本说明。源码包通过 `python3 scripts/build_delivery.py` 构建，包含文件哈希与对应 Git commit，不包含真实数据库、历史导出或本机虚拟环境。

部署人员在联网 Windows 下载 Python 完整安装程序并准备 wheelhouse，通过交付介质带入每台离线目标机。Python 安装程序和 wheels 不上传 GitHub，不能把 Mac 的 `.venv` 复制到 Windows。

## 2. 联网 Windows：下载完整 Python

指定候选版本：**CPython 3.14.9，Windows x64，普通 GIL 构建**。

- [官方发布与校验信息](https://www.python.org/downloads/release/python-3149/)
- [完整安装程序 python-3.14.9-amd64.exe](https://www.python.org/ftp/python/3.14.9/python-3.14.9-amd64.exe)
- 发布页所列 SHA-256：`a247a0b612244a803f507b39fb7121e3602e4160d59c3b99d98e735289c53394`

```powershell
Get-FileHash .\python-3.14.9-amd64.exe -Algorithm SHA256
Get-AuthenticodeSignature .\python-3.14.9-amd64.exe
```

对照官方页面检查哈希和有效发布者签名，记录下载日期及来源。若指定补丁版无法取得，不自行改 minor/架构，先重新确认交付组合。不要下载需要联网再取运行时的 install manager、Microsoft Store 应用或 embeddable ZIP。

使用完整安装程序按当前用户安装，保留 pip、venv 和 Python launcher，选择标准运行时。准备机与目标机使用同一完整安装程序。安装到本机用户可写位置，无需专门安装 Word、SQLite 服务、Git、Node.js、Docker 或编译工具。

## 3. 联网 Windows：准备 wheelhouse

解压干净源码包到本地目录，打开 PowerShell 进入其中的 `MeasureM`。检查解释器：

```powershell
py -3.14 -c "import sys,struct,sysconfig; print(sys.version); print(struct.calcsize('P')*8); assert sys.version_info[:3] == (3,14,9) and struct.calcsize('P') == 8 and not sysconfig.get_config_var('Py_GIL_DISABLED')"
```

手动执行下面命令（逐条检查成功后再继续）：

```powershell
py -3.14 -m venv .prepare-venv
.\.prepare-venv\Scripts\python.exe -m pip --disable-pip-version-check --no-cache-dir download --only-binary=:all: --dest .\wheelhouse -r .\requirements-windows.lock.txt
Get-ChildItem .\wheelhouse\*.whl | Get-FileHash -Algorithm SHA256
```

锁定文件包含全部18个运行依赖的精确版本和 wheel 哈希，pip 会自动校验；必须全部成功，不可用源码包替代、不加 `--no-deps`。不要在线升级目标机 pip。当前标准库 venv 附带的 pip 即可执行这些命令。

也可以在允许执行本地脚本的环境运行 `& .\scripts\prepare_windows.ps1`，自动检查环境、下载和记录清单；若系统策略不允许执行脚本，直接使用上面的手动命令，无需修改全局执行策略。该脚本要求新的 `.prepare-venv` 与 `wheelhouse` 目录，避免混入旧包。

准备后的介质内容：源码 ZIP及其 `.sha256`、完整 Python安装程序、wheelhouse目录及部署人员记录的哈希/准备环境。源码 ZIP解压后应得到完整 MeasureM目录；wheelhouse另行放到该目录下。

## 4. 离线目标机：首次安装

先核对介质哈希，使用同一 Python 完整安装程序安装。将源码解压到新的本机可写目录，例如 `%LOCALAPPDATA%\MeasureM`，不要嵌套成 `MeasureM\MeasureM`，不要使用 Program Files、网络共享、UNC路径或云同步目录。

目录下首次应没有 `measurem.db`、历史 `generated` 或旧 `.venv`。不要覆盖已有安装。以下步骤以 PowerShell为例：

```powershell
Set-Location "$env:LOCALAPPDATA\MeasureM"
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip --disable-pip-version-check install --no-index --no-cache-dir --only-binary=:all: --find-links=.\wheelhouse -r .\requirements-windows.lock.txt
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe scripts\verify_install.py
```

`--no-index`禁止查网络索引；`--no-cache-dir`避免旧缓存掩盖缺包，lock内没有下载URL，缺包即失败。无需运行 Activate.ps1或激活虚拟环境。找不到 `py` 时，使用已核实的 Python 3.14.9 x64完整路径替换创建 venv那条命令。

`verify_install.py`使用临时数据库和虚构演示数据验证初始化、导入、叙述覆盖/恢复和导出；不会打开或修改正式数据库。应出现阶段1、阶段2和总体验证PASS。

## 5. 启动与使用

双击 `start_windows.bat` 或执行：

```powershell
.\start_windows.bat
```

浏览器手动打开 **http://127.0.0.1:8002**。保留服务窗口；结束时在服务窗口按 Ctrl+C。浏览器关闭不代表服务停止。同一台电脑只启动一份，端口8002不改为8000或8001。

启动器使用自身目录下 `.venv\Scripts\python.exe`，固定工作目录。`run.py`在Windows和Mac上都将数据库、静态资源和导出定位到项目根目录，检查该目录可写，仅监听127.0.0.1、关闭开发热重载。

首次真实启动自动建立空 `measurem.db`，无预置项目，不需要导入旧数据库。界面可先用 `examples\demo-project.xlsx`验证上传、节点、叙述编辑和导出；这是虚构测试项目，正式使用可在界面删除，再上传自己的输入文件。空白模板位于 `examples\input-template.xlsx`。

- 所有数据保存在本机 `measurem.db`；服务端副本保存到 `generated\项目ID\`，浏览器下载目录由浏览器设置。
- 目录必须允许创建SQLite日志和导出文件。无需把generated目录作为外部静态目录开放。
- 修改源码后需要手动重启。默认 `/docs`、`/redoc`关闭，`/openapi.json`仍提供本地API结构。
- 局域网其他电脑不能通过本机IP访问；无需断开内部局域网。

## 6. Excel要求与业务限制

程序生成 `.xlsx`不需要安装Office或Word；但**openpyxl不计算公式**。

1. 正式输入文件含公式时，先用桌面Excel重算并保存，避免导入空的或旧的公式缓存。
2. 输出文件用已授权、可离线使用的桌面Excel打开重算保存；验证无修复提示、无异常公式错误。Excel 2016是现有公式的兼容目标，最终以实际版本验收为准。WPS/LibreOffice须另行验证。
3. 宋体、黑体、Arial等字体影响显示/打印，运行不下载或分发字体。
4. 内置日历只覆盖2024—2026年；其他年份仅按周末判断。需要其他年份时先审核补充，不联网自动更新。
5. 本次没有改工时算法或随机种子。同进程内重叠节点工时一致；跨重启/跨机器完全复现不作保证。

## 7. 常见问题

| 现象 | 处理 |
|---|---|
| 浏览器打不开 | 检查服务窗口是否仍在运行、地址是否127.0.0.1:8002 |
| 8002被占用 | `Get-NetTCPConnection -LocalPort 8002`，或 `netstat -ano \| findstr :8002`；确认是否已有MeasureM，不杀未知进程、不改端口 |
| 缺.venv/ModuleNotFoundError | 用指定Python重新创建干净venv，确认离线安装命令成功；不要对系统Python盲目pip install |
| 找不到匹配wheel/哈希不一致 | 核对标准CPython3.14 x64与完整wheelhouse；回准备机补齐同一lock，不在目标机编译或绕过哈希 |
| 无法保存导出 | 关闭Excel占用的同名文件，检查目录权限、磁盘空间和路径长度，再重试 |
| 阶段名称无效 | 输入表阶段名去掉文件名禁用字符、目录分隔符、末尾空格/句点，缩短名称后重新导入 |
| 输入数据为空或公式显示为空 | 检查是否用桌面Excel重算保存；对照输入Sheet与列顺序 |
| 修改源码不生效 | 正常停止服务后重新启动；当前不启用自动热重载 |

异常堆栈/导入信息写入服务控制台，不增加日志上传。维护备份前先停止服务，把数据库及相关导出复制到获准的本地备份位置；不要向GitHub提交业务数据。卸载或清理数据须明确授权，本交付不提供自动清库。

## 8. 精简验收与签收

- [ ] 干净Windows目录/虚拟环境，禁外网、无pip缓存，离线安装与pip check成功。
- [ ] verify_install.py成功，首次真实启动为空项目列表。
- [ ] 无缓存浏览器加载本地Alpine.js，导入→编辑→导出演示流程成功，无外网资源请求。
- [ ] 8002只监听127.0.0.1；内部局域网其他电脑无法访问。
- [ ] 从其他当前目录启动、含空格/中文路径可用；端口占用/文件占用提示可理解。
- [ ] 代表性阶段1、双周、阶段、里程碑在实际Excel版本打开重算并检查关键值/格式。

部署人员记录：Windows版本/架构、Python补丁版、Excel版本、源码commit、wheelhouse哈希、验收日期、结果与未通过项。不要求新增CI、全量人工穷举或逐项反复跑完整测试。

## 9. Mac继续使用

已有Mac环境不需要安装Windows lock或重装依赖，继续 `python3 run.py`。数据库与generated位置不变；新增本地Alpine资源、回环监听与关闭热重载同样生效。新建Mac环境时使用 `requirements.txt`；Windows lock中的pydantic-core哈希对应win_amd64，不适用于Mac。
