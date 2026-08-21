# 页织工坊 Agent 集成指南

本指南适用于页织工坊 v0.2.0 的本地 Agent JSON CLI 和 Codex Skill。Agent 接口不会另起一套转换算法：它与桌面 GUI 共用同一任务目录、参数规范、引擎检测、取消机制和输出校验，因此自动化结果与在 GUI 中选择相同任务及参数时遵循相同处理逻辑。

## 为什么采用 JSON CLI + Codex Skill

当前实现由两层组成：

1. **稳定 JSON CLI**：所有 Agent 和脚本都可调用的统一底座，负责能力发现、参数校验、任务执行、进度事件、取消和退出码。
2. **Codex Skill**：自然语言入口和操作规范。明确、可逆的常见任务默认一次调用；未知或复杂任务才进入能力发现、独立预校验和执行流程。

当前版本不内置常驻 MCP Server。这是针对本地文件处理场景的产品选择，而不是否定 MCP：

- CLI 在源码版和 Windows 便携版中都能直接运行，不需要用户配置端口、后台服务或额外账户。
- 每次任务使用独立进程，路径授权、取消和故障边界更直观，也便于其他 Agent、PowerShell 或 CI 调用。
- 页织工坊任务通常耗时较长且会生成本地文件，JSON Lines 已能稳定传递实时进度和最终结果。
- 稳定 CLI 本身就是未来 MCP 适配器的底座。确有多会话共享、集中队列或远程能力发现需求时，可增加薄 MCP 层，无需重写处理引擎。

## 选择正确的命令入口

| 环境 | 命令前缀 |
| --- | --- |
| Windows 便携版 | `.\LayoutLoom-CLI.exe agent` |
| 源码虚拟环境（运行过 `install.ps1`） | `.\.venv\Scripts\layoutloom-agent.exe` |
| 已安装 Python 包 | `layoutloom-agent` |
| 源码通用写法 | `.\.venv\Scripts\python.exe -m docuforge.cli agent` |

以下示例使用源码虚拟环境。便携版用户只需把开头的 `.\.venv\Scripts\layoutloom-agent.exe` 换成 `.\LayoutLoom-CLI.exe agent`。

## 最快推荐方式：一次调用

当任务 ID、输入文件和输出目录已经明确时，直接使用 `quick-run`。它在内存中构造请求，并在同一进程内完成参数、路径、引擎、取消、输出锁和最终文件校验，不要求先生成请求文件，也不要求单独运行 `validate`：

```powershell
.\.venv\Scripts\layoutloom-agent.exe quick-run word.to_pdf `
  "C:\资料\合同.docx" `
  --output-dir "C:\资料\输出" `
  --param engine=wps `
  --allow-root "C:\资料" `
  --format jsonl
```

已知非敏感参数可重复使用 `--param key=value`。在当前面向 WPS 优化的安装中，Word/Excel/PPT 转 PDF 默认显式使用 `--param engine=wps`；只有用户明确选择其他引擎，或 WPS 已明确报告不可用时才改变。密码不能出现在命令行，必须放在临时 UTF-8 JSON 参数对象中并通过 `--params-file` 传入。

对 Word 转 PDF、Excel 转 PDF、PPT 转 PDF、PDF 合并/拆分/压缩、常见图片处理和视频转码等明确另存任务，Agent 应优先使用这一入口，不应重复执行协议检查、目录查询、参数描述和独立预校验，也不应再次要求用户确认已经清楚授权的普通另存操作。

Office 快速任务只启动一个 `quick-run`。调用方应持续读取 JSONL，最多等待 90 秒；若仍无最终事件，只发送一次温和中断并再等待最多 15 秒清理。旧任务和包装器完全退出、输出锁释放前不得重试，也不得在原终端会话中输入第二条转换命令。页织工坊会自动为同名结果生成唯一文件名，因此无需事先检查或轮询旧输出文件。

## 发现与复杂任务工作流

以下流程用于首次诊断、未知任务、陌生参数、敏感参数、多辅助资源、复杂批量或可能移动源文件的操作；不要求每个常见任务都完整重复一遍。

### 1. 检查协议

```powershell
.\.venv\Scripts\layoutloom-agent.exe protocol --pretty
```

返回应用版本、Agent 协议版本、传输格式和稳定退出码。v0.2.0 的协议版本为 `1.0`。

### 2. 查询任务目录

```powershell
.\.venv\Scripts\layoutloom-agent.exe catalog --query "PDF Word" --pretty
.\.venv\Scripts\layoutloom-agent.exe catalog --group "图片效果与编辑" --pretty
```

`catalog` 默认不执行所有外部引擎的完整探测，以免单纯查询目录时启动耗时检查。需要获知所有候选引擎的当前状态时添加 `--probe`；需要完整参数摘要时添加 `--full`。Agent 不应凭经验猜测任务 ID。

### 3. 读取任务参数

```powershell
.\.venv\Scripts\layoutloom-agent.exe describe pdf.merge --pretty
```

`describe` 返回输入扩展名、最少/最多文件数、参数类型、默认值、选项范围、输出策略、引擎状态和精度说明。创建请求前应以这里返回的定义为准。

### 4. 创建 UTF-8 JSON 请求

```json
{
  "schema_version": "1.0",
  "request_id": "merge-example-001",
  "operation": "pdf.merge",
  "inputs": [
    "D:\\资料\\封面.pdf",
    "D:\\资料\\正文.pdf"
  ],
  "output_dir": "D:\\资料\\输出",
  "parameters": {
    "filename": "完整文档"
  },
  "options": {
    "expand_globs": false
  }
}
```

请求文件最大为 1 MiB。字段含义：

- `schema_version`：当前使用 `1.0`。
- `request_id`：可选的调用方追踪 ID；省略时由页织工坊生成。
- `operation`：必须来自 `catalog` 或 `describe`。
- `inputs`：输入文件数组。相对路径按请求文件所在目录解析；为了审计清楚，推荐绝对路径。
- `output_dir`：输出文件夹，不是具体输出文件名。
- `parameters`：只允许该任务公开声明的参数。
- `options.expand_globs`：默认 `false`；显式设为 `true` 时才展开通配符。

重复 JSON 字段、未知顶层字段、未知选项、未知参数、无效选择、缺失文件和不支持的扩展名都会直接校验失败。

### 5. 先校验，不写入输出

```powershell
.\.venv\Scripts\layoutloom-agent.exe validate --request .\request.json --pretty `
  --allow-root "D:\资料"
```

`validate` 检查请求结构、输入文件、参数、所需引擎和路径边界，但不会创建输出目录或处理文件。可重复提供 `--allow-root`：

```powershell
.\.venv\Scripts\layoutloom-agent.exe validate --request .\request.json --pretty `
  --allow-root "D:\输入资料" `
  --allow-root "E:\输出结果"
```

一旦使用 `--allow-root`，所有输入、输出目录和 path 类型辅助参数都必须位于至少一个获准根目录内。

### 6. 执行并读取 JSONL 事件

```powershell
.\.venv\Scripts\layoutloom-agent.exe run --request .\request.json --format jsonl `
  --allow-root "D:\资料"
```

标准输出的每一行都是独立 JSON 对象，`seq` 单调递增。事件包括：

- `accepted`：请求已接受，尚未完成。
- `progress`：包含 `fraction`、`percent` 和阶段说明。
- `cancel_requested`：已收到优雅取消请求，正在收尾。
- `result`：唯一的最终成功或部分成功结果。
- `error`：唯一的最终错误结果；取消时可能包含 `partial_result`。

调用方必须持续读取到最终 `result` 或 `error`。不要把 `accepted`、任意中间进度或 `progress=100` 当作完成。只想读取单个最终 JSON 时可以使用：

```powershell
.\.venv\Scripts\layoutloom-agent.exe run --request .\request.json --format json --pretty
```

### 7. 核验输出

最终结果会列出 `outputs`、`warnings`、`completed_inputs`、`failed_inputs` 和 `cancelled_inputs` 等信息。Agent 在向用户报告成功前，还应确认每个报告的输出路径真实存在且文件非空。

独立批量任务中，单个输入失败不会立即停止整个队列。退出码为 `5` 时，应保留并报告成功输出，同时逐项说明失败输入。

## 稳定退出码

| 退出码 | 机器含义 | 调用方处理 |
| ---: | --- | --- |
| `0` | success | 核验并报告全部输出 |
| `2` | invalid request or usage | 修正 JSON、任务 ID、参数、路径或命令 |
| `3` | engine unavailable | 告知缺少的 WPS、Microsoft Office、LibreOffice、FFmpeg、Poppler 或其他引擎，不要暗中换引擎 |
| `4` | handled runtime failure | 读取错误信息或逐文件失败原因 |
| `5` | partial success | 保留成功输出并报告失败清单 |
| `70` | internal error | 保留日志和请求，作为未预期错误处理 |
| `130` | cancelled | 检查 `partial_result`，保留已完成输出 |

## 取消和批量行为

运行中发送一次 `Ctrl+C`、`SIGTERM` 或 Windows `SIGBREAK` 即可请求优雅取消。页织工坊会让当前任务进入收尾流程，保留已完成文件并清理未完成或空的临时输出。调用方应等待最终事件和退出码，不应一开始就强杀进程；只有进程长期无法响应时才由宿主执行更高级别的故障处置。

有先后顺序的组合任务会保持输入顺序。可独立处理的批量任务会在单个文件失败后继续执行，并在最终结果中汇总。

## 安全边界

- Agent 只能调用公开目录中的任务，不能通过请求自行声明新任务或新参数。
- 使用 `--allow-root` 可以把一次调用严格限制在用户授权的文件夹内，建议面向第三方 Agent 时默认启用。
- `image.rename` 的 `move=true` 会移动原文件，Agent 模式默认阻止。只有用户明确授权后，宿主才可在命令行添加 `--allow-source-mutation`；请求 JSON 无法给自己提升权限。
- 密码只应存放在临时 UTF-8 请求文件中，不要使用人工 CLI 的 `-p` 参数传密码，也不要把请求内容写入日志。校验和结果只返回敏感参数名，不返回密码值。
- Agent 不应在引擎不可用时私自切换用户指定的处理模式。特别是 `pdf.to_word` 的 Microsoft Word 原生模式会单独检查真实 Word 是否可用。
- `video.repair_slides_ppt` 依赖 GUI 快速补修窗口创建的方案，属于“GUI 预建方案 + Agent 执行”任务，不能由纯 JSON 请求替代框选和人工选帧。
- 页织工坊默认另存输出。应把源文件与输出目录分开，并在开始前确认用户对输入和输出根目录拥有合法处理权限。

## 安装 Codex Skill

### 便携版

在解压后的页织工坊目录运行：

```powershell
.\LayoutLoom-CLI.exe agent install-skill --pretty
```

### 源码版

```powershell
.\.venv\Scripts\layoutloom-agent.exe install-skill --pretty
```

默认安装到当前用户的 Codex skills 目录；如需指定其他根目录：

```powershell
.\.venv\Scripts\layoutloom-agent.exe install-skill `
  --skills-dir "D:\CodexData\skills" --pretty
```

如果 `layoutloom-agent` 已存在，确认升级时添加 `--force`。安装器会先备份旧 Skill，再写入当前 CLI 的实际命令位置。安装完成后需重新打开 Codex，使其重新发现 Skill。

之后可以直接向 Codex 表达本地文件处理目标，例如：

> 使用页织工坊把 D:\资料\封面.pdf 和 D:\资料\正文.pdf 按这个顺序合并，输出到 D:\资料\输出；只允许访问 D:\资料。

Skill 会让明确的常见任务直接执行 `quick-run → 核验输出`；只有未知或复杂任务才使用“检查协议/查询目录/describe → 创建请求 → validate → run → 核验输出”。两条路径都不会绕过用户授权、路径限制或引擎依赖。

## 供其他 Agent 或未来 MCP 使用

任何能启动本地进程并读取 UTF-8 标准输出的宿主，都可以直接集成 `LayoutLoom-CLI.exe agent`。建议宿主：

1. 缓存 `protocol` 和不带 `--probe` 的目录结果，并按应用/协议版本失效。
2. 对已知、非破坏性、参数明确的操作直接调用 `quick-run`；不要在每次任务前重复 `describe` 和独立 `validate`。
3. 仅对未知参数、复杂批量、敏感参数、外部引擎不确定或源文件变更任务执行独立 `describe`/`validate`。
4. 把 JSONL 标准输出作为机器通道；把标准错误保存为诊断信息，不与 JSONL 混合解析。
5. 为每次调用设置明确的允许根目录，并由宿主控制源文件修改权限。
6. 使用退出码决定成功、部分成功、取消或错误，不通过文本关键字猜测状态。

未来 MCP Server 最适合做这一 CLI 的薄封装：MCP tool schema 可由 `catalog/describe` 生成，实际任务仍交给相同的 `validate/run` 流程。这样可以保持 GUI、CLI、Skill 和 MCP 的任务定义一致，避免多套实现产生精度与安全差异。
