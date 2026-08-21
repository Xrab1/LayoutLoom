# 页织工坊（LayoutLoom）源代码获取说明

页织工坊是采用 **GNU AGPL v3 或更高版本**发布的自由开源软件。

> **运行环境提示：** 本项目已对桌面版 WPS Office 与 Microsoft Office 的真实 COM 引擎完成独立识别和定向适配；自动模式按 WPS → Microsoft Office → LibreOffice 尝试，也可锁定单一引擎。LibreOffice 作为兼容回退，复杂版式建议比较本机 WPS 与 Microsoft Office 的实际结果。

Windows 便携版应与同版本的 `LayoutLoom-Source-<版本>.zip`、`LICENSE`、`THIRD_PARTY_NOTICES.md` 和 `SHA256SUMS.txt` 一同发布。源码压缩包包含构建该版本所需的页织工坊源代码、测试、构建脚本、Agent JSON CLI 和可分发 Codex Skill；体积较大的第三方二进制引擎不重复放入源码包，其固定版本、来源和许可证见 README、准备脚本及便携包中的 `THIRD_PARTY_LICENSES`。

任何人都可以免费使用、研究、修改和再分发本软件。分发修改版，或通过网络向用户提供修改版功能时，请依照 AGPL 提供对应源代码并保留许可证和版权说明。

WPS Office、Microsoft Office 与 LibreOffice 不是本项目便携包的一部分，也不随页织工坊分发。软件只检测并调用用户本机已有的合法安装；Office 自动选择优先级仍为 WPS → Microsoft Office → LibreOffice。

## v0.2.0 Agent 接口

源码版提供 `layoutloom-agent` 入口及 `python -m docuforge.cli agent ...` 调用方式；Windows 便携版提供独立控制台入口 `LayoutLoom-CLI.exe`，并在 `agent_skill\layoutloom-agent` 中携带同版本 Codex Skill。相关实现、协议、安装方式和安全边界见 [`AGENT_INTEGRATION.md`](AGENT_INTEGRATION.md)。

Agent 接口完全在本机运行，与 GUI 共用公开任务目录和处理实现。当前版本不内置云端服务或常驻 MCP Server，也不会为了 Agent 调用开放本地网络端口；采用稳定的标准输入/输出 JSON 协议，方便任何自动化宿主直接调用，也为第三方按需封装 MCP 提供统一底座。

已知、非破坏性的常见任务可通过 `quick-run` 在单个进程内完成请求构造、完整校验和执行；未知或复杂任务仍可使用 `catalog`、`describe`、`validate` 与 `run` 的显式流程。两种入口共用同一处理实现，不会产生精度不同的第二套引擎。

若再分发修改后的 Agent CLI、Skill 或基于修改版页织工坊提供网络功能，仍应依照 GNU AGPL v3 或更高版本提供对应源代码，并保留许可证和版权说明。具体义务以 `LICENSE` 原文为准。
