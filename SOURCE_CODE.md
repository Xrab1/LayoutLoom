# 页织工坊（LayoutLoom）源代码获取说明

页织工坊是采用 **GNU AGPL v3 或更高版本**发布的自由开源软件。

> **运行环境提示：** 本项目已对桌面版 WPS Office 与 Microsoft Office 的真实 COM 引擎完成独立识别和定向适配；自动模式按 WPS → Microsoft Office → LibreOffice 尝试，也可锁定单一引擎。LibreOffice 作为兼容回退，复杂版式建议比较本机 WPS 与 Microsoft Office 的实际结果。

Windows 便携版应与同版本的 `LayoutLoom-Source-<版本>.zip`、`LICENSE`、`THIRD_PARTY_NOTICES.md` 和 `SHA256SUMS.txt` 一同发布。源码压缩包包含构建该版本所需的页织工坊源代码、测试和构建脚本；体积较大的第三方二进制引擎不重复放入源码包，其固定版本、来源和许可证见 README、准备脚本及便携包中的 `THIRD_PARTY_LICENSES`。

任何人都可以免费使用、研究、修改和再分发本软件。分发修改版，或通过网络向用户提供修改版功能时，请依照 AGPL 提供对应源代码并保留许可证和版权说明。

WPS Office、Microsoft Office 与 LibreOffice 不是本项目便携包的一部分，也不随页织工坊分发。软件只检测并调用用户本机已有的合法安装；Office 自动选择优先级仍为 WPS → Microsoft Office → LibreOffice。
