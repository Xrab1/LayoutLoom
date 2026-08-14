# 页织工坊（LayoutLoom）第三方组件与许可证说明

页织工坊主项目采用 GNU Affero General Public License v3.0 或更高版本（AGPL-3.0-or-later）。发布包还包含或调用下列主要第三方组件。实际发布时应保留各组件随附的完整许可证文件；本清单不替代原许可证正文。

| 组件 | 主要用途 | 许可证/说明 |
| --- | --- | --- |
| CPython 3.12、Tcl/Tk | 便携版语言运行时与桌面界面 | PSF License 与 Tcl/Tk BSD 风格许可证；原文随发布包保留 |
| PyInstaller | Windows one-folder 打包 | GPL-2.0-or-later，并带允许分发打包应用的特殊例外 |
| PyMuPDF / MuPDF | PDF 渲染、版面分析与转换 | GNU AGPL v3 或 Artifex 商业许可；本项目选择 AGPL 路线 |
| Poppler | PDF 页面渲染 | GPL 系列；作为独立本地可执行引擎随包分发，实际版本见同目录构建清单 |
| FFmpeg | 视频转码、压缩、裁剪、音频提取 | 随构建而定；当前便携引擎为 gyan.dev GPLv3 构建，许可证、源码提交号和构建选项随包保留 |
| pdf2docx、pypdf、pdfplumber、pdf2image | PDF 结构与文本处理 | MIT/BSD 等宽松许可证，以各项目随附文件为准 |
| Pillow、NumPy、OpenCV | 图片处理与视觉分析 | HPND/BSD/Apache-2.0 等，以各项目随附文件为准 |
| python-docx、openpyxl、python-pptx | Office Open XML 处理 | MIT |
| ReportLab | PDF 生成 | BSD |
| tkinterdnd2 | 文件拖放 | MIT；底层 TkDND 许可证随组件保留 |
| Real-ESRGAN NCNN Python binding | 可选 GPU 图像增强 | BSD-3-Clause；模型和底层 NCNN/Vulkan 组件按各自许可证发布 |
| pywin32 | Windows COM 桥接 | PSF |

WPS Office、Microsoft Office 和 LibreOffice 不属于页织工坊发布包，也不会被复制或重新分发。软件只检测和调用用户电脑上已有的合法安装；WPS 与 Microsoft Office 均使用各自的真实 COM 自动化接口并接受独立适配，LibreOffice 可作为用户自行安装的第三顺位兼容转换引擎。

`build.ps1` 会把虚拟环境中各 Python 包的 `LICENSE`、`COPYING`、`NOTICE` 等原始文件自动汇入 `dist/LayoutLoom/THIRD_PARTY_LICENSES/Python-packages`。发布维护者仍应在每次升级依赖后重新核对许可证清单、二进制来源、源代码链接和实际构建选项。
