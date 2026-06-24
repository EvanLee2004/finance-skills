# env-doctor（环境管家）

财务技能包的**基础设施技能**：让别的技能能跑起来。任何技能因为缺 Python 库或系统工具跑不动时，agent 查这里的清单，按**国内镜像优先**把环境装齐、再重试。

- **研发维护**：李明昊 / 甲骨易财务部，2026-06。
- **纯提示词驱动，无脚本**：核心是 [`config/依赖与安装清单.md`](config/依赖与安装清单.md) 这张活表——"哪个技能要什么、缺了怎么装（国内/国外）"。加新技能就改这张表。
- **不碰业务数据**：只装/配环境，不读应收表、发票、工资等任何数据。

## 怎么用

1. **同事主动**说「配下环境 / 把依赖装齐」→ agent 照清单「一把全装」命令装齐（国内镜像）。
2. **任何技能报缺依赖**（ModuleNotFoundError、找不到 soffice/tesseract 等）→ agent 自动查清单、按国内镜像装上、重试原技能，不用同事操心。

## 覆盖范围（截至 2026-06-24，8 个技能）

- Python 核心库：pandas、openpyxl、pypdf、pdfplumber、pdf2image、python-docx、python-pptx、markitdown、lxml、defusedxml、Pillow
- PDF 高级（按需）：pytesseract、reportlab、pypdfium2
- 系统级：LibreOffice、poppler、tesseract、Node.js
- Python ≥3.10（可选校验路径需要）

详见 [`config/依赖与安装清单.md`](config/依赖与安装清单.md)。
