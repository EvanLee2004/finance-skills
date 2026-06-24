# pdf · PDF 处理技能

**甲骨易财务部自研**通用文档技能（研发维护：李明昊，2026-06）。服务于财务同事在 opencode 上的日常 PDF 需求，与 receivables-merge 等业务技能同属 `finance-skills` 技能包。

## 能干什么

- 读文/抽表、合并拆分、旋转、水印、填表、转图
- 进阶：OCR 扫描件（需额外装 pytesseract，见 SKILL.md）

## 结构

```
pdf/
├── SKILL.md          # agent 行为指南 + 触发词
├── README.md         # 本文件
├── reference.md      # 库与命令参考
├── forms.md          # 填表专题
└── scripts/          # 转图、填表、结构提取等
```

## 依赖（纯本地）

- Python：pypdf、pdfplumber、reportlab；转图另需 pdf2image
- 系统：poppler（pdftoppm）

## 维护

研发：李明昊 · 甲骨易财务部 · 改动走 `finance-skills` 仓库 commit/push，同步重打 `部署/财务技能包.zip`。