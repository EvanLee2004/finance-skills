# docx · Word 文档处理技能

**甲骨易财务部自研**通用文档技能（研发维护：李明昊，2026-06）。服务于财务同事在 opencode 上的日常 Word 建/改/解析，与 receivables-merge 等业务技能同属 `finance-skills` 技能包。

## 能干什么

- 新建/编辑 .docx（报告、备忘录、函件、带格式文档）
- 读正文、改批注与修订、插图、unpack/repack 精细改 XML
- 复杂校验走 `scripts/office/validate.py`（需本机 LibreOffice）

## 结构

```
docx/
├── SKILL.md          # agent 行为指南 + 触发词
├── README.md         # 本文件
└── scripts/          # 解包/打包/校验/批注等 Python 工具
```

## 依赖（纯本地）

- Python：python-docx、defusedxml
- 系统：LibreOffice（`soffice`）、pandoc（按需）

## 维护

研发：李明昊 · 甲骨易财务部 · 改动走 `finance-skills` 仓库 commit/push，同步重打 `部署/财务技能包.zip`。