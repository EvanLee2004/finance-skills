# pptx · PPT 演示文稿处理技能

**甲骨易财务部自研**通用文档技能（研发维护：李明昊，2026-06）。服务于财务同事在 opencode 上的日常演示文稿需求，与 receivables-merge 等业务技能同属 `finance-skills` 技能包。

## 能干什么

- 新建/编辑 .pptx，改模板、加页、抽正文
- 缩略图预览：`python3 scripts/thumbnail.py <文件.pptx>`

## 结构

```
pptx/
├── SKILL.md          # agent 行为指南 + 触发词
├── README.md         # 本文件
├── editing.md        # 编辑参考
├── pptxgenjs.md      # 从零创建参考
└── scripts/          # thumbnail、office 工具链等
```

## 依赖（纯本地）

- Python：python-pptx、Pillow、defusedxml
- 系统：LibreOffice（`soffice`，缩略图转 PDF 用）

## 维护

研发：李明昊 · 甲骨易财务部 · 改动走 `finance-skills` 仓库 commit/push，同步重打 `部署/财务技能包.zip`。