# 终极项目 · 财务部 Agent 员工

把财务部自动化任务做成 opencode skills，用自然语言驱动。不修改原项目文件。

## 当前 Skills

| Skill | 项目 | 状态 | 验收 |
|-------|------|------|------|
| `receipt-export` | 项目二·回单查询导出 | ✅ 已交付 | 12/12 PASS |

## 目录结构

```
终极项目财务部Agent员工/
├── skills/receipt-export/SKILL.md    ← opencode skill 文件
├── scripts/
│   ├── receipt_export.py             ← CLI 脚本（零硬编码路径）
│   └── receipt_core/                 → symlink → 原项目核心代码
├── 工作区/                            ← 运行时自动创建的任务工作区
│   └── 回单任务/
│       ├── input/                    ← PDF 放这里
│       └── output/                   ← 导出结果
└── CLAUDE.md                         ← 本文件
```

## 设计原则

- **零硬编码路径**：脚本用 `Path(__file__).resolve().parent` 自发现位置，`check` / `init` 命令验证环境
- **不复制原代码**：`scripts/receipt_core` 是 symlink，指向原项目 `receipt_export_tool/receipt_core/`
- **智能工作区**：用户不指定路径时，自动用 `工作区/回单任务/input/` 和 `output/`
- **Skill = 行为指南**：SKILL.md 告诉 AI 怎么跟员工对话、什么时候问、什么时候直接干

## 添加新 Skill

1. 确认原项目有可调用的代码
2. `scripts/` 里建 symlink（如需）或写 CLI 封装
3. `skills/<name>/` 里写 SKILL.md（参考 receipt-export 的格式）
4. 用真实数据跑全部验收测试
5. 记录到本文

## 对应原项目

| 原项目 | symlink |
|--------|--------|
| `3_项目二_…/程序/receipt_export_tool/receipt_core/` | `scripts/receipt_core` |
