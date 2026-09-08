# worktree 台账 · finance-skills

> 规则：`AI开发规范/AI协作工作体系/worktree台账与开工检查.md`  
> 对账：2026-09-08。磁盘优先。

| 任务 | 工位 | 状态 | 备注 |
|------|------|------|------|
| 琪哥发票入金蝶 `qige-invoice-to-kingdee` | 正式 clone `项目/长期项目/财务部skills/finance-skills` · `main` | 使用中 | 兼容入口；不开 fork |
| 启动话术 + 更新技能 | 同上正式 clone · `main` | 使用中 | 同树 |
| 合入李尚核销 | 同上正式 clone · `main` | 使用中 | 已 merge `gitee/agent/ar-hexiao-daily-1-3-0` @ `c3ea05f` |
| 金蝶入账 `kingdee-posting` | 正式 clone · `main` | 使用中 | GitHub `b527ee1` 已含 017 三入口；本地又合入李尚核销 |
| 收款入金蝶 `016` | `~/.grok/worktrees/skills-finance-skills/kingdee-receipt-016` | 使用中 | 未并入正式 clone。017 已覆盖收款主路径 |
| 三模块合一 `017` 后续 | `~/.grok/worktrees/skills-finance-skills/kingdee-posting-unify` detached `9ddbb67` | 使用中 | 比 GitHub 多「多条1131拆腿 / 点头新建客户」。本轮不并入 |
| 月度损益表 `pl-dept-report` `013` | 正式 clone · `main` | 使用中 | 网页登录改 `xingchen.local.json` |
| 序时账入金蝶 `kingdee-gl-import` `014` | 同上正式 clone · `main` | 使用中 | 湖南分抄作业 |

`grok worktree list` 本仓两间 fork：`kingdee-posting-unify`、`kingdee-receipt-016`。未并入正式 clone。
