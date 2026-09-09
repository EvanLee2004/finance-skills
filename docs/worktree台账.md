# worktree 台账 · finance-skills

> 规则：`AI开发规范/AI协作工作体系/worktree台账与开工检查.md`  
> 对账：2026-09-09。磁盘优先。

| 任务 | 工位 | 状态 | 备注 |
|------|------|------|------|
| 琪哥发票入金蝶 `qige-invoice-to-kingdee` | 正式 clone `项目/长期项目/财务部skills/finance-skills` · `main` | 使用中 | 兼容入口；不开 fork |
| 启动话术 + 更新技能 | 同上正式 clone · `main` | 使用中 | 同树 |
| 合入李尚核销 | 同上正式 clone · `main` | 使用中 | 已 merge `gitee/agent/ar-hexiao-daily-1-3-0` @ `c3ea05f` |
| 金蝶入账 `kingdee-posting` | 正式 clone · `main` | 使用中 | GitHub `1dc1b7b`；勿写（损益表脏改占用） |
| 收款入金蝶 `016` | `~/.grok/worktrees/skills-finance-skills/kingdee-receipt-016` | 使用中 | 未并入正式 clone。017 已覆盖收款主路径 |
| 三模块合一 `017` 后续 | `~/.grok/worktrees/skills-finance-skills/kingdee-posting-unify` detached `9ddbb67` | 使用中 | 比 GitHub 多「多条1131拆腿 / 点头新建客户」。本轮不并入 |
| 收款少问能跑完 `018` | `/Users/evanlee/.grok/worktrees/skills-finance-skills/kingdee-receipt-018` · `kingdee-receipt-018` @ `0fb2292` | 使用中 | 基线 `1dc1b7b`。职员部门回退 + 待确认候选。未 push，等人说并入 / 可删 |
| 月度损益表 `pl-dept-report` `013` | 正式 clone · `main` | 使用中 | 2026-09-09 薪酬台账接入，未 commit |
| 序时账入金蝶 `kingdee-gl-import` `014` | 同上正式 clone · `main` | 使用中 | 湖南分抄作业 |

`git worktree list` 本仓三间 fork：`kingdee-posting-unify`、`kingdee-receipt-016`、`kingdee-receipt-018`。未并入正式 clone。
