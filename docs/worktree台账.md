# worktree 台账 · finance-skills

> 规则：`AI开发规范/AI协作工作体系/worktree台账与开工检查.md`  
> 对账：2026-09-09。磁盘优先。

| 任务 | 工位 | 状态 | 备注 |
|------|------|------|------|
| 琪哥发票入金蝶 `qige-invoice-to-kingdee` | 正式 clone `项目/长期项目/财务部skills/finance-skills` · `main` | 使用中 | 兼容入口；不开 fork |
| 启动话术 + 更新技能 | 同上正式 clone · `main` | 使用中 | 同树 |
| 合入李尚核销 | 同上正式 clone · `main` | 使用中 | 已 merge `gitee/agent/ar-hexiao-daily-1-3-0` @ `c3ea05f` |
| 金蝶入账 `kingdee-posting` | 正式 clone · `main`（已含 018 + 拆腿/点头新建） | 使用中 | 本机已含收款找销售 `82d6a26`；推云后同事才能拉到 |
| 收款入金蝶 `016` | `~/.grok/worktrees/skills-finance-skills/kingdee-receipt-016` | 可删 | 018 已覆盖收款主路径并进 main |
| 三模块合一 `017` 后续 | `~/.grok/worktrees/skills-finance-skills/kingdee-posting-unify` detached `9ddbb67` | 可删 | 拆腿/点头新建已并入 main `c7a03c5` |
| 收款少问能跑完 `018` | `/Users/evanlee/.grok/worktrees/skills-finance-skills/kingdee-receipt-018` · `kingdee-receipt-018` @ `207f5b2` | 可删 | 已进 GitHub `main` |
| 月度损益表 `pl-dept-report` `013` | 正式 clone · `main` | 使用中 | 五条口径 + 结转发生额与核对同口径；总部左列优先科目余额 |
| 序时账入金蝶 `kingdee-gl-import` `014` | 同上正式 clone · `main` | 使用中 | 湖南分抄作业 |

`git worktree list` 本仓三间 fork：`kingdee-posting-unify`、`kingdee-receipt-016`、`kingdee-receipt-018`。
