# worktree 台账 · finance-skills

> 规则：`AI开发规范/AI协作工作体系/worktree台账与开工检查.md`  
> 对账：**2026-09-21**。磁盘优先。

| 任务 | 工位 | 状态 | 备注 |
|------|------|------|------|
| 琪哥发票入金蝶 `qige-invoice-to-kingdee` | — | 已下线 | 2026-09-15 源码夹删除；销项走 `kingdee-posting` |
| 启动话术 + 更新技能 | 同上正式 clone · `main` | 使用中 | 同树 |
| 合入李尚核销 | 同上正式 clone · `main` | 使用中 | 已 merge `gitee/agent/ar-hexiao-daily-1-3-0` @ `c3ea05f` |
| 金蝶入账 `kingdee-posting` | 正式 clone · `main`（已含 018 + 拆腿/点头新建） | 使用中 | 付款已拆出；本技能只销项/收款 |
| 供应商付款 `kingdee-payment` | 正式 clone · `main`（不开新 worktree） | 使用中 | 2026-09-21 已推 `5d35e23`（摘要改销方名、全批一张凭证、去掉网页登录） |
| 收款入金蝶 `016` | ~~`~/.grok/worktrees/skills-finance-skills/kingdee-receipt-016`~~ | **已拆**（2026-09-21） | 独有提交 `3d91164`（`receipt_io.py` 路线）已被 main 的 `assist_xlsx.py` + `convert.py` 收款实现取代；已归档 bundle |
| 三模块合一 `017` 后续 | ~~`~/.grok/worktrees/skills-finance-skills/kingdee-posting-unify`~~ | **已拆**（2026-09-21） | 拆腿/点头新建已并入 main（`create-new-customers` 在 `convert.py` 里）；已归档 bundle |
| 收款少问能跑完 `018` | ~~`~/.grok/worktrees/skills-finance-skills/kingdee-receipt-018`~~ | **已拆**（2026-09-21） | 分支 tip `207f5b2` 已在 main，分支与远程分支保留；30 个未提交改动已存 patch |
| 月度损益表 `pl-dept-report` `013` | 正式 clone · `main` | 使用中 | 已双端 `40008ff`：先源再核算 + 材料夹/期间不符修补 |
| 序时账入金蝶 `kingdee-gl-import` `014` | 同上正式 clone · `main` | 使用中 | 湖南分抄作业 |
| 018 李尚平台并入正本 | 正式 clone · `main`（人要求不开 worktree） | 使用中 | 全员 version 1.0；删 platform-skills；只推 Gitee |

**本仓当前无任何 worktree fork**（2026-09-21 清空；`git worktree list` 只剩正式 clone 的 `main@5d35e23`）。

归档（防丢，共 ~124KB）：`技能/金蝶/工作区/_worktree归档_20260921/`
- `016-kingdee-receipt-draft.bundle`、`unify-9ddbb67.bundle`（各含独有提交，可 `git fetch <bundle>` 取回）
- `018-未提交改动.patch`、`018-未跟踪清单.txt`
