---
name: update-finance-skills
description: >-
  从 Gitee 更新或安装财务部官方技能包。当用户说「更新 / 更新财务skills / 更新财务技能 /
  更新财务技能包 / 更新技能包 / 升级技能 / 同步技能 / 拉最新 / 安装财务skills /
  装财务技能 / 有哪些财务技能 / 该用哪个技能 / 技能怎么说 / 启动时说什么 /
  不知道用哪个skill」时用本技能。核销、应收合并、九点下单、琪哥发票等具体业务
  不要用本技能去跑，只负责更新/导览。核销已合入 main（李尚最新），更新时一并覆盖。
---

# 更新财务技能 / 说什么用哪个

同事说「更新」或「该用哪个」时用本技能。**源永远先 Gitee** `https://gitee.com/Lee157/finance-skills` 的 `main`。

## 启动时说

- **更新**：`更新财务skills`（也可：更新财务技能 / 更新技能包 / 拉最新）
- **第一次装**：`安装财务skills`
- **不知道用哪个**：`有哪些财务技能` 或把活用大白话说一遍，对照下面这张表

## 更新时你必须做

1. 跑本技能脚本，不要手搓复制、不要问她要 zip：

```bash
python3 "<本skill目录>/scripts/update.py"
```

Windows 同事的目录一般是 `%USERPROFILE%\.config\opencode\skills\`。找不到再问她。

2. 脚本会：优先 `git clone/pull` **Gitee main**（不通才 GitHub）→ 只覆盖白名单 → **本机已有 `config/` 默认不覆盖**（核销除外：`ar-hexiao-daily` 的 config 跟仓库走，李尚口径在仓里）→ 根上放下 `财务技能包_来源与更新.md` 和 `财务技能_说什么用哪个.md`。
3. **禁止**清空整个 skills；**禁止**删白名单外她自己装的夹。
4. 汇报照抄脚本打印：SHA、更新了哪些、config 保住没。提醒**重启 opencode**。

核销已合入 `main`（李尚 `agent/ar-hexiao-daily-1-3-0`）。`pack.json` 的 `overwrite_config` 含 `ar-hexiao-daily`：核销业务规则跟仓库走。`config.local.json` / 金蝶 `kingdee.local.json` 一类本机凭据仍不进仓、不覆盖。

## 说什么用哪个（先对这张表，再开对应技能）

| 她说的话 | 用这个技能 |
|----------|------------|
| 更新 / 更新财务skills / 安装财务skills | **本技能** |
| 跑本周应收 / 应收合并 / 做应收all | `receivables-merge` |
| 把 all 拆给各销售 / 按销售拆分 | `split-by-sales` |
| 做这个月劳务发票统计 / 劳务发票核对 | `labor-invoice-check` |
| 申报表重命名 / 代扣代缴 PDF 改名 | `withholding-report-rename` |
| 本周抽谁 / 合规抽查 | `compliance-spot-check` |
| 对比追觅进度 / 追觅这两版 | `dreame-ar-progress-diff` |
| 跑本月部门费用归集 / 部门科目余额表 | `dept-expense-alloc` |
| 跑昨天的核销 / 日清 / 重扫挂账 | `ar-hexiao-daily`（只跑，不覆盖她本机那份） |
| 九点下单 / 跑下单汇总 | `order-daily-summary` |
| 销项发票入金蝶 / 付款入金蝶 / 收款入金蝶 / 琪哥发票入金蝶 | `kingdee-posting`（旧触发仍可进销项） |
| 序时账入金蝶 / 抄作业入金蝶 | `kingdee-gl-import` |
| 月度损益表 / 科目余额表 | `pl-dept-report` |
| 帮我理清需求 / 我不知道怎么说 | `task-clarifier` |
| 配下环境 / 缺库跑不起来 | **不是技能**。当场清华镜像装，失败再阿里 / 中科大 / 默认源。本机若还剩 `env-doctor` 夹，当废纸 |
| 做 Excel / 改这张表（零散、不够独立技能） | `xlsx` |
| 做 Word / 改这份 Word | `docx` |
| 做 PPT | `pptx` |
| 处理 PDF / 合并 pdf | `pdf` |

细表见仓内 `skills/财务技能_说什么用哪个.md`（更新后会放到本机 skills 根上）。

## 硬禁止

- 不要用 GitHub 当第一源。
- 不要为了「目录只保留白名单」去删她别的技能。
- 不要为了「目录整齐」去删她别的技能。核销 config 跟仓库（见 `overwrite_config`）。
- 不要在更新时跑核销写表或金蝶引入。
