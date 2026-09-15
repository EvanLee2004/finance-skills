# 财务部 skills

一件活一个夹，都在 `skills/`。没有 `platform-skills/`。

对 OpenCode 说人话即可。开口对照：[`skills/财务技能_说什么用哪个.md`](skills/财务技能_说什么用哪个.md)。

更新：说 **「更新财务skills」**。别人推上 Gitee 了，不等于你电脑已经更新。

## 版本号（给李尚 / 给 agent）

2026-09-15 起，每个 skill 从 **1.0** 计。写在该夹 `SKILL.md` YAML 的 `version` 字段。

| 你改了什么 | 怎么记版本 |
|------------|------------|
| 改了某个 `skills/<id>/` 的 scripts、config 或 SKILL.md | **只给这一个** skill：`1.0` → `1.1` → `1.2`。同一 commit 改根 README 本表 |
| 没改到的 skill | **禁止**动它的 version |
| 破坏触发词 / 输入输出契约 | 仍 +0.1，并在该 skill README 写一句「同事要改口令」 |

发布：测绿后 `git push gitee main`。不要为版本号单独开夹。

## 技能表（全部 1.0）

| 同事怎么说 | id | version | 干什么 |
|------------|----|---------|--------|
| 跑本周应收 | receivables-merge | 1.0 | 分年表合成应收 all（含催收参考、原生透视） |
| 把 all 拆给各销售 | split-by-sales | 1.0 | 一人一份 |
| 本周抽谁 | compliance-spot-check | 1.0 | 合规抽查建议名单 |
| 跑昨天的核销 | ar-hexiao-daily | 1.0 | 出纳核销日清（极速业务脚本） |
| 出合并报表 | consolidated-statements | 1.0 | 八主体合并底稿；可从金蝶拉五个账套 |
| 月度损益表 | pl-dept-report | 1.0 | 星辰拼损益表+利润表 |
| 销项/付款/收款入金蝶 | kingdee-posting | 1.0 | 填凭证引入表，人去点引入 |
| 序时账入金蝶 | kingdee-gl-import | 1.0 | 序时账 → 官方引入表 |
| 劳务发票核对 | labor-invoice-check | 1.0 | 有票 / 800 以下 / 无票 |
| 申报表重命名 | withholding-report-rename | 1.0 | 代扣代缴 PDF 改名 |
| 部门费用归集 | dept-expense-alloc | 1.0 | 用友按人拆部门 |
| 九点下单 | order-daily-summary | 1.0 | 下单数据（万元） |
| 追觅进度对比 | dreame-ar-progress-diff | 1.0 | 多版 list diff |
| 琪哥发票入金蝶 | qige-invoice-to-kingdee | 1.0 | 兼容入口，进销项 |
| 更新财务skills | update-finance-skills | 1.0 | 从 Gitee 拉白名单 |
| 帮我理清需求 | task-clarifier | 1.0 | 含糊时先问清 |
| Excel / Word / PPT / PDF | xlsx / docx / pptx / pdf | 1.0 | 零散改文档 |

`project-detail-to-ledger` 在仓内，version 1.0，未进更新白名单。

## 仓库

| | |
|--|--|
| **发布 / 同事拉** | Gitee https://gitee.com/Lee157/finance-skills **main** |
| 形态 | 一个 monorepo，`skills/<id>/` |
| 开发交付 | 测绿 → `git push gitee main` |

同事说「更新财务skills」→ 只覆盖白名单，保留本机 `config/`（核销业务规则跟仓库）。

## 开发

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m pytest -q
```

改 skill 后：本地测绿 → **`git push gitee main`**。默认不再打 zip。

真表、口令、客户明细不要进这个仓库。
