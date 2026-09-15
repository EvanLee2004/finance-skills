# 财务部 skills

一件活一个夹，都在 `skills/`。

开口对照：[技能表](#技能表)。源在 Gitee `Lee157/finance-skills` 的 `main`。

## 版本

每个 skill 的 `SKILL.md` 里有 `version`。这次起业务 skill 是 **1.1**。

改某个 `skills/<id>/` 的脚本、config 或 SKILL.md，只给那一个 +0.1，并改本表。没改的不要动版本。测绿后 `git push gitee main`。

## 技能表

按使用人排。名称后括号是谁用。

### 亮晶

| 名称 | id | version | 简介 |
|------|----|---------|------|
| 应收账款合并（亮晶） | receivables-merge | 1.1 | 分年表合成应收 all，带催收参考和透视 |
| 应收按销售拆分（亮晶） | split-by-sales | 1.1 | 一张 all 拆成一人一份 |
| 合规文件抽查（亮晶） | compliance-spot-check | 1.1 | 建议本周抽查名单 |
| 劳务发票核对（亮晶） | labor-invoice-check | 1.1 | 有票 / 800 以下 / 无票 |
| 九点下单统计（亮晶） | order-daily-summary | 1.1 | 下单数据（万元） |
| 追觅应收进度对比（亮晶） | dreame-ar-progress-diff | 1.1 | 两版追觅 list 的进度差 |

### 明妹

| 名称 | id | version | 简介 |
|------|----|---------|------|
| 应收核销日清（明妹） | ar-hexiao-daily | 1.1 | 智云取数、判定、写盈亏和流转副本 |

### 斯佳

| 名称 | id | version | 简介 |
|------|----|---------|------|
| 金蝶入账（斯佳） | kingdee-posting | 1.1 | 销项 / 付款 / 收款填引入表，人去点引入 |
| 月度损益表（斯佳） | pl-dept-report | 1.1 | 星辰多账套拼损益表和利润表 |
| 序时账入金蝶（斯佳） | kingdee-gl-import | 1.1 | 序时账转官方引入表 |
| 部门费用归集分摊（斯佳） | dept-expense-alloc | 1.1 | 用友按人拆部门费用 |
| 合并报表（斯佳） | consolidated-statements | 1.1 | 八主体资负利润现金流底稿 |
| 代扣代缴申报表重命名（斯佳） | withholding-report-rename | 1.1 | 申报 PDF 按公司名+金额改名 |

### 通用

| 名称 | id | version | 简介 |
|------|----|---------|------|
| Excel（通用） | xlsx | 1.1 | 读写和整理表格 |
| Word（通用） | docx | 1.1 | 读写 Word |
| PPT（通用） | pptx | 1.1 | 读写演示文稿 |
| PDF（通用） | pdf | 1.1 | 读、拆、合并 PDF |

已下线：`qige-invoice-to-kingdee`（销项已在金蝶入账）、`task-clarifier`、`update-finance-skills`、`env-doctor`。

## 仓库

Gitee：https://gitee.com/Lee157/finance-skills · 分支 `main`

改完测绿后 `git push gitee main`。真表、口令、客户明细不进仓。
