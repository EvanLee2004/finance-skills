---
name: pl-dept-report
description: >-
  每月从金蝶云星辰取出有账套的主体，拼成斯佳「损益表 + 利润表」Excel（月度损益表 / 科目余额表 / 部门科目余额表）。
  当用户说「月度损益表 / 科目余额表 / 出本月损益表 / 损益表利润表 / 部门科目余额表 / 跑斯佳那张表 / 金蝶损益表」时用本技能。
  总部只读 API 能打就打；另外 4 本星辰账吃引出 xlsx；山东/四川/济南无账套留空。公式自写核对，禁止改数凑平。
  用友按人拆部门费用走 dept-expense-alloc，不要和本技能抢。
---

# 月度损益表 / 科目余额表

同事说人话出一张 Excel：**只有** `损益表` 和 `利润表` 两张 sheet。金额发生额由 `scripts/convert.py` 写；合计、核对、父行、利润勾稽是 Excel 公式。

## 启动时说

- **推荐**：月度损益表
- 也可：科目余额表 / 出本月损益表 / 损益表利润表 / 部门科目余额表 / 跑斯佳那张表 / 金蝶损益表

## 0. 红线

1. 禁止自写 Python 读她的表、禁止心算、禁止改发生额凑平。只准调本技能 `scripts/`。
2. 对话只报：期间、有源/缺源账套短名、未映射部门个数、核对非 0 的科目编码个数、产物路径。禁止金额、口令、Client Secret。
3. 密钥只读本机 `~/.config/finance/kingdee.local.json` 与 `kingdee-pl.local.json`。文件已在就直接用，不要问她填应用号，不要去开放平台点重置密钥 / 解除授权 / 购买。
4. 禁止点金蝶引入 / 审核 / 过账；禁止解绑总部。
5. 山东分公司、四川分公司、济南子公司没有星辰账套：单元格留空，不是 0、不是总部拷贝。
6. 未映射部门金额不进任何部门列，只进运行报告。
7. 真实引出 xlsx 不进 git。Playwright 不是本技能依赖，不要写进同事 README。

## 1. 认期间和引出

没说月份 → 上一个已过完的公历月（9 月 1 日出 8 月）。她给了文件夹就 `--input-dir`；没给就扫当前工作区 / 桌面。引出按 **sheet / 表头** 认（科目余额表、核算项目余额表、利润表），不靠文件名。

```bash
python3 "<本skill目录>/scripts/inspect_inputs.py" --input-dir <绝对目录>
python3 "<本skill目录>/scripts/convert.py" --period YYYYMM --input-dir <绝对目录> --out <目录>/月度损益表_YYYYMM.xlsx
```

系统 python 缺库时脚本会切到仓内 `.venv`。不要对系统 Python `pip install`。缺环境说「配下环境」，转 env-doctor。

本机已有密钥时 convert **不要**加 `--no-api`（测试才加）。总部走只读 API；文化/上海/湖南分/湖南子用本机浏览器引出：

```bash
python3 "<本skill目录>/scripts/xingchen_export.py" --period YYYYMM --out-dir <引出目录>
```

Playwright 只装在本机取数用，不要写进同事 README、不要 `playwright install` 当必装。引出成功后再跑 convert。没有引出的星辰账列空着。

## 2. 收尾（照抄）

> 月度损益表已出，期间 YYYYMM。有源账套：…；缺源：…。未映射部门 N 个。核对非 0 的科目编码 M 个。表在 \<路径\>，运行报告在同目录。请打开看两个 sheet，不要改数凑平。

有核对非 0：只报科目编码个数，停下让她看报告。不要问「要不要改数」。

## 3. 会变的在哪

| 文件 | 改什么 |
|------|--------|
| `config/账套清单.json` | 8 家顺序、5 本星辰账 |
| `config/部门名映射.json` | 金蝶档案名 → Excel 列（大客户→KA） |
| `config/版式.json` | 科目树、列、冻结（无金额） |
| `config/引出列名.json` | 引出表头别名 |
| 本机 `kingdee.local.json` | 连接器；不进仓 |
| 本机 `kingdee-pl.local.json` | 本技能账套表；不进仓 |

## 防漏

| 用户说法 | 用哪个 |
|----------|--------|
| 月度损益表 / 科目余额表 / 跑斯佳那张表（金蝶） | **本技能** |
| 部门费用归集 / 用友按人拆 | `dept-expense-alloc` |
| 销项/付款/收款入金蝶 | `kingdee-posting` |
