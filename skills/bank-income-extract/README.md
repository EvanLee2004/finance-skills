# bank-income-extract（银行日记账收入提取汇总）

把一份银行日记账 Excel 里**每个 sheet（每个银行/现金账户）**的"收入"记录筛出来汇总成一张表，
并把可疑行挑进『待人工核对』。给出纳李明妹用（应收核销自动化 · 模块①）。

## 直接命令行跑（不经 opencode 时）
```bash
python3 scripts/extract_income.py 日记账.xlsx
python3 scripts/extract_income.py 日记账.xlsx -o 收入汇总.xlsx
python3 scripts/extract_income.py 日记账.xlsx --config 别处/识别规则.md
```
依赖 `openpyxl`。纯本地、不联网、不调 AI、不花额度。

## 输出
- **收入汇总** sheet：`渠道（银行/账户）| 日期 | 到账客户名称 | 金额 | 币种`，按 sheet 顺序排；
  缺客户名的行标黄。
- **待人工核对** sheet（有可疑行才建）：疑似漏标收入、收入行缺客户名。

## 识别规则
一行算收入 = 「类型」列含"收入/到账" **且**「借方(增加)」有金额（专门过滤"类型=收入但金额空"
的预填行）。客户名取「摘要」、金额取「借方」、渠道取 sheet 名。规则全在 `config/识别规则.md`，
改表不改码。

## 目录
```
bank-income-extract/
├── SKILL.md              技能说明（agent 行为指引）
├── README.md             本文件
├── config/识别规则.md     ★活表：收入词/列名/跳过sheet/外币，改表不改码
├── scripts/extract_income.py
└── tests/test_extract.py  合成回归（14 断言）
```

## 测试
```bash
python3 tests/test_extract.py    # 期望 All passed ✅
```

## 边界
- 找不到表头或缺关键列的 sheet 会安全跳过并在小结说明，不静默丢数据。
- 不同币种不跨币种加总（分币种报小计）。
- "疑似漏标"只提示、不自动进汇总（防止把非收入误当收入）。
