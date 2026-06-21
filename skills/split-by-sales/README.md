# split-by-sales · 应收按销售人员拆分技能

财务部 skills 的第 2 个。把 **receivables-merge（合并）产出的「应收 all」** 按销售人员拆成一人一份带下拉框的 Excel（亮晶姐模板）。是「合并 → **拆分** → 发给销售」链条的中间一棒。

## 结构（标准四件套）
```
split-by-sales/
├── SKILL.md            # agent 行为指南（触发词 + 人在环 + 怎么跑/收尾）
├── scripts/split.py    # 拆分实现（CLI）：分组/模板/下拉/对账
├── config/拆分规则.md   # 会变的：坏账桶忽略名单 + GM单独成sheet的接手人（人维护）
├── tests/test_robustness.py  # 稳定性回归（合成假数据，12/12）
└── 工作区/{input,output}/
```

## 用法
```bash
python3 scripts/split.py --input <应收all.xlsx绝对路径> --date 0604
# --out-dir 可省（默认输出到 all 同目录的 拆分_<日期>/）；--date 不给用当天
```
依赖：`python3 -m pip install openpyxl`

## 口径（项目五已验收）
- 坏账桶（高美杰1）整行忽略、不分给任何人。
- 名字「X-高美杰」归到 X；其中 GM_OWNERS（于占国）的高美杰行单独进「GM订单」sheet，其余（梁玲玲）并入主表。
- 销售人员为空 → `_销售人员为空_请人工处理.xlsx`。
- 跑完**对账**：分出 + 空名 + 忽略 == 输入行数。

## 验收 / 链路
- 真实 all（`2026.6.4应收all`，3243 行）→ 15 位销售、于占国含 GM订单 74、忽略 10、对账 3233+0+10=3243 ✓。
- 链路：`receivables-merge` 产出 all → 本技能拆 → 对账对得上（实测通过）。

## 数据安全
真实财务数据不进仓库；只含代码 + 文档 + 规则配置。
