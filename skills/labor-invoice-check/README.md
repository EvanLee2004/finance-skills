# labor-invoice-check（劳务发票核对）

个人劳务付款前的发票核对（第一阶段）。给同事用 opencode 说人话跑；也可命令行直接跑。

## 干嘛的
两张表 → 一张核对结果：
- 输入：**待支付清单**（国内个人：姓名/应付金额/备注/身份证号）+ **发票统计台账**（销售方名称/纳税人识别号/合计金额）。
- 处理：按**身份证号**把同一人多张发票的合计金额**求和**，与应付金额比；实习生/外国人豁免、≤800放行、>800缺票或未开票标黄。
- 输出：`主核对表`（标黄缺票/未开票行）+ `不付名单(催票)` + `可付名单` + `运行报告`。

## 命令行
```bash
# 认文件（不跑）
python3 scripts/check.py --inspect --input-dir <目录>
# 跑核对
python3 scripts/check.py --list 待支付清单.xlsx --invoice 发票台账.xlsx --out 结果.xlsx
# 省略 --list/--invoice 则从 工作区/input 自动认；省略 --out 落清单同目录
```

## 跑测试
```bash
python3 tests/test_robustness.py   # 合成金标 + 真实数据冒烟，应全绿
```

## 改规则（不改代码）
- 门槛/容差/实习生关键字/外国人识别 → `config/业务规则.md`
- 列名认别名 → `config/列名别名.json`

## 结构
```
labor-invoice-check/
├── SKILL.md            给 agent 的操作说明（人在环）
├── scripts/check.py    核对核心（确定性、可复现）
├── config/             业务规则.md（活参数）+ 列名别名.json
├── tests/              test_robustness.py
└── 工作区/input,output  默认收发文件处（不进 git）
```

## ⚠ 数据红线
清单含身份证/账号/电话 PII → 只在本地跑、结果别外传；身份证号仅作匹配键，输出默认不展示付款列。
