# payroll-info-match（工资表清洗与信息匹配）

给同事用 opencode 说人话跑；也可命令行直接跑。

## 干嘛的
两张表 → 一张干净工资底表：
- 输入：**工资原始表**（明妹给的未处理版：多 sheet，姓名/身份证号码/基本工资~实发工资，全是公式）+ **员工信息表（权威）**（按月 sheet：姓名/地区/部门/岗位/身份证号/电话）。
- 处理：认工资 sheet → 定人员行范围(排除合计/参考数据) → 删无表头空列 → 公式转数值两位小数 → 合计反算校验 → 按姓名匹配当月花名册(重名/未匹配兜底) → 出底表。
- 输出：`主表` + `待人工核实`(标黄) + `匹配成功` + `运行报告`。

## 命令行
```bash
# 认文件 + 看建议月份（不跑）
python3 scripts/clean_match.py --inspect --input-dir <目录>
# 跑清洗匹配（--month 必须显式传，不接受自动猜测）
python3 scripts/clean_match.py --payroll 工资原始表.xlsx --employee 员工信息表.xlsx --month 202606 --out 结果.xlsx
```

## 跑测试
```bash
python3 tests/test_robustness.py   # 合成金标 + 真实数据冒烟，应全绿
```

## 改规则（不改代码）
- 合计容差/小数位/合计行关键字/空行停止阈值 → `config/业务规则.md`
- 列名别名 + 所属公司映射 → `config/列名别名.json`

## 结构
```
payroll-info-match/
├── SKILL.md                给 agent 的操作说明（人在环）
├── scripts/clean_match.py  清洗+匹配核心（确定性、可复现）
├── config/                 业务规则.md（活参数）+ 列名别名.json
├── tests/                  test_robustness.py
└── 工作区/input,output      默认收发文件处（不进 git）
```

## ⚠ 数据红线
工资表含身份证号/工资金额 PII → 只在本地跑、结果别外传。
