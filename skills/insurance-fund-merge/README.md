# insurance-fund-merge（社保公积金合并与在职台账）

给同事用 opencode 说人话跑；也可命令行直接跑。

⚠ **这是"先做出来试"的版本，没走完需求确认书流程**——详见 `SKILL.md` 开头和 `config/业务规则.md` 第一节。

## 干嘛的
工资表 + 社保/公积金台账（或湖南这类合并的"五险一金台账"）→ 跨月累计的薪酬明细：
- 输入：**当月工资表**（含姓名+基本工资即可，不限定来源）+ **该主体社保台账/公积金台账**（分开两个文件，或湖南这类合并一个文件）。
- 处理：识别单位缴纳部分金额列（不认位置只认表头关键字，兼容2行合并表头/上海"企业部分"/湖南组合台账等多种真实结构）→ 按姓名匹配进工资表 → 组织架构名册只增不删 → 编外人员按名单单独拆出 → 跨月累计（幂等，重跑不翻倍）→ 按"发工资=在职"算在职月份/人员状态 → 生成多月对比宽表。
- 输出：`薪酬明细` + `组织架构新` + `编外人员` + `待人工核实` + `薪酬汇总` + `运行报告`。

## 命令行
```bash
# 认文件（不跑）——新主体第一次用务必先跑这个
python3 scripts/merge_insurance.py --inspect --input-dir <目录>

# 独立社保+公积金两个文件的主体（甲骨易/上海/文化传媒/园创园/语丰泰达）
python3 scripts/merge_insurance.py --entity 甲骨易 --month 202606 \
  --payroll <工资表.xlsx> --insurance <社保台账.xlsx> --fund <公积金台账.xlsx> \
  [--master <累计薪酬台账.xlsx>] [--out <输出.xlsx>]

# 社保公积金合并一个文件的主体（湖南分公司/湖南子公司）
python3 scripts/merge_insurance.py --entity 湖南分公司 --month 202606 \
  --payroll <工资表.xlsx> --combined <五险一金台账.xlsx> \
  [--master <累计薪酬台账.xlsx>] [--out <输出.xlsx>]
```

## 跑测试
```bash
python3 tests/test_robustness.py   # 4种真实表头结构的合成金标 + 真实数据冒烟，应全绿
```

## 改规则（不改代码）
- 编外人员名单 → `config/编外人员名单.json`
- 在职判定口径/已知不做的范围/开工方式说明 → `config/业务规则.md`
- 已验证过真实数据的主体清单 → `config/主体配置.json`

## 结构
```
insurance-fund-merge/
├── SKILL.md                    给 agent 的操作说明（人在环）
├── scripts/merge_insurance.py  表头识别+合并+跨月累计+在职判定核心（确定性、可复现）
├── config/                     业务规则.md + 编外人员名单.json + 主体配置.json（活参数）
├── tests/                      test_robustness.py
└── 工作区/input,output          默认收发文件处（不进 git）
```

## ⚠ 数据红线
含身份证号/工资/社保金额 PII → 只在本地跑、结果别外传。
