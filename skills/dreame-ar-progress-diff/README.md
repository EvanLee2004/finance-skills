# dreame-ar-progress-diff（追觅应收进度对比）

给亮晶用的追觅客户「应收进度」多版本对比技能。人话触发；也可命令行直跑。

## 干嘛的

多份格式相近的追觅 list Excel（含「应收进度」sheet）→ 一份对比报告：

| 输入 | 处理 | 输出 |
|------|------|------|
| 2～N 个时间点的进度表 | 人名智能对齐、期间并集、值+底色 diff、公式当备注忽略 | `结论摘要/列结构/值变化/颜色变化/预计付款/明细/运行报告` |

典型场景：0703 vs 0709 vs 0713，看谁催了、谁回款了、谁收了 PO、有没有新月账。

## 命令行

```bash
# 认文件（不跑）
python3 scripts/compare.py --inspect --input-dir <目录>

# 明确点名（推荐，旧→新）
python3 scripts/compare.py --files 旧.xlsx 中.xlsx 新.xlsx --out 报告.xlsx

# 目录自动认 + 按文件名日期排序
python3 scripts/compare.py --input-dir <目录>
```

依赖：`python3 -m pip install openpyxl`（无需 pandas）。

## 跑测试

```bash
python3 tests/test_robustness.py   # 合成金标 + 可选真实桌面样本；应全绿
```

## 改规则（不改代码）

| 文件 | 改什么 |
|------|--------|
| `config/业务规则.md` | sheet 名、表头行、停止词、是否忽略公式 |
| `config/子列识别.json` | PO/发票/预计付款等表头别名 |
| `config/颜色图例.json` | 底色 RGB → 中文名 |

## 结构（标准四件套）

```
dreame-ar-progress-diff/
├── SKILL.md                 # agent 行为指南
├── scripts/compare.py       # 确定性对比 CLI
├── config/                  # 活配置
├── references/SOP_….md      # 完整口径
├── tests/test_robustness.py
├── evals/evals.json
└── 工作区/{input,output}/   # 默认收发（不进 git）
```

## 设计原则（稳定好用）

- **流程写清在 SKILL.md**：AI 冷启动能懂「表是啥、要比啥、怎么收尾」，不用亮晶从头讲。  
- **会变的不写死**：列名/人数/月份/具体金额/文件名 → 当次解析；表头别名/颜色/停止词 → `config/`。  
- **钉死的只留稳定性**：期间并集、人名对齐、公式忽略、源只读、报告结构 → 脚本。  
- 历史「值变化 29」等数字只作回归金标，不是每次业务 KPI。

## 验收状态（2026-07-15）

- 合成回归：全绿  
- 真实样本冒烟：结构与口径符合 SOP（具体条数随当次数据变）

## 数据红线

真实 list / 报告含客户与金额 → 不进 git；只在本地跑。
