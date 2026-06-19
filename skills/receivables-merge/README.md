# receivables-merge · 应收账款合并技能

财务部 Agent 平台的样板技能（第一个符合《技能标准规范》的 skill）：把每周手工做的"应收账款合并"自动化——合并多年份分表、算账龄、回填销售标注、按维护表做销售归属、删已回款行、出透视汇总。

## 结构（标准四件套）

```
receivables-merge/
├── SKILL.md            # agent 行为指南（触发词 + 怎么找文件/跑/复核）
├── scripts/merge.py    # S1–S8 确定性实现（CLI）
├── references/
│   ├── SOP_应收账款合并.md   # 完整流程与口径（视频+成品实证）
│   └── 思路_skills化.md       # 设计思路
├── config/                   # 「活」的配置：会变/要判断的都在这，改表不改码
│   ├── 列名别名.json          # S1 认列规则
│   ├── 销售归属维护表.md       # 销售归属（离职→接手、客户→重分配）；人维护、agent 每次确认
│   └── 业务规则.md            # 特殊批次 / 跳过的 sheet / 业务规则说明；agent 跑前先读
├── tests/test_robustness.py  # 稳定性回归测试（缺文件/坏输入/确定性）
├── evals/evals.json          # skill-creator 测试 prompt
└── 工作区/{input,output}/    # 运行时数据（.gitignore）
```

## 用法

```bash
# 1) 先识别三个输入（按内容认，不靠文件名）
python3 scripts/merge.py --inspect --input-dir <放文件的目录>

# 2) 跑合并
python3 scripts/merge.py \
  --source 源台账.xlsx --ref 回填源.xlsx \
  --rules config/销售归属维护表.md \
  --out 工作区/output/应收all.xlsx
# --ref 省→跳回填；--rules 省→跳归属；--base-month 202606 补往月
```

依赖：`python3 -m pip install pandas openpyxl`

## 验收状态（2026-06-18）

用真实历史数据（`2026.6.4日应收` → 对照成品 `2026.6.4应收all`）验证：
- **所有业务规则已确认正确**：账龄口径、新智云单号补位、删应收=0行、**复合名只对高美杰**、**客户重分配全局优先**、销售归属三路分流（梁玲玲-高美杰/于占国-高美杰/高美杰1）。
- 销售人员分布 10/20 完全吻合，其余为 ±1~12 的小差。
- 残留 ~0.8%(约25行) 差额已定位为**数据来源对不齐**（手上源表比成品少 ~26 行高美杰，缺当周真实回填源），**非逻辑错误**。
- 待补：拿到当周真实回填源后做一次完整复现验证。

## 数据安全

源台账/回填源/成品/真实维护表均含敏感财务数据，**一律 .gitignore、不进仓库**。仓库只含代码 + 文档 + 维护表空模板。
