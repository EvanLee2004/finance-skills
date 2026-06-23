# 财务部 skills

一个个 skill，专门解决财务部各种**复杂、反复**的活——合并、拆分、对账、回单、报表……

每个 skill 都是「**自然语言驱动 + agent 照流程用 Python 干活**」：财务同事说人话，找文件 / 跑 / 复核全归 agent，不用改文件名、摆文件夹。

## 技能清单

| skill | 解决什么 | 状态 |
|-------|----------|------|
| [receivables-merge](skills/receivables-merge/) | **应收账款合并**：合并分年表、算账龄、按上一版回填标注、按维护表做销售归属、删已回款行、结转老坏账、出透视汇总 | ✅ 真实数据验证 · grok 9/10 |
| [split-by-sales](skills/split-by-sales/) | **按销售拆分**：把应收 all 按销售人员拆成一人一份带下拉框 Excel（账龄降序排好、坏账桶忽略、GM单独成sheet、对账）——接在合并之后 | ✅ 回归12/12 · 链路通 |
| [docx](skills/docx/) | **Word 文档**：创建/编辑/解析 .docx | ✅ 已入库 |
| [pptx](skills/pptx/) | **PPT 演示文稿**：创建/编辑 .pptx | ✅ 已入库 |
| [xlsx](skills/xlsx/) | **Excel 表格**：读写/公式/清洗 .xlsx | ✅ 已入库 |
| [pdf](skills/pdf/) | **PDF 处理**：合并/拆分/填表/OCR 等 | ✅ 已入库 |
| …更多 | 回单查询 / 销售反馈汇总 / 费用归集 等 | 规划中 |

## 每个 skill 长什么样（标准）

见 [docs/技能标准规范.md](docs/技能标准规范.md)：标准四件套 `SKILL.md + scripts/ + references/ + config/`；核心原则**人说人话、脏活归 agent**；会变的东西（认列、归属规则）外置成可维护的配置表，规则变改表不改码。

## 数据安全

真实财务数据（源台账 / 回填源 / 成品等批量数据）**不进仓库**（见 `.gitignore`）；本库**私有**。各 skill 的维护表（如销售变化表）作为可长期维护的配置保留在库内。
