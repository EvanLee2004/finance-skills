# 财务部 skills

一个个 skill，专门解决财务部各种**复杂、反复**的活——合并、拆分、对账、核销、报表……

每个 skill 都是「**自然语言驱动 + agent 照流程用 Python 干活**」：财务同事说人话，找文件 / 跑 / 复核全归 agent，不用改文件名、摆文件夹。

**当前源码共 14 个技能**（与官方分发包 `财务技能包_v1.0.16` 一致）。装到同事机器请用飞书下发的 zip + 使用手册；本仓是源码与版本真相源。

## 技能清单（三层：业务技能 + 行为/环境 + 通用基座）

技能分三层：

- **业务技能**：封装财务部某个具体的活（口径、归属规则写进 `config/`，结果可逐行复现）
- **行为 / 环境**：理清需求、装依赖——不碰业务口径，但所有业务技能都用得上
- **通用基座**：处理四类文档（Excel/PDF/Word/PPT）的底层能力，给业务技能"打下手"、也兜住够不上独立技能的零散文档活

### 业务技能（财务专有，config 驱动、可复现）· 8 个

| skill | 解决什么 | 状态 |
|-------|----------|------|
| [receivables-merge](skills/receivables-merge/) | **应收账款合并**：合并分年表、算账龄、按上一版回填标注、按维护表做销售归属、删已回款行、结转老坏账、出透视汇总 | ✅ 真实数据验证 · 回归通过 · 已入包 |
| [split-by-sales](skills/split-by-sales/) | **按销售拆分**：把应收 all 按销售人员拆成一人一份带下拉框 Excel（账龄降序、坏账桶忽略、GM 单独 sheet、对账）——接在合并之后 | ✅ 回归通过 · 链路通 · 已入包 |
| [labor-invoice-check](skills/labor-invoice-check/) | **劳务发票核对**：待支付清单(国内个人)×发票台账，按身份证号求和多张发票、实习生/外国人豁免、≤800 放行、>800 缺票/未开票标黄催票 → 主核对表+不付名单+可付名单 | ✅ 真实数据验证 · 回归通过 · 已入包 |
| [withholding-report-rename](skills/withholding-report-rename/) | **代扣代缴申报表重命名**：一批「代扣代缴、代收代缴税款报告表」PDF 批量改名成 `{纳税人名称}{金额合计}.pdf`；默认 copy 不动原件、出对照表，抽不到的进待人工 | ✅ 回归 5/5 · opencode 实测 · 已入包 |
| [compliance-spot-check](skills/compliance-spot-check/) | **合规文件抽查**：吃应收 all（+可选抽查历史）→ 本周建议名单（营销人员｜客户｜交付月份）；未反馈优先、已反馈月份跳过、覆盖在职；只推荐不自动发邮件 | ✅ 合成回归通过 · 已入包 · 待真实 all 试用 |
| [dreame-ar-progress-diff](skills/dreame-ar-progress-diff/) | **追觅应收进度对比**：多版追觅 list「应收进度」按人名对齐、期间并集，出值/底色/列结构 diff（预计付款忽略公式）+ 结论摘要 | ✅ 回归 47/47 · 真实金标 · **已入包 v1.0.15+** |
| [dept-expense-alloc](skills/dept-expense-alloc/) | **部门费用归集分摊（月度）**：用友余额+收入底稿+人员归属+按人费用 → 部门科目余额表+利润表，主体合计=部门合计核对≈0 | ✅ v1.0.0 可交付 · **已入包** · 待真实月份试用 |
| [ar-hexiao-daily](skills/ar-hexiao-daily/) | **应收核销日清**：出纳每日 T-1 核销判定 + 今日工作清单 + 挂账重扫（人在环；第一版只判不写用户原表、**永不写智云**） | ✅ 历史回放 135/135 · **已入包 v1.0.16** · 待工位真 T-1 验收 |

> 链路示意：`receivables-merge` → `split-by-sales`（旁路 `compliance-spot-check`）；出纳核销独立走 `ar-hexiao-daily`。  
> **规划中（未建 skill）**：销售反馈汇总 等。  
> **已下线 / 迁出**：`payroll-info-match`、`insurance-fund-merge`（不做）；`bank-income-extract` 已改独立 Windows exe（日记账挑收入），不再随本包维护。

### 行为 / 环境 · 2 个

| skill | 解决什么 | 状态 |
|-------|----------|------|
| [task-clarifier](skills/task-clarifier/) | **理清需求**：需求含糊时先用带选项的选择题问清「要干啥 / 文件在哪 / 口径」，再动手——绝不猜 | ✅ 已入包（改编自 trailofbits/skills，CC BY-SA 4.0） |
| [env-doctor](skills/env-doctor/) | **环境管家**：缺 Python 库 / LibreOffice·poppler·tesseract / Python 版本太老时，查《依赖与安装清单》按**国内镜像优先**装齐再重试。纯提示词、不碰业务数据 | ✅ 清单覆盖全包技能 · 清华镜像实装验证 · 已入包 |

### 通用基座（处理四类文档；改自 Anthropic 官方 office skills）· 4 个

| skill | 解决什么 | 状态 |
|-------|----------|------|
| [xlsx](skills/xlsx/) | **Excel 表格**：读写/公式/清洗/出成品 .xlsx；`recalc.py` 校验零公式错误 | ✅ 已入库 · 已入包 |
| [pdf](skills/pdf/) | **PDF 处理**：读文抽表、合并拆分、旋转水印、填表、加解密、转图、OCR | ✅ 已入库 · 已入包 |
| [docx](skills/docx/) | **Word 文档**：创建/编辑/解析 .docx，批注修订、插图、提正文 | ✅ 已入库 · 已入包 |
| [pptx](skills/pptx/) | **PPT 演示文稿**：做幻灯片、改模板、抽正文、合并拆分 deck | ✅ 已入库 · 已入包 |

**合计：8 业务 + 2 行为/环境 + 4 基座 = 14。**

> **环境依赖（部署到同事机器时注意）**：① 四类通用基座的"校验"脚本 `office/validate.py` 用了 `match` 语法，**需 Python ≥3.10**（3.9 会报 SyntaxError）——核心读写不受影响，仅可选校验步骤受限。② `xlsx/recalc.py`、`pptx/thumbnail.py`、`docx/accept_changes.py` 依赖 **LibreOffice（soffice）**重算/转图/接受修订；没装 LibreOffice 时这几个功能降级，openpyxl/python-docx/pypdf 的基本读写仍正常。  
> **以上环境问题统一交给 `env-doctor` 处理**——任何技能缺库/缺工具，agent 查它的清单按国内镜像装齐再重试。  
> 应收核销日清另依赖 **`xlrd`**（老式 `.xls` 日记账）；安装提示词见使用手册。

## 每个 skill 长什么样（标准）

见 [docs/技能标准规范.md](docs/技能标准规范.md)：标准四件套 `SKILL.md + scripts/ + references/ + config/`；核心原则**人说人话、脏活归 agent**；会变的东西（认列、归属规则）外置成可维护的配置表，规则变改表不改码。

## 新做涉及 Excel 的技能：怎么用 xlsx 这个基座（架构约定）

结论：**xlsx 当"工具箱 + 规范"，不当"代码母本去 fork"。** 三种姿势按场景选——

1. **够不上独立技能的零散 Excel 活**（临时加列、做张小表、洗个乱表）→ **不必新建技能**，直接让 agent 用 xlsx 这个通用基座干。
2. **新的业务 Excel 技能**（如费用归集、回单台账）→ **照四件套新建独立技能**（自己的 `scripts/` 用 openpyxl/pandas 直接写、业务规则进自己的 `config/`），**不要把 xlsx 的代码 fork 进来**——xlsx 自带几百个 XML schema，业务技能用不上，fork 只会臃肿、还得跟着升级。业务技能可**调用** xlsx 的 `recalc.py` 做"零公式错误"自检、按 xlsx 的配色/数字格式规范出成品，但**依赖关系是"调用/参照"，不是"继承代码"**。
3. **要深改 Excel 底层 XML**（普通 openpyxl 干不了的，如复杂图表、特殊样式）→ 借 xlsx 的 `office/unpack.py`、`pack.py` 解包改包。

> 一句话：业务技能保持**独立、config 驱动、可复现**（这是它的价值）；xlsx 提供**通用能力 + 出品规范 + 自检工具**。两层解耦——业务规则变了改业务技能的 config，文档处理能力升级了升基座，互不牵连。

## 数据安全

真实财务数据（源台账 / 回填源 / 成品 / 核销运行工作区等批量数据）**不进仓库**（见 `.gitignore`）；本库**私有**。各 skill 的维护表（如销售变化表）作为可长期维护的配置保留在库内。分发包 zip 只含技能源码，不含 `工作区/` 运行产物。

## 分发与版本

| 项 | 说明 |
|----|------|
| 源码仓 | 本仓库 `skills/`（GitHub 私有） |
| 同事安装 | 飞书下发的 `财务技能包_vX.Y.Z.zip` +《财务技能使用手册》 |
| 当前对齐 | **v1.0.16 / 14 技能**（2026-07-23 起含 `ar-hexiao-daily`） |

改 skill 后：本地测绿 → commit/push 本仓 → 按 `发布/` 规程重打 zip 并升手册（zip 与手册不在本仓内）。
