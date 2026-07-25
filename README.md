# 财务部 skills

一个个 skill，专门解决财务部各种**复杂、反复**的活——合并、拆分、对账、核销、报表……

## 许可证

本仓库默认采用 **MIT License**（见根目录 `LICENSE`）。

- **为什么 MIT**：公开技能仓需要「clone 即可用、可改、可再分发」；MIT 是最常见的宽松许可证之一，与多数 Python 工具链兼容，条款短、无专利额外条款。
- **对比过 Apache-2.0**：Apache-2.0 带明确专利授权与 NOTICE 要求，更重；本仓以财务脚本技能为主、无多贡献方专利博弈场景，MIT 更轻。
- **待拍板**：最终许可证选型属产品决策，默认按 MIT 落地，若需 Apache-2.0 / 其他请明昊确认后改。

## 开发 / 测试

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m pytest -q         # 仓根；配置见 pytest.ini（同名测试文件已改为唯一 basename）
```


CI：`.github/workflows/pytest.yml` 在 push/PR 到 `main` 时跑同一套 `pytest`。

每个 skill 都是「**自然语言驱动 + agent 照流程用 Python 干活**」：财务同事说人话，找文件 / 跑 / 复核全归 agent，不用改文件名、摆文件夹。

**当前源码共 15 个技能**。**本仓是唯一源码与版本真相源**（monorepo，不是一 skill 一仓）。

> ### ⭐ 更新去哪？就这一个仓库 · 2026-07-25 起 git 即分发
>
> | | |
> |--|--|
> | **GitHub** | https://github.com/EvanLee2004/finance-skills |
> | **Gitee（国内优先）** | https://gitee.com/Lee157/finance-skills |
> | **分支** | `main` |
> | **形态** | **一个 monorepo**（全部 skill 在 `skills/` 下） |
> | **开发交付** | 测绿 → `git push origin main`（双端）→ **完成**。**默认不再打 zip、不再飞书发压缩包** |
> | **同事更新** | 对 opencode 说一句 **「更新财务skills」** → 拉云端 main → **只覆盖白名单** → 保留本地 config → **不动自装技能** |
>
> 详细流程、白名单、config 铁律、可复制提示词 → 必读 **[SOURCE.md](./SOURCE.md)**。  
> 装进 opencode 后精简版见 `skills/财务技能包_来源与更新.md`。  
> ⚠️ **git push 了 ≠ 同事本机已更新**；同事要说一次「更新财务skills」。

## 技能清单（三层：业务技能 + 行为/环境 + 通用基座）

技能分三层：

- **业务技能**：封装财务部某个具体的活（口径、归属规则写进 `config/`，结果可逐行复现）
- **行为 / 环境**：理清需求、装依赖——不碰业务口径，但所有业务技能都用得上
- **通用基座**：处理四类文档（Excel/PDF/Word/PPT）的底层能力，给业务技能"打下手"、也兜住够不上独立技能的零散文档活

### 业务技能（财务专有，config 驱动、可复现）· 9 个

| skill | 解决什么 | 状态 |
|-------|----------|------|
| [receivables-merge](skills/receivables-merge/) | **应收账款合并**：合并分年表、算账龄、按上一版回填标注、按维护表做销售归属、删已回款行、结转老坏账、出透视汇总 | ✅ 真实数据验证 · 回归通过 · 已入包 |
| [split-by-sales](skills/split-by-sales/) | **按销售拆分**：把应收 all 按销售人员拆成一人一份带下拉框 Excel（账龄降序、坏账桶忽略、GM 单独 sheet、对账）——接在合并之后 | ✅ 回归通过 · 链路通 · 已入包 |
| [labor-invoice-check](skills/labor-invoice-check/) | **劳务发票核对**：待支付清单(国内个人)×发票台账，按身份证号求和多张发票、实习生/外国人豁免、≤800 放行、>800 缺票/未开票标黄催票 → 主核对表+不付名单+可付名单 | ✅ 真实数据验证 · 回归通过 · 已入包 |
| [withholding-report-rename](skills/withholding-report-rename/) | **代扣代缴申报表重命名**：一批「代扣代缴、代收代缴税款报告表」PDF 批量改名成 `{纳税人名称}{金额合计}.pdf`；默认 copy 不动原件、出对照表，抽不到的进待人工 | ✅ 回归 5/5 · opencode 实测 · 已入包 |
| [compliance-spot-check](skills/compliance-spot-check/) | **合规文件抽查**：吃应收 all（+可选抽查历史）→ 本周建议名单（营销人员｜客户｜交付月份）；未反馈优先、已反馈月份跳过、覆盖在职；只推荐不自动发邮件 | ✅ 合成回归通过 · 已入包 · 待真实 all 试用 |
| [dreame-ar-progress-diff](skills/dreame-ar-progress-diff/) | **追觅应收进度对比**：多版追觅 list「应收进度」按人名对齐、期间并集，出值/底色/列结构 diff（预计付款忽略公式）+ 结论摘要 | ✅ 回归 47/47 · 真实金标 · **已入包 v1.0.15+** |
| [dept-expense-alloc](skills/dept-expense-alloc/) | **部门费用归集分摊（月度）**：用友余额+收入底稿+人员归属+按人费用 → 部门科目余额表+利润表，主体合计=部门合计核对≈0 | ✅ v1.0.0 可交付 · **已入包** · 待真实月份试用 |
| [ar-hexiao-daily](skills/ar-hexiao-daily/) | **应收核销日清**：出纳按核销日取智云数 → SOD 级判定 → 一份《核销日清》→ **她确认** → 写前复核 → 统一写盈亏明细 + 流转安全子集；含跑批台账查漏天（**永不写智云**） | ✅ 测试 186 · **opencode 端到端实测通过**（715 格与她手填逐格一致）· 待工位真 T-1 验收 |
| [order-daily-summary](skills/order-daily-summary/) | **九点下单统计**：登录智云抓下单表 → 组织架构归多语（不含运保）/数据/游戏/其他 →「下单数据(万元)」xlsx | ✅ 单测 24 · **内网真机复测通过（2026-07-24）** |

> 链路示意：`receivables-merge` → `split-by-sales`（旁路 `compliance-spot-check`）；出纳核销独立走 `ar-hexiao-daily`；亮晶下单日报走 `order-daily-summary`。  
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

**合计：9 业务 + 2 行为/环境 + 4 基座 = 15。**

> **环境依赖（部署到同事机器时注意）**：① 四类通用基座的"校验"脚本 `office/validate.py` 用了 `match` 语法，**需 Python ≥3.10**（3.9 会报 SyntaxError）——核心读写不受影响，仅可选校验步骤受限。② `xlsx/recalc.py`、`pptx/thumbnail.py`、`docx/accept_changes.py` 依赖 **LibreOffice（soffice）**重算/转图/接受修订；没装 LibreOffice 时这几个功能降级，openpyxl/python-docx/pypdf 的基本读写仍正常。  
> **以上环境问题统一交给 `env-doctor` 处理**——任何技能缺库/缺工具，agent 查它的清单按国内镜像装齐再重试。  
> 应收核销日清另依赖 **`xlrd`**（老式 `.xls` 日记账）；安装提示词见使用手册。

## 每个 skill 长什么样（标准）

标准四件套 `SKILL.md + scripts/ + references/ + config/` + 一份 `README.md`；
核心原则**人说人话、脏活归 agent**；会变的东西（认列、归属规则、阈值）外置成配置表，**规则变改表不改码**。

| 规范 | 说明 |
|---|---|
| [docs/技能标准规范.md](docs/技能标准规范.md) | 四件套结构、交互模型、config 驱动 |
| **[docs/技能README模板.md](docs/技能README模板.md)** | **每个业务技能的 README 必须讲清三段**：① 这技能干嘛 ② **没有它的时候手工怎么做** ③ **现在人和 AI 怎么配合干**。范本见 `ar-hexiao-daily` / `receivables-merge` |
| **[docs/仓库边界_什么进什么不进.md](docs/仓库边界_什么进什么不进.md)** | **什么该进这个仓库、什么留本地**（真实数据 / 需求收集 / 录音 / 凭据一律不进），含 push 前自检命令 |

> **为什么 README 一定要写"手工原样"**：不写清她原来怎么干，就没人说得清这技能到底省了什么，
> 也没人能判断自动化之后有没有漏掉她原来会做的动作。这是给领导、给接手的人、给未来的 agent 看的。

### 文档边界（一句话）

**技能仓**回答「这技能能干吗、怎么跑、口径是什么」；
**项目文件夹**（`财务部skills/技能/<中文名>/`、各长期项目）回答「需求怎么来的、她原话说了啥、试过哪些方案」。
录音、逐字稿、需求确认 docx、真实 Excel、给同事的成品 Word —— **都不进本仓**。详见上表第三行。

## 新做涉及 Excel 的技能：怎么用 xlsx 这个基座（架构约定）

结论：**xlsx 当"工具箱 + 规范"，不当"代码母本去 fork"。** 三种姿势按场景选——

1. **够不上独立技能的零散 Excel 活**（临时加列、做张小表、洗个乱表）→ **不必新建技能**，直接让 agent 用 xlsx 这个通用基座干。
2. **新的业务 Excel 技能**（如费用归集、回单台账）→ **照四件套新建独立技能**（自己的 `scripts/` 用 openpyxl/pandas 直接写、业务规则进自己的 `config/`），**不要把 xlsx 的代码 fork 进来**——xlsx 自带几百个 XML schema，业务技能用不上，fork 只会臃肿、还得跟着升级。业务技能可**调用** xlsx 的 `recalc.py` 做"零公式错误"自检、按 xlsx 的配色/数字格式规范出成品，但**依赖关系是"调用/参照"，不是"继承代码"**。
3. **要深改 Excel 底层 XML**（普通 openpyxl 干不了的，如复杂图表、特殊样式）→ 借 xlsx 的 `office/unpack.py`、`pack.py` 解包改包。

> 一句话：业务技能保持**独立、config 驱动、可复现**（这是它的价值）；xlsx 提供**通用能力 + 出品规范 + 自检工具**。两层解耦——业务规则变了改业务技能的 config，文档处理能力升级了升基座，互不牵连。

## 数据安全

> ### ⚠️ 本仓是 **PUBLIC 公开仓**（2026-07-25 核实）
>
> GitHub `EvanLee2004/finance-skills` = **public**，Gitee 同为可匿名访问。
> **提交任何东西前先假设"全世界都看得到"。**
>
> 当前已公开、且**存在于 git 历史中**（删文件无法消除，需改写历史）：
> - `order-daily-summary/config/销售组织架构.xlsx`：**26 位同事真实姓名 + 部门归属**
> - 公司**内网地址** `192.168.10.167:18880`（7 个文件）
> - 若干真实同事姓名硬编码在规则/测试里（合并、拆分、抽查等）
>
> **新增内容一律按公开标准写**：不写客户名、不写金额、不写账号、不新增真实姓名。

真实财务数据（源台账 / 回填源 / 成品 / 核销运行工作区等批量数据）**不进仓库**。
各 skill 的维护表作为可长期维护的配置保留在库内——**公开仓下需逐张确认其中不含个人信息**。

**机器执行的那一半** = `.gitignore`；**给人和 agent 看的那一半** = **[docs/仓库边界_什么进什么不进.md](docs/仓库边界_什么进什么不进.md)**（含 push 前三条自检命令）。两边冲突以文档为准，并把 `.gitignore` 补齐。

## 分发与版本

| 项 | 说明 |
|----|------|
| **更新真相源** | **[SOURCE.md](./SOURCE.md)**（仓库 / 白名单 / 一句话更新 / config 铁律） |
| 源码仓（双端） | **GitHub** `EvanLee2004/finance-skills` + **Gitee** `Lee157/finance-skills` · 分支 **`main`** |
| **开发交付** | `git push origin main` 即上云；**默认不打 zip** |
| **同事更新** | 对 opencode 说 **「更新财务skills」**（从云端 main 白名单覆盖；保留 config；不动自装技能） |
| 版本信号 | 远端 `main` 的 **git short SHA**（不以 zip 版本号为准） |
| opencode 内说明 | 安装后应有 `skills/财务技能包_来源与更新.md` |

## 同事本机：更新财务 skills（主路径）

### 最短：一句话

对 opencode 说：

```
更新财务skills
```

Agent 从 Gitee/GitHub `main` 拉最新 → 只覆盖财务包白名单 15 夹 → 保留本机 `config/` → **不动你自己做的其他 skill** → 汇报 SHA → 你重启 opencode。

完整规则与长提示词见 **[SOURCE.md](./SOURCE.md) 第四节**（与手册 v19 第三节 B 段一致）。

### 开发机已 clone 时（可选手工）

```bash
cd <本机 finance-skills 路径>
git pull origin main    # 国内：git pull gitee main
git log -1 --oneline
# 再把 skills/ 下白名单覆盖到 ~/.config/opencode/skills/（保留本机各 skill 的 config/）
# 重启 opencode
```

### 可复制长提示词（与 SOURCE 第四节相同）

```
更新财务skills

请按官方 monorepo 把本机财务部官方技能更新到云端最新，全程自动完成，要点「允许访问」就允许。

【唯一源】
- Gitee（优先）: https://gitee.com/Lee157/finance-skills  分支 main
- GitHub（备）: https://github.com/EvanLee2004/finance-skills  分支 main
- 仓内路径: skills/<技能id>/
- 若本机已有「财务技能包_来源与更新.md」或 SOURCE.md，先读确认。

【红线·只动财务包，别碰我别的技能】
- 只能更新/新增下面白名单文件夹；白名单以外一律不删、不改、不移动、不覆盖。
- 禁止清空整个 skills 目录；禁止「只保留这 15 个」；禁止重命名白名单外的夹。

【财务包白名单】
receivables-merge、split-by-sales、labor-invoice-check、withholding-report-rename、compliance-spot-check、dreame-ar-progress-diff、dept-expense-alloc、ar-hexiao-daily、order-daily-summary、task-clarifier、xlsx、docx、pptx、pdf、env-doctor
（另：把「财务技能包_来源与更新.md」放到 skills 目录根。）

【装到哪】
~/.config/opencode/skills/（Windows = %USERPROFILE%\.config\opencode\skills\）

【怎么取源】
优先 git clone/pull 上述仓库 main（国内优先 Gitee）。不要问我要 zip；不要等我发压缩包。

【步骤】
1）git 拉到最新 main，记下 short SHA。
2）只对白名单：用仓内 skills/ 覆盖安装目录源码；新技能整夹复制。
3）⚠ 保留我本地 config：某技能本机已有 config/ 则绝不覆盖；没有才从仓库复制。
4）可选：若存在已下线旧夹 payroll-info-match / insurance-fund-merge / bank-income-extract 才删它们。
5）依赖可顺手补（国内镜像）：
   pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pandas openpyxl xlrd pypdf pdfplumber pdf2image python-docx python-pptx markitdown lxml defusedxml Pillow requests playwright
   playwright install chromium

【汇报·必须逐条】
- 更新到的 git short SHA + 远端（Gitee/GitHub）
- 更新/新增了哪些财务技能
- 白名单外其他技能：必须写「未动」
- 我的 config/维护表保住没
- 提醒我重启 opencode
```

### 双端 push（开发侧 · 本机已配好）

```bash
# origin：fetch 走 GitHub；push 同时推 GitHub + Gitee
git push origin main          # 一键双端 = 交付完成

# 只推某一端时：
git push https://github.com/EvanLee2004/finance-skills.git main
git push gitee main
```

当前 `git remote -v` 应为：

- `origin` fetch → GitHub  
- `origin` push → GitHub **和** Gitee  
- `gitee` → 仅 Gitee（备用）

改 skill 后：本地测绿 → **`git push origin main`（双端）即交付**。默认**不**再打 zip、不挂 Release 附件。
