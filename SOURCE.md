# 财务技能包 · 来源与更新（单一真相源）

> **以后凡「更新技能 / 版本从哪来 / push 到哪 / 同事怎么更新」先看本文件。**  
> 本 monorepo = 财务部全部官方 skill 的**唯一源码仓**；同事本机 opencode 里的官方技能夹是它的**安装副本**。  
> **2026-07-25 起：默认分发 = git push 云端；同事一句话「更新财务skills」即可拉最新。不再发 zip 作为主路径。**

---

## 一、仓库在哪（就这一个项目）

| 用途 | 地址 |
|------|------|
| **主仓（开发 push / fetch）** | GitHub **https://github.com/EvanLee2004/finance-skills** |
| **国内镜像（同事 pull 优先）** | Gitee **https://gitee.com/Lee157/finance-skills** |
| 分支 | 默认 **`main`**（日常只跟 main） |

本地开发目录（明昊机器）：

```text
…/项目/长期项目/财务部skills/finance-skills/
```

- 所有业务 skill 都在 **`skills/<技能id>/`**
- **不是**每个 skill 一个仓库；**一个仓装全部**
- 应收核销日清 `ar-hexiao-daily`、九点下单、部门费用……全部在本 monorepo 里；**没有**单独平行仓库要同步
- 开发改完：测绿 → **`git push origin main`**（一键双端 GitHub+Gitee）→ 云端即最新；**不必再打 zip**

---

## 二、谁该干什么（一句话版）

| 角色 | 动作 |
|------|------|
| **开发（明昊 / AI）** | 在本仓改码 → 测绿 → `git push origin main`（GitHub+Gitee）。**到此交付完成**；不用再打包 zip、不用飞书发压缩包。 |
| **同事本机（已装过财务 skills）** | 对 opencode 说一句：**「更新财务skills」**（或「更新财务技能 / 更新财务技能包」）。Agent 从云端 `main` 拉最新 → **只覆盖财务包白名单** → **保留你本地 config** → **不动你自己做的其他 skill**。 |
| **同事本机（首次安装）** | 说「安装财务skills」或粘手册第三节 A 段；Agent 从 Gitee/GitHub clone 后按白名单装入 opencode。 |

> ⚠️ **push 了 ≠ 同事本机已更新。** 云端更新后，同事要说一次「更新财务skills」才会同步到本机。  
> ⚠️ **更新 ≠ 清空 skills。** 只动官方白名单 15 夹；同事自己装的 / 自己写的 skill **一律不删不改**。

---

## 三、装到本机的哪（opencode）

| 系统 | 技能目录（一般） |
|------|------------------|
| macOS / Linux | `~/.config/opencode/skills/` |
| Windows | `%USERPROFILE%\.config\opencode\skills\` |

目录里每个官方 skill 一个夹（如 `labor-invoice-check`、`ar-hexiao-daily`）。  
**更新只动财务包白名单夹**；同事自己装的其他 skill 不许删、不许改、不许挪。

**财务包白名单（15）**  
`receivables-merge` · `split-by-sales` · `labor-invoice-check` · `withholding-report-rename` · `compliance-spot-check` · `dreame-ar-progress-diff` · `dept-expense-alloc` · `ar-hexiao-daily` · `order-daily-summary` · `task-clarifier` · `xlsx` · `docx` · `pptx` · `pdf` · `env-doctor`

另：根下说明文件 `财务技能包_来源与更新.md` 一并覆盖更新（方便下次还能找到本说明）。

**本地 config 铁律**：本机已有的 `config/`（销售归属、组织架构、`config.local.json` 等）**更新时不覆盖**；只有本机没有该技能 config 时才从云端新包装入。

---

## 四、同事一句话更新（主路径 · 必读）

### 触发语（任选其一即可）

- **「更新财务skills」**（推荐，最短）
- 「更新财务技能」/「更新财务技能包」/「把财务 skills 更新到最新」

### Agent 必须执行的步骤（写死）

1. **读来源**：若本机 skills 目录已有 `财务技能包_来源与更新.md` 或本仓 `SOURCE.md`，先打开确认仓库地址与本文一致。  
2. **取最新源码**（优先国内 Gitee，GitHub 作备）：  
   - 本机已有 clone：`git -C <clone路径> pull`（Gitee 或 GitHub 的 `main`）  
   - 本机没有 clone：浅克隆到临时目录  
     `git clone --depth 1 -b main https://gitee.com/Lee157/finance-skills.git`  
     （Gitee 不通再试 `https://github.com/EvanLee2004/finance-skills.git`）  
3. **白名单覆盖**到 opencode skills 目录：  
   - 仅对白名单 15 个夹：覆盖 `SKILL.md`、`scripts/`、`README.md`、`references/` 等源码；白名单里本机还没有的夹整夹复制  
   - **保留**本机各技能已有 `config/`（见第三节铁律）  
   - 把 `skills/财务技能包_来源与更新.md` 放到 skills 目录根  
4. **禁止**：清空整个 skills；删除/改动白名单外任何夹；为「只保留这 15 个」去删同事自装技能。  
5. **可选清理**（仅当存在才删，且只删这些已下线官方旧夹）：`payroll-info-match`、`insurance-fund-merge`、`bank-income-extract`。  
6. **汇报**（必须逐条）：更新到的 **git short SHA**；更新/新增了哪些白名单技能；白名单外其他技能是否「未动」；config 是否保留；提醒 **重启 opencode**。

### 可复制提示词（同事 / Agent 通用）

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
~/.config/opencode/skills/（Windows = %USERPROFILE%\.config\opencode\skills\）；定位不到就找本机 opencode 实际加载技能的目录。

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
- 白名单外其他技能：必须写「未动」（若误动立刻说明）
- 我的 config/维护表保住没
- 提醒我重启 opencode
```

---

## 五、开发侧 push（明昊 / AI）

```bash
cd …/财务部skills/finance-skills
# 测绿后
git push origin main    # origin 已配置双 push：GitHub + Gitee
```

- **不要**再默认打 zip、不要默认飞书发压缩包  
- zip 仅作无网/无 git 的**极端备用**（历史 `发布/` 里旧包可归档；新发版默认不做）  
- 版本真相 = **`git rev-parse --short HEAD`**（云端 `main` tip），不再靠 `财务技能包_vX.Y.Z.zip` 当主版本号  

当前 remote 期望：

- `origin` fetch → GitHub  
- `origin` push → GitHub **和** Gitee  
- `gitee` → 仅 Gitee（备用）

---

## 六、版本怎么对

| 信号 | 含义 |
|------|------|
| `git rev-parse --short HEAD`（远端 main） | **真实版本**（开发与同事更新后都应对齐这个） |
| 使用手册版本（v19…） | 给人看的说明版本，可落后于 main 若干功能 commit |
| 历史 zip `财务技能包_vX.Y.Z` | **旧分发形态**；2026-07-25 起不再作为主更新路径 |

同事问「我是不是最新」：更新后看 Agent 汇报的 short SHA，是否等于  
`https://gitee.com/Lee157/finance-skills` 的 `main` 最新提交。

---

## 七、和旁边文件夹的关系（别搞混）

| 路径 | 是什么 | 进本仓 git 吗 |
|------|--------|----------------|
| `财务部skills/finance-skills/` | **本仓 = 唯一 skill 源码** | ✅ |
| `财务部skills/技能/<中文名>/` | 本地资料（方案/测试数据/录音） | ❌ 一般不进本仓 |
| `财务部skills/发布/` | 给人看的手册等；**不再以 zip 为主交付** | 工作区另管 |
| 同事本机「自己做的 skill」 | 不在白名单内的夹 | ❌ 更新财务 skills **绝不动** |

---

## 八、给 opencode / 其他 AI 的硬提示

```
财务部官方 skill 只维护在这一个 monorepo：
  GitHub: EvanLee2004/finance-skills
  Gitee:  Lee157/finance-skills
  分支:   main
用户说「更新财务skills」= 从上述仓库 pull/clone main → 白名单覆盖进 opencode skills
  → 保留各技能本地 config/ → 绝不碰白名单外的技能。
不要再要求用户下 zip；不要自建平行仓库；不要只改本机 skills 却不回推开发仓。
回推代码只在有写权限的开发机对 main 做；同事机默认只读 pull。
```
