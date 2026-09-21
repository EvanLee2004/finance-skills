# 财务技能包 · 来源与更新（单一真相源）

> **以后凡「更新技能 / 版本从哪来 / push 到哪 / 同事怎么更新」先看本文件。**  
> 本 monorepo = 财务部全部官方 skill 的**唯一源码仓**；同事本机 opencode 里的官方技能夹是它的**安装副本**。  
> **2026-07-25 起：默认分发 = git push Gitee `main`。不再发 zip。**

---

## 一、仓库在哪（就这一个项目）

| 用途 | 地址 |
|------|------|
| **发布源（开发 push / 同事 pull）** | Gitee **https://gitee.com/Lee157/finance-skills** |
| 分支 | 默认 **`main`** |
| 版本 | 每个 skill 的 `SKILL.md` YAML `version`，2026-09-15 起全员 **1.0**；改哪个 skill 只给那个 +0.1 |

本地开发目录（明昊机器）：

```text
…/项目/长期项目/财务部skills/finance-skills/
```

- 所有业务 skill 都在 **`skills/<技能id>/`**
- **不是**每个 skill 一个仓库；**一个仓装全部**
- 应收核销日清 `ar-hexiao-daily`、九点下单、部门费用……全部在本 monorepo 里；**没有**单独平行仓库要同步
- 开发改完：测绿 → **`git push gitee main`** → 云端即最新；**不必再打 zip**。默认不推 GitHub。

---

## 二、谁该干什么（一句话版）

| 角色 | 动作 |
|------|------|
| **开发** | 改码 → 该 skill `version` +0.1 → 测绿 → `git push gitee main`。 |
| **同事本机** | 说「更新财务部skills」= 从 **Gitee** `https://gitee.com/Lee157/finance-skills` 的 `main` 拉到 opencode skills（不要 GitHub）。只覆盖 `pack.json` 白名单；本机 `config/` 默认保留（核销跟仓库）；不要清空自己另装的技能。 |

---

## 三、装到本机的哪（opencode）

| 系统 | 技能目录（一般） |
|------|------------------|
| macOS / Linux | `~/.config/opencode/skills/` |
| Windows | `%USERPROFILE%\.config\opencode\skills\` |

目录里每个官方 skill 一个夹（如 `labor-invoice-check`、`ar-hexiao-daily`）。  
**更新只动财务包白名单夹**；同事自己装的其他 skill 不许删、不许改、不许挪。

白名单见仓库根 `pack.json`。已下线：`env-doctor`、`qige-invoice-to-kingdee`、`task-clarifier`、`update-finance-skills`、`compliance-spot-check`。

核销跟 `main`：更新时覆盖 `ar-hexiao-daily` 的仓内 `config/`。凭据只留 `*.local.json`。

另：根下说明文件 `财务技能包_来源与更新.md` 一并覆盖更新（方便下次还能找到本说明）。

**本地 config 铁律**：本机已有的 `config/`（销售归属、组织架构、`config.local.json` 等）**更新时不覆盖**；只有本机没有该技能 config 时才从云端新包装入。

---

## 四、同事怎么更新

从 Gitee `Lee157/finance-skills` 的 `main` 拉到本机 opencode skills。只覆盖 `pack.json` 白名单。本机已有 `config/` 默认不覆盖（核销跟仓库）。不要清空白名单外自己装的夹。开口对照 `skills/财务技能_说什么用哪个.md`。

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
财务部官方 skill 只维护在 Gitee Lee157/finance-skills 的 main。
白名单见 pack.json。不要清空同事自己装的技能。不要发 zip。
开发机测绿后 git push gitee main。
```
