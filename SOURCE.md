# 财务技能包 · 来源与更新（单一真相源）

> **以后凡「更新技能 / 版本从哪来 / push 到哪」先看本文件。**  
> 本 monorepo = 财务部全部 skill 的**唯一源码仓**；同事本机 opencode 里的文件夹是它的**安装副本**。

---

## 一、仓库在哪（就这一个项目）

| 用途 | 地址 |
|------|------|
| **主仓（开发 push / fetch）** | GitHub **https://github.com/EvanLee2004/finance-skills** |
| **国内镜像（同事下包 / 国内 pull）** | Gitee **https://gitee.com/Lee157/finance-skills** |
| 分支 | 默认 **`main`**（日常只跟 main，别跟杂分支） |

本地开发目录（明昊机器）：

```text
…/项目/长期项目/财务部skills/finance-skills/
```

- 所有业务 skill 都在子目录 **`skills/<技能id>/`**
- **不是**每个 skill 一个仓库；**一个仓装全部**，改完 `git push origin main` 双端更新

---

## 二、谁该干什么

| 角色 | 更新动作 |
|------|----------|
| **开发（明昊 / AI）** | 在本仓改码 → 测绿 → `git push origin main`（GitHub+Gitee）→ 需要给同事用时再打 zip / 升手册版本 |
| **同事本机（已能访问私有仓）** | `git pull` 本仓 `main`，再把 `skills/*` **白名单覆盖**进 opencode skills 目录（**保留本机 config**） |
| **同事本机（只用 zip）** | 飞书/Gitee Release 下最新 `财务技能包_vX.Y.Z.zip` → 用手册/README 里的「更新提示词」让 opencode 覆盖安装 |

> ⚠️ **push 了 ≠ 同事本机已更新。** 同事 opencode 不会自动跟着 git 走，必须再走「pull 同步」或「zip 覆盖」其中一条。

---

## 三、装到本机的哪（opencode）

| 系统 | 技能目录（一般） |
|------|------------------|
| macOS / Linux | `~/.config/opencode/skills/` |
| Windows | `%USERPROFILE%\.config\opencode\skills\` |

目录里每个财务 skill 一个夹（如 `labor-invoice-check`、`receivables-merge`）。  
**更新只动财务包白名单夹**；同事自己装的其他 skill 不许删。

**财务包白名单（15）**  
`receivables-merge` · `split-by-sales` · `labor-invoice-check` · `withholding-report-rename` · `compliance-spot-check` · `dreame-ar-progress-diff` · `dept-expense-alloc` · `ar-hexiao-daily` · `order-daily-summary` · `task-clarifier` · `xlsx` · `docx` · `pptx` · `pdf` · `env-doctor`

**本地 config 铁律**：本机已有的 `config/`（销售归属、组织架构、config.local 等）**更新时不覆盖**；只有本机没有该技能 config 时才从新包装入。

---

## 四、推荐更新路径（按优先级）

### A. 开发机 / 已 clone 本仓

```bash
cd <finance-skills 本地路径>
git pull origin main          # 或 git pull gitee main（国内）
git log -1 --oneline          # 确认 tip，例如 7e43a80
# 再同步 skills/* → opencode skills 目录（白名单覆盖、跳过 config）
# 然后重启 opencode
```

### B. 同事只收 zip（当前默认分发）

1. 拿到新的 `财务技能包_vX.Y.Z.zip` 放到桌面  
2. 粘贴 README「同事本机：更新财务技能包」整段提示词到 opencode  
3. 允许访问 → 跑完 → **重启 opencode**

### C. 问 AI「帮我更新财务技能」时

Agent 应：

1. **先读本文件**（或 `skills/财务技能包_来源与更新.md`）确认仓库地址  
2. 有 git 副本 → 优先 `git pull` 上述仓库 `main`  
3. 无 git、只有 zip → 按 zip 提示词覆盖  
4. 汇报：更新到哪个 **commit short SHA** 或 **zip 版本号**；config 是否保留；是否需重启

---

## 五、版本怎么对

| 信号 | 含义 |
|------|------|
| `git rev-parse --short HEAD` | 源码真实版本（开发侧以它为准） |
| `财务技能包_vX.Y.Z.zip` | 给同事的打包点；**可能落后于 main 若干 commit** |
| 使用手册版本（v18…） | 给人看的说明版本，与 zip 大致对齐 |

同事问「我是不是最新」：

- 有 git：`git log -1 --oneline` 是否等于远端 `main`  
- 只有 zip：飞书/Release 上的版本号是否最新；并对照本文件「当前分发包」说明（见 README）

---

## 六、和旁边文件夹的关系（别搞混）

| 路径 | 是什么 | 进 git 吗 |
|------|--------|-----------|
| `财务部skills/finance-skills/` | **本仓 = 唯一 skill 源码** | ✅ 本仓库 |
| `财务部skills/技能/<中文名>/` | 本地资料（方案/测试数据/录音） | ❌ 一般不进本仓 |
| `财务部skills/发布/` | zip + 给同事手册 | 另管；附件可挂 Release |

---

## 七、给 opencode / 其他 AI 的硬提示

```
财务部官方 skill 只维护在这一个 monorepo：
  GitHub: EvanLee2004/finance-skills
  Gitee:  Lee157/finance-skills
  分支:   main
更新时去上述仓库 pull 或用官方 zip 覆盖；不要自建平行仓库、不要只改本机 skills 却不回推。
回推代码只在有写权限的开发机对 main 做；同事机默认只读 pull / 装包。
```
