---
name: order-daily-summary
description: >-
  自动登录智云抓「下单」表，按组织架构把销售归到多语（不含运保）/数据/游戏/其他，按日期汇总成「下单数据(万元)」Excel。
  当用户说「九点下单 / 跑下单汇总 / 统计昨天下单 / 下单数据万元 / 按部门下单 / 亮晶下单统计 / 下单日报」，
  或要按工作日规则拉上周五～周日/昨天的下单并出部门万元表时，用本技能。
---

# 九点下单统计

工作日早上用的**下单万元汇总**：登录公司内网智云 → 按日期规则筛下单 → 用 `config/销售组织架构.xlsx` 归四部门 → 出亮晶版式「下单数据」表。

**重要**：部门口径来自**销售组织架构**（本地化→多语不含运保），**不是**利润看板用的智云表内「部门」字段。

## 前置（缺一不可）

1. **内网**：能访问 `http://192.168.10.167:18880`（公司网 / VPN / 部署机）
2. **凭据**（二选一，**勿把密码写进对话记录**）：
   - 本机 `config.local.json`（从 `config/config.local.example.json` 复制到 skill 根目录 `config.local.json`，填 username/password）
   - 环境变量 `ZHIYUN_USER` + `ZHIYUN_PASSWORD`
3. 依赖：`playwright` + chromium、`requests`、`openpyxl`（缺了让 env-doctor 或 `pip install playwright requests openpyxl && playwright install chromium`）

## 你（agent）该怎么干

**核心原则：财务同事说一句「跑九点下单」，你出表、报数字；别甩接口 JSON、别回显密码。**

### ⭐ 人在环

- **缺凭据**：只问一次——「请在本机写好 config.local.json 或设置 ZHIYUN_USER/ZHIYUN_PASSWORD（我不会把密码写进日志）」。
- **未匹配销售非空**：提示她改 `config/销售组织架构.xlsx`（A 销售 / B 分类），改完可重跑。
- **常规**：一句话汇报窗口、明细行数、总计万元、分部门、文件路径。

### 1. 跑脚本（写死绝对路径）

把 `<本skill目录>` 换成**本 SKILL.md 所在文件夹的绝对路径**：

```bash
# A. 正路：工作日早上 09:00–09:15 实时抓（九点快照）
python3 "<本skill目录>/scripts/run.py" \
  --out "<交付目录绝对路径>" \
  [--today YYYY-MM-DD] \
  [--config "<config.local.json 绝对路径>"] \
  [--detail]

# B. 离线：喂 9 点导出的下单明细（与旧 exe 同输入，数字锁导出时刻）
python3 "<本skill目录>/scripts/run.py" \
  --out "<交付目录绝对路径>" \
  --from-xlsx "<早高峰导出的下单.xlsx 绝对路径>" \
  [--today YYYY-MM-DD] \
  [--no-date-filter]
```

- 默认按**今天**算窗口：周一=上周五～周日；周二～周五=昨天；周末若触发=上周五～昨天。
- 要对某固定日复盘：`--today 2026-07-24`（周五→窗口=7/23）。
- `--detail`：明细含 **SO / 订单名称**（实时默认可选；离线模式默认开明细）。
- **⚠ 下午实时抓到的「昨日」常小于 9 点**：智云可改期。差一截不是汇总 bug；要对齐 exe/九点，用路径 A 准点跑或路径 B 喂导出。

### 2. 跑完收尾（人话）

- 「窗口 7/23～7/23，明细 11 行，总计 4.26 万；多语 4.26；未匹配无。表在 xxx.xlsx」
- **若晚于 9:15 且数字明显小于早高峰/旧 exe**：主动说明「实时晚跑 vs 九点快照」差在改期单（点名 SO 若有），**别谎称 skill 算错也别谎称 exe 错**。
- 有未匹配 → 列名单，请她补组织架构。

## 会变的东西（改表不改码）

| 文件 | 用途 |
|------|------|
| `config/销售组织架构.xlsx` | A 销售 → B 本地化/数据/游戏/其他 |
| `config/业务规则.md` | 日期规则、金额字段、展示名说明 |
| `config.local.json`（gitignore） | 账号密码，**禁止提交** |

## 与看板 / 旧 exe

见 `references/与exe及看板差异.md`。分析层对齐死程序 `process_order.py`；取数对齐看板登录+GetFilterRows。
