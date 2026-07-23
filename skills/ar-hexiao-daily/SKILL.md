---
name: ar-hexiao-daily
description: >-
  应收核销日清：编排出纳李明妹每天回款核销（智云取数、核销判定、工作清单、回填审核单、
  确认后写盈亏明细副本、挂账重扫）。人在环硬闸：AI 回填前必须出审核 Excel，她确认才写表。
  当用户说「跑昨天的核销 / 做核销 / 应收核销 / 回款核销 / 核销清单 / 挂账还有哪些没做 /
  日清 / 昨天到账核销 / 重扫挂账 / 收入提取 / 拉智云 / 抓核销 / 导智云 /
  回填审核 / 确认回填 / 可以写了 / 按这个写」，用本技能。
---

# 应收核销日清（人在环 · 回填审核闸）

**主战场从第 6 步开始**（明妹钉死）：智云按核销日筛 → 判定 → **回填审核单** → 她确认 → 写工作副本「明细」。  
1–4 她提才做；第 5 步销售不绑隔天。

**数据源**：第 6 步智云 = 核销真相；第 7 步流转表 = 备份 + 写单号 +「是否更新」状态，**回填金额不从第 7 推**。  
**写表**：只动盈亏「明细」；其它 sheet 禁止。  
**两回填**：①盈亏五列（脚本可写，**必须先审核**）②流转单号/是否更新（**只建议，默认不自动写流转表**）。  
**永不写智云。** 金额只由 `scripts/*.py` 算。

## 0. 按她的话直接跳

| 她说的话 | 直接做 | 命令链 |
|---|---|---|
| 跑昨天的核销 / 出清单 / 日清 | **6→判定→清单→校验→审核单，然后停** | `classify` → `build_worklist` → `validate_plan` → `build_review_sheet` |
| 确认 / OK / 可以写 / 按这个写 | **仅在有审核单后** `apply --confirmed` | 见 §3.1 |
| 拉智云 / 导智云 | 取数进 01_ | `fetch_zhiyun.py` |
| 挑收入 / 日记账 | 第 2 步（次要） | `extract_income.py` |
| 挂账 / 重扫 / 空的和部分 | 重扫 | `rescan_holds.py --ledger …` |
| 我文件放好了 / 缺什么 | 盘点 | `inspect_inputs.py` |

**「跑昨天的核销」= 出到审核单为止，禁止顺手 apply。**

### 判定+审核单跑完 · 结尾照抄（一个字别改）

```
✅ 昨天的核销判定完了，回填还没写进你的表
📁 工作清单：<今日工作清单_YYYYMMDD.xlsx>
📁 回填审核单：<回填审核单_YYYYMMDD.xlsx>  ← 先看这个
   · 「盈亏回填·拟写入」N 笔 —— 改前→改后、依据都在表里
   · 「冲突·需你定」C 笔 —— 程序不会自动写
   · 「流转表·是否更新建议」—— 空/是/部分（空不要公式；部分未更新标红）
   · 挂账 M 笔、异常 K 笔见工作清单
🔒 你的盈亏表还一个字节都没动
👉 打开审核单扫一眼；没问题回我「确认」或「可以写」，我再写入明细副本（写前自动备份）
```

⛔ 禁止说「填进智云」。⛔ 禁止未确认就 apply。

## 0.5 硬禁止

1. 禁止自写 Python 读她的 Excel；只准调 `scripts/`。  
2. 禁止把客户名、金额明细打进对话。  
3. 缺输入就停下报缺，禁止凑合。  
4. **禁止跳过回填审核单直接 `apply_to_copy`**（脚本无 `--confirmed` 也会拒绝）。  
5. 禁止自动写到账流转表（第一版只出建议；她要手填或以后另开）。

## 0.6 怎么跟她说话

短 · 像人 · 主动要料。细则 → `references/跟明妹沟通.md`。

- ❌ 报脚本名/E 码/auto=27  
- ✅ 「判完了：**N 笔拟写入、C 笔冲突**。审核单在 <路径>，你说确认我再写。」

## 0.8 判不准就问她

弱命中 / 一 SO 多行 / 手续费 / 尾差 / 跨年 → 带选项问（见 `跟明妹沟通.md` + `task-clarifier`）。答案记 `config/业务规则.md`。

## 1. 红线

1. 金额只由脚本算。  
2. 对话不回显客户名金额。  
3. 分笔一律 hold，禁止均分/比例。  
4. 永不写智云；账密不落盘。  
5. 写盈亏只许「明细」；`--in-place` 写前备份、写后回读，不过不替换。  
6. **回填 = 审核单 → 她确认 → `apply --confirmed`**。  
7. 回填数据源 = 第 6 步智云，不用流转表金额。  
8. 列名模糊匹配；找不到列 → 报错列实际表头。

## 2. 工作区

```
工作区/
  01_智云导出/    ← 智云导出
  02_我的表副本/  ← 盈亏、流转副本（apply 前只读；确认后可 --in-place 写盈亏明细）
  03_台账/        ← 挂账台账
  04_产出/        ← 清单、审核单、校验计划、变更清单
```

## 3. 流程编排（11 步）

每步三句：`✅ 做完了` / `📁 路径` / `👉 下一步 + 触发词`。

| 步 | 名称 | 谁做 | 入口 |
|----|------|------|------|
| 1–4 | 日记账→挑收入→流转→建回款 | **她做**（次要） | 不提别串 |
| 5 | 销售核销 | 销售 | 不绑隔天 |
| **6** | 智云筛核销日 | 她导/fetch | → `01_` |
| 6b | 核销判定 | 脚本 | `classify_hexiao.py` |
| 7 | 流转备份/写单号 | 她+清单建议 | **不自动写流转表**；建议见审核单/清单 |
| **8a** | 工作清单 | 脚本 | `build_worklist.py` |
| **8b** | 校验 + **回填审核单** | 脚本 | `validate_plan` → `build_review_sheet` → **停等确认** |
| **8c** | 写明细副本 | 脚本 | 确认后 `apply_to_copy --confirmed`（可 `--in-place`） |
| 9 | 挂账重扫 | 脚本 | `rescan_holds`；空+部分 |
| 10–11 | 核对/收工 | 她 | 比其它 sheet 总数 |

### 3.1 回填硬闸（plan → validate → review → confirm → execute）

```bash
SKILL="<本skill目录>"
WS="$SKILL/工作区"
LEDGER="$WS/02_我的表副本/<盈亏副本.xlsx>"   # 她固定用的那份

python3 "$SKILL/scripts/verify_sources.py" snapshot --workspace "$WS"
python3 "$SKILL/scripts/inspect_inputs.py"  --workspace "$WS"
python3 "$SKILL/scripts/classify_hexiao.py" --workspace "$WS"   # 或 --fixture …
python3 "$SKILL/scripts/build_worklist.py"  --workspace "$WS"
python3 "$SKILL/scripts/validate_plan.py" \
  --plan "$WS/04_产出/判定结果_*.json" --ledger "$LEDGER" \
  --out "$WS/04_产出/写入计划_校验后.json"
python3 "$SKILL/scripts/build_review_sheet.py" \
  --checked "$WS/04_产出/写入计划_校验后.json" \
  --ledger "$LEDGER" --workspace "$WS"
# ★ 到这里必须停。把审核单路径给她。禁止继续 apply。

# 仅当她说「确认/OK/可以写/按这个写」：
python3 "$SKILL/scripts/apply_to_copy.py" \
  --checked "$WS/04_产出/写入计划_校验后.json" \
  --ledger "$LEDGER" --confirmed --in-place
# 头几次现场可去掉 --in-place，出新文件并排验

python3 "$SKILL/scripts/verify_sources.py" verify --workspace "$WS"
```

确认词（同义）：`确认` `OK` `ok` `可以写` `按这个写` `没问题写吧` `写吧`。  
她说改/跳过某笔 → 先改计划或加 `--force` 只写无冲突的，**重新出审核单**再确认。

### 3.2 流转「是否更新」三态（只建议）

| 状态 | 填法 | 公式 | 颜色 |
|------|------|------|------|
| 全部订单都没检索到/不能更 | **留空** | **不要公式** | — |
| 全部检索到并已更新 | **是** | 可套她惯用公式，否则字面「是」 | — |
| 部分更新 | **部分** | 不要写成「是」 | **未更新 SO 标红** |

实现：`flow_ledger.flow_status_policy`；清单「按到账汇总」+ 审核单「流转表·是否更新建议」。

### 开关

- `--flow-complete`：当天**所有渠道**流转表都给了才加（否则不误判 E0）。  
- `--ledger`（rescan）：给了才对挂起项真重判。  
- `--confirmed`（apply）：**人工确认后才写**。  
- `--in-place`（apply）：就地写同一份副本（写前备份）。

规则 → `config/业务规则.md`、`references/规则_*.md`；给她看 → `references/使用说明_给明妹.md`。

## 4. 你（agent）怎么干

1. 盘点输入；缺料一次说清。  
2. 主路径跑到 **build_review_sheet** 停下，用 §0 照抄结尾。  
3. 等确认词 → 再 `apply --confirmed`。  
4. 对话只报笔数/路径/成败，不报明细。  
5. 依赖：`pip install openpyxl xlrd`；智云另需 requests/playwright。  
6. 测试：`python3 -m pytest tests/ -v`（应 ≥120 例全绿）。

## 5. 验收

- pytest 全绿。  
- 无 `--confirmed` 时 apply 必须非 0。  
- 真跑后 `verify_sources.py verify` 退出 0（未确认前源表哈希不变）。

## 6. 已知边界（不许夸大）

- 预存/分笔/手续费/尾差/一 SO 多行/跨年：见业务规则；未拍板默认 hold。  
- 流转表**不自动写**；红字字体 openpyxl 难稳读。  
- 她公式的具体式子以她表为准，程序给策略不发明公式。
