# -*- coding: utf-8 -*-
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

# 测试数据根（项目内，只读）
TEST_DATA = Path(
    "/Users/evanlee/Documents/甲骨易实习/项目/长期项目/应收核销自动化（李明妹）/测试数据"
)
FIXTURE = TEST_DATA / "步骤6_核销判定" / "智云取数夹具_20260708整天53笔_含标准答案.json"
BANK_XLSX = TEST_DATA / "步骤2_收入提取" / "银行日记账_样例_7账户.xlsx"
LEDGER_SMALL = TEST_DATA / "步骤7_回填" / "盈亏表_1月样例_小体积.xlsx"
LEDGER_FULL = TEST_DATA / "步骤7_回填" / "盈亏核算表2026全年_副本.xlsx"

# 主回归闸：2026-07-22 真实 13 笔 + 明妹当天手工填完的副本（金标）
# 真实财务数据不进仓库；本地没有就自动跳过相关用例。
GOLD_DIR = TEST_DATA / "步骤6_核销判定" / "20260722_真实13笔_金标"
