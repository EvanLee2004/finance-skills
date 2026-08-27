# -*- coding: utf-8 -*-
import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

# 测试数据根（项目内，只读）。可用 AR_HEXIAO_TEST_DATA 覆盖；默认相对本 skill 定位。
# 不写死本机用户名路径（公开仓可 clone）。
_default_test_data = (
    Path(__file__).resolve().parents[1]  # skills/ar-hexiao-daily
    / ".."
    / ".."
    / ".."
    / ".."
    / "应收核销自动化（李明妹）"
    / "测试数据"
).resolve()
TEST_DATA = Path(os.environ.get("AR_HEXIAO_TEST_DATA", str(_default_test_data)))
FIXTURE = TEST_DATA / "步骤6_核销判定" / "智云取数夹具_20260708整天53笔_含标准答案.json"
BANK_XLSX = TEST_DATA / "步骤2_收入提取" / "银行日记账_样例_7账户.xlsx"
LEDGER_SMALL = TEST_DATA / "步骤7_回填" / "盈亏表_1月样例_小体积.xlsx"
LEDGER_FULL = TEST_DATA / "步骤7_回填" / "盈亏核算表2026全年_副本.xlsx"

# 主回归闸：2026-07-22 真实 13 笔 + 明妹当天手工填完的副本（金标）
# 真实财务数据不进仓库；本地没有就自动跳过相关用例。
GOLD_DIR = TEST_DATA / "步骤6_核销判定" / "20260722_真实13笔_金标"

REAL_FIXTURES = {
    "金标盈亏表": GOLD_DIR / "02_我的表副本" / "2026年盈亏核算表1-12月（副本）.xlsx",
    "银行日记账": BANK_XLSX,
    "全年盈亏表": LEDGER_FULL,
    "汇款到账流转表": TEST_DATA / "步骤7_回填" / "到账流转表_汇款7月_副本.xlsx",
    "微信到账流转表": TEST_DATA / "步骤7_回填" / "到账流转表_微信全年_副本.xlsx",
}


def pytest_sessionstart(session) -> None:
    """Private CI can require every controlled real-structure regression fixture."""
    del session
    if os.environ.get("AR_HEXIAO_REQUIRE_REAL_FIXTURES", "").strip() != "1":
        return
    missing = [name for name, path in REAL_FIXTURES.items() if not path.is_file()]
    if missing:
        raise pytest.UsageError(
            "受控真实结构回归已启用，但缺少测试数据：" + "、".join(missing)
        )
