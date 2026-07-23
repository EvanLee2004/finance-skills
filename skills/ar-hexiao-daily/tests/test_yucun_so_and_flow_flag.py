# -*- coding: utf-8 -*-
"""2026-07-23 口径：预存明细读 SO；流转空/部分=待办。"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import classify_hexiao as C  # noqa: E402
import rescan_holds as R  # noqa: E402


def test_flow_needs_update_empty_and_partial():
    assert R.flow_needs_update(None) is True
    assert R.flow_needs_update("") is True
    assert R.flow_needs_update("  ") is True
    assert R.flow_needs_update("部分") is True
    assert R.flow_needs_update("部份") is True
    assert R.flow_needs_update("是") is False
    assert R.flow_needs_update("yes") is False


def test_load_excel_reads_yucun_so(tmp_path):
    """核销明细带 SO 列时，预存展开记录必须带上 SO（以前写死 so=''）。"""
    ws_root = tmp_path / "工作区"
    exp = ws_root / "01_智云导出"
    exp.mkdir(parents=True)

    # 回款记录：一笔预存
    wb = openpyxl.Workbook()
    w = wb.active
    w.append(
        [
            "回款记录ID",
            "回款类型",
            "核销日期",
            "到账日期",
            "到账金额/原币",
            "手续费/原币",
            "原币币种",
            "核销状态",
            "开票客户",
        ]
    )
    w.append(
        [
            "AR26079999",
            "预存回款",
            "2026-07-21",
            "2026-07-20",
            100,
            0,
            "人民币CNY",
            "核销成功",
            "测试客户",
        ]
    )
    wb.save(exp / "回款记录_test.xlsx")

    # 对账：预存应 0 行——给一个空表头即可
    wb2 = openpyxl.Workbook()
    w2 = wb2.active
    w2.append(
        [
            "回款记录ID",
            "SO",
            "SOD",
            "回款额/原币",
            "核销日期",
            "客户名称",
            "结算周期",
        ]
    )
    wb2.save(exp / "回款核销对账_test.xlsx")

    # 核销明细：带 SO
    wb3 = openpyxl.Workbook()
    w3 = wb3.active
    w3.append(
        [
            "回款记录NUM",
            "核销日期",
            "本次核销金额",
            "币种",
            "汇率",
            "SO",
            "订单名称",
        ]
    )
    w3.append(
        ["AR26079999", "2026-07-21", 100, "人民币CNY", 1, "SO26070111", "某单"]
    )
    wb3.save(exp / "核销明细_test.xlsx")

    recs = C.load_excel_exports(ws_root)
    assert recs is not None
    yucun = [r for r in recs if r.get("channel") == "yucun" or "预存" in (r.get("huikuan_type") or "")]
    assert yucun, recs
    assert any(r.get("so") == "SO26070111" for r in yucun), yucun
