# -*- coding: utf-8 -*-
"""核销明细应保留本币金额，并累计目标日之前的历史核销。"""
import datetime as dt
from pathlib import Path

import openpyxl

import classify_hexiao as C


def _xlsx(path: Path, headers, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_loader_keeps_current_local_and_cross_ar_cumulative_writeoffs(tmp_path):
    export_dir = tmp_path / "01_智云导出"
    _xlsx(
        export_dir / "回款记录_20260724.xlsx",
        [
            "回款记录ID", "核销日期", "到账日期", "到账金额/原币", "到账金额/本币",
            "手续费/原币", "原币币种", "回款类型", "核销状态", "开票客户",
        ],
        [["AR1", dt.date(2026, 7, 24), dt.date(2026, 7, 24), 1030.47, 7410.12,
          0, "美元USD", "", "核销成功", "客户甲"]],
    )
    _xlsx(
        export_dir / "订单交付_20260724.xlsx",
        ["回款记录ID", "SO", "交付额/原币", "汇率", "结算币种", "订单名称"],
        [["AR1", "SO1", 4083.60, None, "美元USD", "订单甲"]],
    )
    _xlsx(
        export_dir / "核销明细_20260724.xlsx",
        ["回款记录NUM", "核销日期", "本次核销金额", "本次核销金额/本币", "SO"],
        [
            # 历史首款来自另一父回款 AR0；累计必须仍按 SO 合并到当前 AR1。
            ["AR0", dt.date(2026, 6, 11), 3053.13, 21949.18, "SO1"],
            ["AR1", dt.date(2026, 7, 24), 1030.47, 7410.12, "SO1"],
        ],
    )
    _xlsx(
        export_dir / "订单明细_20260724.xlsx",
        ["SO", "SOD", "交付额/原币"],
        [["SO1", "SOD1", 4083.60]],
    )

    payment = C.load_exports(tmp_path, dt.date(2026, 7, 24))[0]
    assert payment["writeoffs"] == {"SO1": 1030.47}
    assert payment["writeoffs_local"] == {"SO1": 7410.12}
    assert payment["cumulative_writeoffs"] == {"SO1": 4083.60}
    assert payment["cumulative_writeoffs_local"] == {"SO1": 29359.30}
