# -*- coding: utf-8 -*-
import datetime as dt
import json
from pathlib import Path

import openpyxl
import pytest

import classify_hexiao as C


def _xlsx(path: Path, headers, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def _bundle(root: Path, day: str, payments, orders, details, sods):
    d = root / "01_智云导出"
    _xlsx(
        d / f"回款记录_{day}.xlsx",
        ["回款记录ID", "核销日期", "到账日期", "到账金额/原币", "到账金额/本币",
         "手续费/原币", "原币币种", "回款类型", "核销状态", "开票客户"],
        payments,
    )
    _xlsx(
        d / f"订单交付_{day}.xlsx",
        ["回款记录ID", "SO", "交付额/原币", "汇率", "结算币种", "订单名称"],
        orders,
    )
    _xlsx(
        d / f"核销明细_{day}.xlsx",
        ["回款记录NUM", "核销日期", "本次核销金额", "SO"],
        details,
    )
    _xlsx(
        d / f"订单明细_{day}.xlsx",
        ["SO", "SOD", "交付额/原币"],
        sods,
    )


def test_target_date_selects_exact_bundle_and_recovers_shifted_parent(tmp_path):
    _bundle(
        tmp_path, "20260727",
        [["AR27", dt.date(2026, 7, 27), dt.date(2026, 7, 24), 100, 100, 0,
          "人民币CNY", "", "核销成功", "客户甲"]],
        [["AR27", "SO27", 100, 1, "人民币CNY", "订单甲"]],
        [],
        [["SO27", "SOD27", 100]],
    )
    _bundle(
        tmp_path, "20260728",
        [["ARL", dt.date(2026, 7, 28), dt.date(2026, 7, 24), 70, 70, 0,
          "人民币CNY", "", "核销成功", "客户乙"]],
        [["ARL", "SOL", 70, 1, "人民币CNY", "订单乙"]],
        [["ARL", dt.date(2026, 7, 27), 70, "SOL"]],
        [["SOL", "SODL", 70]],
    )

    payments = C.load_exports(tmp_path, dt.date(2026, 7, 27))
    assert {p["ar"] for p in payments} == {"AR27", "ARL"}
    shifted = next(p for p in payments if p["ar"] == "ARL")
    assert shifted["parent_hexiao_date"] == dt.date(2026, 7, 28)
    assert shifted["hexiao_date"] == dt.date(2026, 7, 27)
    assert shifted["writeoffs"] == {"SOL": 70.0}
    assert shifted["_source_meta"]["historical_detail_rows"] == 1

    records = C.expand_payments(payments, {})
    coverage = C.source_coverage(payments, records)
    assert coverage["complete"] is True
    assert coverage["expected_order_keys"] == 2
    assert coverage["historical_detail_rows"] == 1


def test_shifted_only_date_can_run_without_same_day_main_bundle(tmp_path):
    _bundle(
        tmp_path, "20260728",
        [["ARL", dt.date(2026, 7, 28), dt.date(2026, 7, 24), 70, 70, 0,
          "人民币CNY", "", "核销成功", "客户乙"]],
        [["ARL", "SOL", 70, 1, "人民币CNY", "订单乙"]],
        [["ARL", dt.date(2026, 7, 26), 70, "SOL"]],
        [["SOL", "SODL", 70]],
    )
    payments = C.load_exports(tmp_path, dt.date(2026, 7, 26))
    assert [p["ar"] for p in payments] == ["ARL"]
    assert payments[0]["hexiao_date"] == dt.date(2026, 7, 26)
    assert payments[0]["writeoffs"] == {"SOL": 70.0}


def test_target_date_does_not_silently_pick_first_of_many_files(tmp_path):
    _bundle(
        tmp_path, "20260727",
        [["AR27", dt.date(2026, 7, 27), dt.date(2026, 7, 27), 10, 10, 0,
          "人民币CNY", "", "", "甲"]],
        [["AR27", "SO27", 10, 1, "人民币CNY", ""]],
        [], [["SO27", "SOD27", 10]],
    )
    _bundle(
        tmp_path, "20260728",
        [["AR28", dt.date(2026, 7, 28), dt.date(2026, 7, 28), 20, 20, 0,
          "人民币CNY", "", "", "乙"]],
        [["AR28", "SO28", 20, 1, "人民币CNY", ""]],
        [], [["SO28", "SOD28", 20]],
    )
    with pytest.raises(C.InputError, match="必须用 --hexiao-date"):
        C.load_exports(tmp_path)
    assert {p["ar"] for p in C.load_exports(tmp_path, dt.date(2026, 7, 28))} == {"AR28"}


def test_blank_order_delivery_is_recovered_from_complete_sod_sum(tmp_path):
    _bundle(
        tmp_path, "20260727",
        [["AR1", dt.date(2026, 7, 27), dt.date(2026, 7, 24), 100, 100, 0,
          "人民币CNY", "", "", "甲"]],
        [["AR1", "SO1", None, 1, "人民币CNY", ""]],
        [],
        [["SO1", "SOD1", 40], ["SO1", "SOD2", 60]],
    )
    payments = C.load_exports(tmp_path, dt.date(2026, 7, 27))
    assert payments[0]["orders"][0]["deliver"] == 100.0
    assert payments[0]["orders"][0]["delivery_source"] == "订单明细SOD合计"
    records = C.expand_payments(payments, {})
    assert {r["sod"] for r in records} == {"SOD1", "SOD2"}
    assert C.source_coverage(payments, records)["recovered_delivery_orders"] == 1


def test_incomplete_sod_amount_does_not_guess_delivery(tmp_path):
    _bundle(
        tmp_path, "20260727",
        [["AR1", dt.date(2026, 7, 27), dt.date(2026, 7, 24), 100, 100, 0,
          "人民币CNY", "", "", "甲"]],
        [["AR1", "SO1", None, 1, "人民币CNY", ""]],
        [],
        [["SO1", "SOD1", 40], ["SO1", "SOD2", None]],
    )
    payments = C.load_exports(tmp_path, dt.date(2026, 7, 27))
    assert payments[0]["orders"][0]["deliver"] is None
    records = C.expand_payments(payments, {})
    assert records[0]["forced_code"] == "E7"
    assert C.source_coverage(payments, records)["recovered_delivery_orders"] == 0


def test_order_level_coverage_catches_missing_so():
    payment = {
        "ar": "AR1", "orders": [{"so": "SO1"}, {"so": "SO2"}],
        "writeoffs": {}, "_source_meta": {},
    }
    with pytest.raises(C.CoverageError, match=r"AR1\|SO2"):
        C.source_coverage(
            [payment],
            [{"ar": "AR1", "so": "SO1", "sod": "SOD1", "amount_orig": 1}],
        )


def test_shifted_detail_audit_groups_by_real_date(tmp_path):
    _bundle(
        tmp_path, "20260728",
        [["ARL", dt.date(2026, 7, 28), dt.date(2026, 7, 24), 70, 70, 0,
          "人民币CNY", "", "", "乙"]],
        [["ARL", "SOL", 70, 1, "人民币CNY", ""]],
        [["ARL", dt.date(2026, 7, 27), 70, "SOL"]],
        [["SOL", "SODL", 70]],
    )
    found = C.find_shifted_detail_dates(tmp_path)
    assert found["2026-07-27"]["rows"] == 1
    assert found["2026-07-27"]["ars"] == ["ARL"]


def test_shifted_detail_assessment_only_flags_missing_keys(tmp_path):
    _bundle(
        tmp_path, "20260728",
        [["ARL", dt.date(2026, 7, 28), dt.date(2026, 7, 24), 70, 70, 0,
          "人民币CNY", "", "", "乙"]],
        [["ARL", "SOL", 70, 1, "人民币CNY", ""]],
        [["ARL", dt.date(2026, 7, 27), 70, "SOL"]],
        [["SOL", "SODL", 70]],
    )
    pending = C.assess_shifted_detail_dates(
        tmp_path, date_from=dt.date(2026, 7, 27), date_to=dt.date(2026, 7, 27)
    )
    assert pending["2026-07-27"]["needs_rerun"] is True

    out = tmp_path / "04_产出"
    out.mkdir()
    (out / "判定结果_20260727.json").write_text(
        json.dumps({"auto": [{"ar": "ARL", "so": "SOL"}]}, ensure_ascii=False),
        encoding="utf-8",
    )
    done = C.assess_shifted_detail_dates(
        tmp_path, date_from=dt.date(2026, 7, 27), date_to=dt.date(2026, 7, 27)
    )
    assert done["2026-07-27"]["needs_rerun"] is False
    assert done["2026-07-27"]["missing_order_keys"] == []
