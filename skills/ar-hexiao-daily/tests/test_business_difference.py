# -*- coding: utf-8 -*-
"""业务值差异：判定、校验、OOXML公式缓存、日清展示端到端回归。"""
import datetime as dt
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import apply_to_copy as A  # noqa: E402
import build_worklist as W  # noqa: E402
import classify_hexiao as C  # noqa: E402
import validate_plan as V  # noqa: E402


HDR = [
    "部门", "销售人员", "客户名称", "单号", "新智云单号", "应收金额",
    "计提金额", "回款明细", "是否结账（是/否）", "收款时间",
    "收款方式(支/汇/现)", "实收金额", "差异",
]


def _ledger(tmp_path, *, with_difference=True, rows=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(HDR if with_difference else HDR[:-1])
    for so, sod, receivable, five, difference in rows or []:
        row = [
            "部", "人", "客", "AB", so, receivable,
            None, None, None, None, None, sod,
        ]
        if with_difference:
            row.append(difference)
        if five:
            row[6:11] = [
                five.get("计提"), five.get("回款明细"), five.get("是否结账"),
                five.get("收款时间"), five.get("收款方式"),
            ]
        ws.append(row)
    path = tmp_path / ("盈亏_有差异.xlsx" if with_difference else "盈亏_无差异.xlsx")
    wb.save(path)
    return path


def _item(row=2, so="SO1", sod="SOD1", *, amount=110.0, difference=-10.0):
    item = {
        "case_id": f"AR1|{so}|{sod}",
        "ar": "AR1",
        "so": so,
        "sod": sod,
        "ledger_row_ref": row,
        "five_cols": {
            "计提": amount,
            "回款明细": amount,
            "是否结账": "是",
            "收款时间": "2026-07-08",
            "收款方式": "汇",
            "实收SOD": sod,
        },
        "current_values": {
            "计提": None, "回款明细": None, "差异": None,
            "是否结账": "", "收款时间": "", "收款方式": None,
            "实收SOD": sod,
        },
    }
    if difference is not None:
        item["derived_cols"] = {"差异": difference}
    return item


def _split_item(row=2, so="SO_SPLIT", sod="SOD_SPLIT"):
    item = _item(row, so, sod, amount=50.0, difference=None)
    item["five_cols"]["计提"] = None
    item["row_operation"] = {
        "type": "split_below",
        "source_receivable": 100.0,
        "paid_receivable": 40.0,
        "unpaid_receivable": 60.0,
        "baseline_receivable": 100.0,
        "paid_side_receivable_total": 40.0,
        "existing_received": 0.0,
        "current_received": 50.0,
        "cumulative_received": 50.0,
        "latest_delivery": 110.0,
        "inserted_five_cols": {
            "计提": None, "回款明细": None, "是否结账": "否",
            "收款时间": None, "收款方式": None, "实收SOD": sod,
        },
    }
    return item


def _synthetic_ledger(receivable=713.24):
    return C.LedgerIndex(synthetic={
        "so": {"SO1": [1]},
        "sod": {},
        "rows": {1: {
            "so": "SO1", "sod": "", "yingshou": receivable,
            "jiti": None, "huikuan": None, "chayi": None,
            "jiezhang": "", "shoukuan_time": None, "shoukuan_way": None,
        }},
    })


def _record(amount=300.0, delivery=765.77, cumulative=765.77):
    return {
        "ar": "AR1", "so": "SO1", "sod": "SOD1",
        "amount_orig": amount, "amount_local": amount,
        "deliver_local": delivery,
        "cumulative_received_local": cumulative,
        "currency": "人民币CNY", "status": "手动核销",
        "hexiao_date": dt.date(2026, 7, 22),
        "shoukuan_date": dt.date(2026, 7, 21),
    }


def test_mismatch_uses_latest_delivery_and_business_difference():
    result = C.classify_one(_record(), _synthetic_ledger(), {}, 0.0, 2026)
    assert result["five_cols"]["计提"] == 765.77
    assert result["five_cols"]["回款明细"] == 300.0
    assert result["five_cols"]["是否结账"] == "是"
    assert result["derived_cols"]["差异"] == -52.53


def test_partial_does_not_write_accrual_or_difference():
    result = C.classify_one(
        _record(amount=300.0, delivery=765.77, cumulative=300.0),
        _synthetic_ledger(),
        {},
        0.0,
        2026,
    )
    assert result["row_operation"]["type"] == "split_below"
    assert result["five_cols"]["计提"] is None
    assert "derived_cols" not in result


def test_delivery_above_baseline_keeps_original_receivable_and_blank_carry_row():
    result = C.classify_one(
        _record(amount=120.0, delivery=150.0, cumulative=120.0),
        _synthetic_ledger(receivable=100.0),
        {},
        0.0,
        2026,
    )

    assert result["bucket"] == "auto"
    assert result["five_cols"]["回款明细"] == 120.0
    assert result["five_cols"]["是否结账"] == "是"
    assert result["five_cols"]["计提"] is None
    assert "derived_cols" not in result
    assert result["row_operation"] == {
        "type": "split_below",
        "receivable_mode": "preserve_baseline_blank_carry",
        "source_receivable": 100.0,
        "baseline_receivable": 100.0,
        "source_row_receivable": 100.0,
        "remaining_unreceived": 30.0,
        "existing_received": 0.0,
        "current_received": 120.0,
        "cumulative_received": 120.0,
        "latest_delivery": 150.0,
        "business_rows": [1],
        "inserted_five_cols": {
            "计提": None,
            "回款明细": None,
            "是否结账": "否",
            "收款时间": None,
            "收款方式": None,
            "实收SOD": "SOD1",
        },
    }


def test_blank_carry_row_accepts_final_payment_and_uses_existing_closeout_rule(tmp_path):
    ledger = _ledger(
        tmp_path,
        rows=[("SO1", "SOD1", 100.0, None, None)],
    )
    first = C.classify_one(
        _record(amount=120.0, delivery=150.0, cumulative=120.0),
        C.LedgerIndex(ledger),
        {},
        0.0,
        2026,
    )
    first_checked = V.validate({"auto": [first]}, V.read_ledger_rows(ledger))
    after_first = tmp_path / "首次特殊部分回款.xlsx"
    A.write_plan(ledger, after_first, first_checked["write"])

    final_record = _record(amount=30.0, delivery=150.0, cumulative=150.0)
    final_record["ar"] = "AR2"
    final = C.classify_one(
        final_record,
        C.LedgerIndex(after_first),
        {},
        0.0,
        2026,
    )

    assert final["bucket"] == "auto"
    assert final["ledger_row_ref"] == 3
    assert final["five_cols"]["计提"] == 150.0
    assert final["five_cols"]["回款明细"] == 30.0
    assert final["five_cols"]["是否结账"] == "是"
    assert final["derived_cols"]["差异"] == -50.0

    final_checked = V.validate({"auto": [final]}, V.read_ledger_rows(after_first))
    assert final_checked["counts"] == {"write": 1, "skip": 0, "conflict": 0}
    after_final = tmp_path / "特殊部分回款最终结清.xlsx"
    A.write_plan(after_first, after_final, final_checked["write"])
    assert A.verify_written(after_final, final_checked["write"]) == []

    values = openpyxl.load_workbook(str(after_final), data_only=True)["明细"]
    assert values.cell(2, 6).value == 100.0
    assert values.cell(2, 8).value == 120.0
    assert values.cell(3, 6).value is None
    assert values.cell(3, 7).value == 150.0
    assert values.cell(3, 8).value == 30.0
    assert values.cell(3, 13).value == -50.0
    formulas = openpyxl.load_workbook(str(after_final), data_only=False)["明细"]
    assert formulas.cell(3, 13).value == "=SUM(F2,F3)-G3"


def test_blank_carry_row_can_split_again_before_delivery_is_fully_received(tmp_path):
    ledger = _ledger(
        tmp_path,
        rows=[("SO1", "SOD1", 100.0, None, None)],
    )
    first = C.classify_one(
        _record(amount=80.0, delivery=150.0, cumulative=80.0),
        C.LedgerIndex(ledger),
        {},
        0.0,
        2026,
    )
    first_checked = V.validate({"auto": [first]}, V.read_ledger_rows(ledger))
    after_first = tmp_path / "特殊部分回款第一次.xlsx"
    A.write_plan(ledger, after_first, first_checked["write"])

    second_record = _record(amount=40.0, delivery=150.0, cumulative=120.0)
    second_record["ar"] = "AR2"
    second = C.classify_one(
        second_record,
        C.LedgerIndex(after_first),
        {},
        0.0,
        2026,
    )
    assert second["bucket"] == "auto"
    assert second["ledger_row_ref"] == 3
    assert second["row_operation"]["receivable_mode"] == "preserve_baseline_blank_carry"
    assert second["row_operation"]["remaining_unreceived"] == 30.0

    second_checked = V.validate({"auto": [second]}, V.read_ledger_rows(after_first))
    assert second_checked["counts"] == {"write": 1, "skip": 0, "conflict": 0}
    after_second = tmp_path / "特殊部分回款第二次.xlsx"
    A.write_plan(after_first, after_second, second_checked["write"])
    assert A.verify_written(after_second, second_checked["write"]) == []

    ws = openpyxl.load_workbook(str(after_second), data_only=True)["明细"]
    assert [ws.cell(row, 6).value for row in (2, 3, 4)] == [100.0, None, None]
    assert [ws.cell(row, 8).value for row in (2, 3, 4)] == [80.0, 40.0, None]
    assert [ws.cell(row, 9).value for row in (2, 3, 4)] == ["是", "是", "否"]


def test_multi_parent_batch_uses_blank_receivables_after_original_baseline(tmp_path):
    ledger = _ledger(
        tmp_path,
        rows=[("SO1", "SOD1", 100.0, None, None)],
    )
    first = _record(amount=80.0, delivery=150.0, cumulative=80.0)
    first["writeoff_sequence_key"] = ["2026-07-22", "HX1", "1", "AR1", "SO1"]
    second = _record(amount=40.0, delivery=150.0, cumulative=120.0)
    second["ar"] = "AR2"
    second["writeoff_sequence_key"] = ["2026-07-22", "HX2", "2", "AR2", "SO1"]

    plan = C.classify_records([first, second], C.LedgerIndex(ledger), {})
    checked = V.validate(plan, V.read_ledger_rows(ledger))

    assert checked["counts"] == {"write": 2, "skip": 0, "conflict": 0}, checked
    out = tmp_path / "特殊分笔回款链.xlsx"
    A.write_plan(ledger, out, checked["write"])
    assert A.verify_written(out, checked["write"]) == []

    ws = openpyxl.load_workbook(str(out), data_only=True)["明细"]
    assert [ws.cell(row, 6).value for row in (2, 3, 4)] == [100.0, None, None]
    assert [ws.cell(row, 8).value for row in (2, 3, 4)] == [80.0, 40.0, None]
    assert [ws.cell(row, 9).value for row in (2, 3, 4)] == ["是", "是", "否"]

    rerun = V.validate(plan, V.read_ledger_rows(out))
    assert rerun["counts"] == {"write": 0, "skip": 2, "conflict": 0}
    comparison = A.build_order_difference(checked["write"], out, hexiao_date="2026-07-22")
    assert comparison["difference_count"] == 0
    assert comparison["matched_count"] == 3


def test_delivery_above_baseline_cumulative_over_delivery_is_held():
    result = C.classify_one(
        _record(amount=40.0, delivery=150.0, cumulative=160.0),
        _synthetic_ledger(receivable=100.0),
        {},
        0.0,
        2026,
    )

    assert result["bucket"] == "hold"
    assert result["code"] == "E5"
    assert "row_operation" not in result


def test_delivery_above_baseline_does_not_close_on_sub_yuan_shortfall():
    result = C.classify_one(
        _record(amount=149.5, delivery=150.0, cumulative=149.5),
        _synthetic_ledger(receivable=100.0),
        {},
        0.0,
        2026,
    )

    assert result["row_operation"]["receivable_mode"] == "preserve_baseline_blank_carry"
    assert result["row_operation"]["remaining_unreceived"] == 0.5
    assert result["five_cols"]["计提"] is None


def test_delivery_above_baseline_does_not_accept_sub_yuan_overpayment():
    result = C.classify_one(
        _record(amount=40.5, delivery=150.0, cumulative=150.5),
        _synthetic_ledger(receivable=100.0),
        {},
        0.0,
        2026,
    )

    assert result["bucket"] == "hold"
    assert result["code"] == "E5"
    assert "row_operation" not in result


def test_delivery_not_above_baseline_keeps_original_conserving_split_mode():
    result = C.classify_one(
        _record(amount=40.0, delivery=100.0, cumulative=40.0),
        _synthetic_ledger(receivable=100.0),
        {},
        0.0,
        2026,
    )

    operation = result["row_operation"]
    assert operation["type"] == "split_below"
    assert "receivable_mode" not in operation
    assert operation["paid_receivable"] == 40.0
    assert operation["unpaid_receivable"] == 60.0


def test_difference_write_skip_and_conflict(tmp_path):
    five = {
        "计提": 110.0, "回款明细": 110.0, "是否结账": "是",
        "收款时间": dt.date(2026, 7, 8), "收款方式": "汇",
    }
    missing = _ledger(
        tmp_path, rows=[("SO1", "SOD1", 100.0, five, None)]
    )
    assert V.check_one(_item(), V.read_ledger_rows(missing))["verdict"] == "write"

    correct = _ledger(
        tmp_path, rows=[("SO1", "SOD1", 100.0, five, -10.0)]
    )
    assert V.check_one(_item(), V.read_ledger_rows(correct))["verdict"] == "skip"

    wrong = _ledger(
        tmp_path, rows=[("SO1", "SOD1", 100.0, five, -9.0)]
    )
    assert V.check_one(_item(), V.read_ledger_rows(wrong))["verdict"] == "conflict"


def test_missing_difference_column_only_blocks_required_difference(tmp_path):
    ledger = _ledger(
        tmp_path,
        with_difference=False,
        rows=[("SO1", "SOD1", 100.0, None, None)],
    )
    rows = V.read_ledger_rows(ledger)
    assert V.check_one(_item(difference=None, amount=100.0), rows)["verdict"] == "write"
    result = V.check_one(_item(), rows)
    assert result["verdict"] == "conflict"
    assert "没有“差异”列" in result["reason"]


def test_ooxml_formula_and_cache_round_trip(tmp_path):
    ledger = _ledger(
        tmp_path, rows=[("SO1", "SOD1", 100.0, None, None)]
    )
    out = tmp_path / "公式缓存.xlsx"
    item = _item()
    A.write_plan(ledger, out, [item])
    assert A.verify_written(out, [item]) == []
    formula_wb = openpyxl.load_workbook(out, data_only=False)
    assert formula_wb["明细"].cell(2, 13).value == "=F2-G2"
    formula_wb.close()
    value_wb = openpyxl.load_workbook(out, data_only=True)
    assert value_wb["明细"].cell(2, 13).value == -10.0
    value_wb.close()


def test_formula_uses_final_row_after_split(tmp_path):
    ledger = _ledger(tmp_path, rows=[
        ("SO_SPLIT", "SOD_SPLIT", 100.0, None, None),
        ("SO1", "SOD1", 100.0, None, None),
    ])
    split = _split_item()
    difference = _item(row=3)
    out = tmp_path / "插行后公式.xlsx"
    A.write_plan(ledger, out, [split, difference])
    assert A.verify_written(out, [split, difference]) == []
    formula_wb = openpyxl.load_workbook(out, data_only=False)
    assert formula_wb["明细"].cell(4, 13).value == "=F4-G4"
    assert formula_wb["明细"].cell(2, 13).value is None
    assert formula_wb["明细"].cell(3, 13).value is None
    formula_wb.close()


def test_worklist_has_three_distinct_difference_columns(tmp_path):
    item = _item()
    result = {
        "auto": [item], "hold": [], "exception": [],
        "counts": {"auto": 1, "hold": 0, "exception": 0, "total": 1},
        "e_code_dist": {"OK": 1}, "payment_count": 1,
    }
    checked = {
        "write": [{**item, "_check": {"verdict": "write", "reason": "可写"}}],
        "skip": [], "conflict": [],
        "counts": {"write": 1, "skip": 0, "conflict": 0},
    }
    out = tmp_path / "核销日清.xlsx"
    W.build_workbook(result, checked, out)
    ws = openpyxl.load_workbook(out)["今日清单"]
    headers = [c.value for c in ws[1]]
    assert "应填_业务值差异" in headers
    assert "当前_业务值差异" in headers
    assert "当前值与计划值的比较差异" in headers
