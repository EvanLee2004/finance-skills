# -*- coding: utf-8 -*-
"""同一 SO 多 SOD：全部结清后才计提，并补填历史已结清行。"""
import datetime as dt
import json

import openpyxl

import apply_to_copy as A
import classify_hexiao as C
import validate_plan as V


def _synthetic(rows):
    so_index = {}
    sod_index = {}
    for row_no, row in rows.items():
        so_index.setdefault(row["so"], []).append(row_no)
        sod_index.setdefault(row["sod"], []).append(row_no)
    return C.LedgerIndex(synthetic={"so": so_index, "sod": sod_index, "rows": rows})


def _record(sod, amount, deliveries):
    return {
        "ar": f"AR_{sod}",
        "so": "SO1",
        "sod": sod,
        "amount_orig": amount,
        "amount_local": amount,
        "deliver_local": deliveries[sod],
        "so_delivery_local": sum(deliveries.values()),
        "all_sods": sorted(deliveries),
        "sod_delivery_local": deliveries,
        "currency": "人民币CNY",
        "status": "手动核销",
        "customer": "测试客户",
        "hexiao_date": dt.date(2026, 8, 11),
        "shoukuan_date": dt.date(2026, 8, 11),
    }


def test_first_settled_sod_defers_accrual_until_siblings_settle():
    ledger = _synthetic({
        2: {"so": "SO1", "sod": "SOD1", "yingshou": 100.0, "jiezhang": "否"},
        3: {"so": "SO1", "sod": "SOD2", "yingshou": 200.0, "jiezhang": "否"},
    })

    result = C.classify_records([_record("SOD1", 100.0, {"SOD1": 100.0, "SOD2": 200.0})], ledger)
    item = result["auto"][0]

    assert item["five_cols"]["是否结账"] == "是"
    assert item["five_cols"]["计提"] is None
    assert item["so_accrual_audit"]["unsettled_sods"] == ["SOD2"]
    assert "W_SO_ACCRUAL_DEFERRED" in item["warning_codes"]


def test_last_sod_releases_current_accrual_and_backfills_prior_sod():
    ledger = _synthetic({
        2: {
            "so": "SO1", "sod": "SOD1", "yingshou": 100.0,
            "jiti": None, "huikuan": 100.0, "jiezhang": "是",
            "shoukuan_time": dt.date(2026, 8, 10), "shoukuan_way": "汇",
        },
        3: {"so": "SO1", "sod": "SOD2", "yingshou": 200.0, "jiezhang": "否"},
    })

    result = C.classify_records([_record("SOD2", 200.0, {"SOD1": 100.0, "SOD2": 200.0})], ledger)
    item = result["auto"][0]

    assert item["five_cols"]["计提"] == 200.0
    backfill = item["so_accrual_backfills"][0]
    assert backfill["so"] == "SO1"
    assert backfill["sod"] == "SOD1"
    assert backfill["ledger_row_ref"] == 2
    assert backfill["business_rows"] == [2]
    assert backfill["accrual"] == 100.0
    assert backfill["current_batch_sods"] == ["SOD2"]
    assert backfill["all_sods"] == ["SOD1", "SOD2"]


def test_backfill_uses_only_last_settled_business_row_of_split_sod():
    ledger = _synthetic({
        2: {"so": "SO1", "sod": "SOD1", "yingshou": 40.0, "huikuan": 40.0, "jiezhang": "是"},
        3: {"so": "SO1", "sod": "SOD1", "yingshou": 60.0, "huikuan": 60.0, "jiezhang": "是"},
        4: {"so": "SO1", "sod": "SOD2", "yingshou": 200.0, "jiezhang": "否"},
    })

    result = C.classify_records([_record("SOD2", 200.0, {"SOD1": 100.0, "SOD2": 200.0})], ledger)
    backfill = result["auto"][0]["so_accrual_backfills"][0]

    assert backfill["ledger_row_ref"] == 3
    assert backfill["business_rows"] == [2, 3]
    assert backfill["accrual"] == 100.0


def _workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append([
        "部门", "销售人员", "客户名称", "单号", "新智云单号", "应收金额",
        "计提金额", "回款明细", "是否结账（是/否）", "收款时间",
        "收款方式(支/汇/现)", "实收金额", "差异",
    ])
    ws.append(["部", "人", "客", "AB", "SO1", 100.0, None, 100.0, "是", "2026-08-10", "汇", "SOD1", None])
    ws.append(["部", "人", "客", "AB", "SO1", 200.0, None, None, "否", None, None, "SOD2", None])
    wb.save(path)


def test_validate_apply_readback_and_rerun_are_idempotent(tmp_path):
    ledger_path = tmp_path / "盈亏.xlsx"
    output_path = tmp_path / "盈亏_已写.xlsx"
    _workbook(ledger_path)
    deliveries = {"SOD1": 100.0, "SOD2": 200.0}

    plan = C.classify_records(
        [_record("SOD2", 200.0, deliveries)], C.LedgerIndex(ledger_path)
    )
    checked = V.validate(plan, V.read_ledger_rows(ledger_path), ledger_path=ledger_path)

    assert checked["counts"] == {"write": 1, "skip": 0, "conflict": 0}
    backfill = checked["write"][0]["so_accrual_backfills"][0]
    assert backfill["_check"]["verdict"] == "write"

    A.write_plan(ledger_path, output_path, checked["write"])
    assert A.verify_written(output_path, checked["write"]) == []
    rows = V.read_ledger_rows(output_path)
    assert rows[2]["计提"] == 100.0
    assert rows[3]["计提"] == 200.0

    rerun_plan = C.classify_records(
        [_record("SOD2", 200.0, deliveries)], C.LedgerIndex(output_path)
    )
    assert not rerun_plan["auto"][0].get("so_accrual_backfills")
    rerun_checked = V.validate(rerun_plan, V.read_ledger_rows(output_path))
    assert rerun_checked["counts"] == {"write": 0, "skip": 1, "conflict": 0}


def _result_with_backfill():
    notice_item = {
        "so": "SO1",
        "sod": "SOD2",
        "ledger_year": 2026,
        "so_accrual_audit": {
            "all_settled": True,
            "all_sods": ["SOD1", "SOD2"],
            "current_batch_sods": ["SOD2"],
        },
        "so_accrual_backfills": [{
            "so": "SO1", "sod": "SOD1", "ledger_row_ref": 2,
            "accrual": 100.0, "difference": None, "current_accrual": None,
            "historical_receipt_time": "2026-07-20", "ledger_year": 2026,
        }],
    }
    return {"auto": [notice_item], "hold": [], "exception": []}


def test_cross_month_notice_uses_prior_result_real_writeoff_date(tmp_path):
    out = tmp_path / "04_产出"
    out.mkdir()
    (out / "判定结果_20260722.json").write_text(json.dumps({
        "hexiao_date": "2026-07-22",
        "auto": [{"so": "SO1", "sod": "SOD1"}],
    }, ensure_ascii=False), encoding="utf-8")
    result = _result_with_backfill()

    notices = C.annotate_cross_month_accruals(
        result, tmp_path, dt.date(2026, 8, 12)
    )

    assert len(notices) == 1
    assert notices[0]["historical_hexiao_dates"] == ["2026-07-22"]
    assert notices[0]["historical_hexiao_months"] == ["2026-07"]
    assert notices[0]["cross_month_status"] == "是"
    assert notices[0]["history_source"] == "历史核销日清"
    nested = result["auto"][0]["so_accrual_backfills"][0]
    assert nested["cross_month_accrual_notice"]["historical_sod"] == "SOD1"


def test_same_month_prior_result_does_not_create_cross_month_notice(tmp_path):
    out = tmp_path / "04_产出"
    out.mkdir()
    (out / "判定结果_20260805.json").write_text(json.dumps({
        "hexiao_date": "2026-08-05",
        "auto": [{"so": "SO1", "sod": "SOD1"}],
    }, ensure_ascii=False), encoding="utf-8")
    result = _result_with_backfill()

    notices = C.annotate_cross_month_accruals(
        result, tmp_path, dt.date(2026, 8, 12)
    )

    assert notices == []


def test_missing_prior_result_uses_ledger_receipt_time_for_cross_month(tmp_path):
    (tmp_path / "04_产出").mkdir()
    result = _result_with_backfill()

    notices = C.annotate_cross_month_accruals(
        result, tmp_path, dt.date(2026, 8, 12)
    )

    assert len(notices) == 1
    assert notices[0]["historical_hexiao_dates"] == ["2026-07-20"]
    assert notices[0]["historical_hexiao_months"] == ["2026-07"]
    assert notices[0]["cross_month_status"] == "是"
    assert notices[0]["history_source"] == "盈亏核算表收款时间"


def test_missing_prior_result_same_month_receipt_is_not_cross_month(tmp_path):
    (tmp_path / "04_产出").mkdir()
    result = _result_with_backfill()
    result["auto"][0]["so_accrual_backfills"][0]["historical_receipt_time"] = "2026-08-05"

    notices = C.annotate_cross_month_accruals(
        result, tmp_path, dt.date(2026, 8, 12)
    )

    assert notices == []


def test_missing_prior_result_and_receipt_time_stays_pending(tmp_path):
    (tmp_path / "04_产出").mkdir()
    result = _result_with_backfill()
    result["auto"][0]["so_accrual_backfills"][0]["historical_receipt_time"] = None

    notices = C.annotate_cross_month_accruals(
        result, tmp_path, dt.date(2026, 8, 12)
    )

    assert len(notices) == 1
    assert notices[0]["cross_month_status"] == "待确认"
    assert notices[0]["history_source"] == "未找到历史日清及盈亏收款时间"
