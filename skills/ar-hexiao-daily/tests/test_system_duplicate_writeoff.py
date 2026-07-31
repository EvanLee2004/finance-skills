from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path

import classify_hexiao as C
import fetch_zhiyun as F
import validate_plan as V
import writeoff_duplicate_audit as W


DAY = dt.date(2026, 7, 31)


def payment(ar="AR1", amount=100, *, local=None, currency="CNY"):
    return {
        "ar": ar,
        "amount_orig": amount,
        "amount_local": local,
        "currency": currency,
        "hexiao_date": DAY,
        "arrival_date": DAY,
        "orders": [{"so": "SO1", "deliver": amount, "currency": currency}],
        "writeoffs": {},
        "writeoffs_local": {},
        "cumulative_writeoffs": {},
        "cumulative_writeoffs_local": {},
        "_source_meta": {"historical_detail_rows": 0, "recovered_deliveries": 0},
    }


def row(record_id, amount=100, *, ar="AR1", so="SO1", local=None,
        currency="CNY", day=DAY, revoked=""):
    return {
        "record_id": record_id,
        "rowid": f"row-{record_id}",
        "ar": ar,
        "date": day,
        "so": so,
        "amount": amount,
        "amount_local": local,
        "currency": currency,
        "revoked": revoked,
        "source": "核销明细_20260731.xlsx",
        "snapshot_date": DAY,
    }


def audit(parent, rows):
    return W.audit_parent_writeoffs(parent, rows)


def test_two_distinct_ids_same_so_amount_recover_one_logical():
    logical, info = audit(payment(), [row("HX1"), row("HX2")])
    assert info["delta_raw"] == -100
    assert info["delta_dedup"] == 0
    assert info["status"] == "recovered"
    assert [item["record_id"] for item in logical] == ["HX1"]


def test_four_distinct_ids_only_count_once_for_h_and_r():
    p = payment()
    raw = [row(f"HX{i}") for i in range(1, 5)]
    current, audits = C.reconcile_writeoff_details([p], {"AR1": p}, raw, DAY)
    assert len(current) == 1
    assert p["writeoffs"] == {"SO1": 100.0}
    assert p["cumulative_writeoffs"] == {"SO1": 100.0}
    assert audits["AR1"]["ignored_record_count"] == 3


def test_parent_400_four_equal_records_is_normal_not_collapsed():
    logical, info = audit(payment(amount=400), [row(f"HX{i}") for i in range(4)])
    assert info["delta_raw"] == 0
    assert info["status"] == "normal"
    assert len(logical) == 4
    assert not info["duplicate_groups"]


def test_negative_099_is_tolerance_without_business_dedup():
    logical, info = audit(payment(amount=99.01), [row("HX1"), row("HX2", amount=0)])
    assert info["delta_raw"] == -0.99
    assert info["status"] == "tolerance"
    assert len(logical) == 2


def test_negative_100_is_tolerance():
    logical, info = audit(payment(amount=99), [row("HX1")])
    assert info["delta_raw"] == -1
    assert info["status"] == "tolerance"
    assert len(logical) == 1


def test_negative_101_exact_duplicate_recovers_within_tolerance():
    logical, info = audit(
        payment(amount=98.99),
        [row("HX1", amount=50), row("HX2", amount=50)],
    )
    assert info["delta_raw"] == -1.01
    assert info["delta_dedup"] == 48.99
    assert info["status"] == "recovered"
    assert len(logical) == 1


def test_over_writeoff_without_exact_duplicate_is_unresolved():
    logical, info = audit(
        payment(),
        [row("HX1", amount=60), row("HX2", amount=50)],
    )
    assert not logical
    assert info["status"] == "unresolved"


def test_duplicate_fold_still_over_is_unresolved():
    logical, info = audit(
        payment(),
        [row("HX1", amount=80), row("HX2", amount=80), row("HX3", amount=30)],
    )
    assert not logical
    assert info["delta_dedup"] == -10
    assert info["status"] == "unresolved"


def test_same_so_different_amounts_are_not_folded():
    logical, info = audit(
        payment(),
        [row("HX1", amount=70), row("HX2", amount=40)],
    )
    assert not logical
    assert not info["duplicate_groups"]


def test_different_so_same_amounts_are_not_folded():
    logical, info = audit(
        payment(),
        [row("HX1", amount=100, so="SO1"), row("HX2", amount=100, so="SO2")],
    )
    assert not logical
    assert not info["duplicate_groups"]


def test_same_record_id_across_snapshots_is_one_physical_not_system_duplicate():
    earlier = row("HX1")
    later = {**row("HX1"), "source": "核销明细_20260801.xlsx",
             "snapshot_date": dt.date(2026, 8, 1)}
    logical, info = audit(payment(), [earlier, later])
    assert len(logical) == 1
    assert info["physical_snapshot_duplicate_count"] == 1
    assert info["status"] == "normal"
    assert not info["duplicate_groups"]


def test_different_ids_identical_fields_survive_raw_stage_until_over():
    rows = [row("HX1"), row("HX2")]
    normal, normal_info = audit(payment(amount=200), rows)
    recovered, recovered_info = audit(payment(amount=100), rows)
    assert len(normal) == 2 and normal_info["status"] == "normal"
    assert len(recovered) == 1 and recovered_info["status"] == "recovered"


def test_revoked_record_is_excluded_from_all_totals():
    logical, info = audit(payment(), [row("HX1"), row("HX2", revoked="是")])
    assert len(logical) == 1
    assert info["raw_total"] == 100
    assert info["revoked_count"] == 1


def test_recovered_duplicate_changes_target_h_and_cumulative_r_together():
    p = payment()
    raw = [row("HX1"), row("HX2")]
    C.reconcile_writeoff_details([p], {"AR1": p}, raw, DAY)
    assert p["writeoffs"]["SO1"] == 100
    assert p["cumulative_writeoffs"]["SO1"] == 100


def test_same_so_across_parents_is_not_cross_parent_folded_but_r_combines():
    p1 = payment("AR1", 100)
    p2 = payment("AR2", 50)
    p2["orders"] = [{"so": "SO1", "deliver": 150, "currency": "CNY"}]
    raw = [
        row("HX1", 100, ar="AR1"),
        row("HX2", 50, ar="AR2"),
    ]
    _, audits = C.reconcile_writeoff_details(
        [p1, p2], {"AR1": p1, "AR2": p2}, raw, DAY
    )
    assert audits["AR1"]["status"] == "normal"
    assert audits["AR2"]["status"] == "normal"
    assert p1["cumulative_writeoffs"]["SO1"] == 150
    assert p2["cumulative_writeoffs"]["SO1"] == 150


def test_local_amounts_are_preferred_for_comparison():
    parent = payment(amount=10, local=100, currency="USD")
    logical, info = audit(
        parent,
        [row("HX1", amount=9, local=100, currency="USD"),
         row("HX2", amount=9, local=100, currency="USD")],
    )
    assert info["comparison_basis"] == "local"
    assert info["status"] == "recovered"
    assert len(logical) == 1


def test_original_amounts_used_only_when_currency_matches():
    logical, info = audit(
        payment(amount=100, currency="USD"),
        [row("HX1", amount=100, currency="USD")],
    )
    assert info["comparison_basis"] == "original"
    assert info["status"] == "normal"
    assert len(logical) == 1


def test_incomparable_currency_or_missing_record_id_overage_is_unresolved():
    _, currency_info = audit(
        payment(amount=100, currency="USD"),
        [row("HX1", amount=200, currency="EUR")],
    )
    _, missing_id_info = audit(payment(), [row("", amount=200)])
    assert currency_info["status"] == "unresolved"
    assert missing_id_info["status"] == "unresolved"


def test_every_raw_row_has_explicit_disposition():
    rows = [
        row("HX1"),
        {**row("HX1"), "snapshot_date": dt.date(2026, 8, 1)},
        row("HX2", revoked="是"),
    ]
    _, info = audit(payment(), rows)
    assert len(info["records"]) == 3
    assert all(item["disposition"] in {
        "kept", "physical_snapshot_duplicate", "system_duplicate_ignored", "revoked"
    } for item in info["records"])


def test_unresolved_parent_cannot_enter_auto_and_audit_tamper_conflicts():
    recovered_rows = [row("HX1"), row("HX2")]
    _, recovered = audit(payment(), recovered_rows)
    audits = {"AR1": recovered}
    plan = {
        "duplicate_writeoff_audits": audits,
        "duplicate_writeoff_audit_sha256": W.audit_fingerprint(audits),
        "auto": [{
            "ar": "AR1", "so": "SO1", "sod": "SOD1", "ledger_row_ref": 2,
            "five_cols": {"计提": 100, "回款明细": 100, "是否结账": "是",
                          "收款时间": "2026-07-31", "收款方式": "汇"},
            "warning_codes": ["W_SYSTEM_DUPLICATE_WRITEOFF_COLLAPSED"],
            "duplicate_writeoff_audit": recovered,
        }],
    }
    rows = {2: {"SO": "SO1", "SOD": "SOD1", "计提": None, "回款明细": None,
                "是否结账": "否", "收款时间": None, "收款方式": None,
                "差异": None, "_差异列存在": True, "应收金额": 100}}
    assert V.validate(plan, rows)["counts"]["write"] == 1
    plan["duplicate_writeoff_audits"]["AR1"]["logical_total"] = 999
    checked = V.validate(plan, rows)
    assert checked["counts"]["conflict"] == 1
    assert "指纹不一致" in checked["conflict"][0]["_check"]["reason"]


def test_actual_unresolved_parent_is_rejected_if_hand_edited_into_auto():
    _, unresolved = audit(
        payment(amount=100),
        [row("HX1", amount=80), row("HX2", amount=30)],
    )
    assert unresolved["status"] == "unresolved"
    audits = {"AR1": unresolved}
    plan = {
        "duplicate_writeoff_audits": audits,
        "duplicate_writeoff_audit_sha256": W.audit_fingerprint(audits),
        "auto": [{
            "ar": "AR1", "so": "SO1", "sod": "SOD1", "ledger_row_ref": 2,
            "five_cols": {"计提": 100, "回款明细": 100, "是否结账": "是",
                          "收款时间": "2026-07-31", "收款方式": "汇"},
            "warning_codes": [],
            "duplicate_writeoff_audit": unresolved,
        }],
    }
    rows = {2: {"SO": "SO1", "SOD": "SOD1", "计提": None, "回款明细": None,
                "是否结账": "否", "收款时间": None, "收款方式": None,
                "差异": None, "_差异列存在": True, "应收金额": 100}}
    checked = V.validate(plan, rows)
    assert checked["counts"]["conflict"] == 1
    assert "未解决" in checked["conflict"][0]["_check"]["reason"]


def test_fetch_contract_exposes_record_identity_and_new_version():
    assert F.MINGXI_COLS[0] == "核销记录NUM"
    assert F.EXPORT_SCHEMA_VERSION == "2026-07-31-writeoff-record-identity-v1"
