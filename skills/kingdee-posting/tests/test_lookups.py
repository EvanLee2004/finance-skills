#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""预处理查找：金蝶往来定科目、回款优先、双销售 hold。不访问网络。"""
from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

import importlib.util

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


lookups = _load("kingdee_posting_lookups_unit", SCRIPTS / "lookups.py")
zhiyun_api = _load("kingdee_posting_zhiyun_unit", SCRIPTS / "zhiyun_api.py")


def _box(raw):
    return lookups.box_from_dict(raw, ["113101", "113102", "113103"], {"于占国": "15"})


def test_single_line_to_accounts():
    box = _box({"ar_balance": [{"customer_code": "1001", "account": "113103", "balance": "1"}]})
    ar, rev, why = lookups.resolve_ar("1001", "2026-08-01", box)
    assert (ar, rev, why) == ("113103", "510103", "")


def test_prior_month_balance_books_without_invoice_month_debit():
    box = _box({"ar_balance": [{"customer_code": "1001", "account": "113103", "balance": "80"}]})
    ar, rev, why = lookups.resolve_ar("1001", "2026-09-07", box)
    assert (ar, rev, why) == ("113103", "510103", "")
    assert why == ""


def test_unknown_customer_has_no_line():
    box = _box({})
    ar, rev, why = lookups.resolve_ar("1001", "2026-08-01", box)
    assert ar == ""
    assert "金蝶往来" in why
    assert "业务线" not in why


def test_zero_balance_falls_back_to_period_debit():
    box = _box(
        {
            "ar_balance": [{"customer_code": "1001", "account": "113103", "balance": "0"}],
            "period_debit": [
                {"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "80"},
            ],
        }
    )
    ar, rev, why = lookups.resolve_ar("1001", "2026-08-26", box)
    assert (ar, rev, why) == ("113103", "510103", "")
    assert "业务线" not in why


def test_zero_balance_without_period_debit_holds_kingdee_ar():
    box = _box({"ar_balance": [{"customer_code": "1001", "account": "113103", "balance": "0"}]})
    ar, rev, why = lookups.resolve_ar("1001", "2026-08-26", box)
    assert ar == ""
    assert rev == ""
    assert "金蝶往来" in why
    assert "业务线" not in why


def test_multi_line_picks_larger_period_debit():
    box = _box(
        {
            "ar_balance": [
                {"customer_code": "1001", "account": "113103", "balance": "1"},
                {"customer_code": "1001", "account": "113102", "balance": "1"},
            ],
            "period_debit": [
                {"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "80"},
                {"customer_code": "1001", "account": "113102", "period": "2026-08", "debit": "20"},
            ],
        }
    )
    ar, rev, why = lookups.resolve_ar("1001", "2026-08-15", box)
    assert (ar, rev, why) == ("113103", "510103", "")


def test_multi_line_without_debit_holds():
    box = _box(
        {
            "ar_balance": [
                {"customer_code": "1001", "account": "113103", "balance": "1"},
                {"customer_code": "1001", "account": "113102", "balance": "1"},
            ]
        }
    )
    ar, _rev, why = lookups.resolve_ar("1001", "2026-08-01", box)
    assert ar == ""
    assert "本期借方" in why
    assert "业务线" not in why


def test_multi_line_equal_debit_holds():
    box = _box(
        {
            "ar_balance": [
                {"customer_code": "1001", "account": "113103", "balance": "1"},
                {"customer_code": "1001", "account": "113102", "balance": "1"},
            ],
            "period_debit": [
                {"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "10"},
                {"customer_code": "1001", "account": "113102", "period": "2026-08", "debit": "10"},
            ],
        }
    )
    ar, _rev, why = lookups.resolve_ar("1001", "2026-08-01", box)
    assert ar == ""
    assert "本期借方" in why
    assert "业务线" not in why


def test_receipt_sales_beats_order_sales():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "甲科技有限公司", "date": "2026-08-01", "amount": "10", "sales": ["于占国"]},
            ],
            "order_sales": {"甲科技有限公司": ["陈霞"]},
        }
    )
    name, why = lookups.resolve_sales("甲科技有限公司", "2026-08-01", Decimal("10.00"), box)
    assert (name, why) == ("于占国", "")


def test_order_sales_when_receipt_missing():
    box = _box({"order_sales": {"甲科技有限公司": ["陈霞"]}})
    name, why = lookups.resolve_sales("甲科技有限公司", "2026-08-01", "10", box)
    assert (name, why) == ("陈霞", "")


def test_dual_order_sales_holds_for_sijia():
    box = _box({"order_sales": {"甲科技有限公司": ["于占国", "陈霞"]}})
    name, why = lookups.resolve_sales("甲科技有限公司", "2026-08-01", "10", box)
    assert name == ""
    assert "斯佳" in why


def test_receipt_sales_peel_company_suffix():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "甲科技", "date": "2026-08-01", "amount": "10", "sales": ["于占国"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("甲科技有限公司", "2026-08-01", Decimal("10.00"), box)
    assert (name, why) == ("于占国", "")


def test_receipt_sales_peel_collection_account():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "甲科技有限公司", "date": "2026-08-01", "amount": "10", "sales": ["于占国"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("甲科技有限公司第二收缴户", "2026-08-01", Decimal("10.00"), box)
    assert (name, why) == ("于占国", "")


def test_receipt_sales_date_amount_when_bank_name_differs():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "开票客户甲", "date": "2026-09-01", "amount": "10", "sales": ["于占国"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("银行收缴户乙", "2026-09-01", Decimal("10.00"), box)
    assert (name, why) == ("于占国", "")


def test_receipt_sales_date_amount_two_sales_holds():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "开票客户甲", "date": "2026-09-01", "amount": "10", "sales": ["于占国"]},
                {"customer": "开票客户乙", "date": "2026-09-01", "amount": "10", "sales": ["陈霞"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("银行收缴户", "2026-09-01", Decimal("10.00"), box)
    assert name == ""
    assert "不唯一" in why


def test_named_receipt_without_sales_does_not_steal_other_customer():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "甲科技有限公司", "date": "2026-09-01", "amount": "10", "sales": []},
                {"customer": "乙科技有限公司", "date": "2026-09-01", "amount": "10", "sales": ["陈霞"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("甲科技有限公司", "2026-09-01", Decimal("10.00"), box)
    assert name == ""
    assert why == "找不到销售"


def test_receipt_history_unique_when_this_payment_missing():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "开票客户甲", "date": "2026-07-01", "amount": "80", "sales": ["于占国"]},
                {"customer": "开票客户甲", "date": "2026-08-01", "amount": "90", "sales": ["于占国"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("开票客户甲", "2026-09-03", Decimal("12.00"), box)
    assert (name, why) == ("于占国", "")


def test_receipt_history_two_sales_holds():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "开票客户甲", "date": "2026-07-01", "amount": "80", "sales": ["于占国"]},
                {"customer": "开票客户甲", "date": "2026-08-01", "amount": "90", "sales": ["陈霞"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("开票客户甲", "2026-09-03", Decimal("12.00"), box)
    assert name == ""
    assert "不唯一" in why


def test_this_payment_beats_receipt_history():
    box = _box(
        {
            "receipt_sales": [
                {"customer": "开票客户甲", "date": "2026-07-01", "amount": "80", "sales": ["陈霞"]},
                {"customer": "开票客户甲", "date": "2026-09-03", "amount": "12", "sales": ["于占国"]},
            ]
        }
    )
    name, why = lookups.resolve_sales("开票客户甲", "2026-09-03", Decimal("12.00"), box)
    assert (name, why) == ("于占国", "")


def test_order_sales_unique_contain():
    box = _box({"order_sales": {"甲科技有限公司": ["陈霞"]}})
    name, why = lookups.resolve_sales("甲科技有限公司北京分公司", "2026-08-01", "10", box)
    assert (name, why) == ("陈霞", "")


def test_order_sales_contain_two_customers_holds():
    box = _box(
        {
            "order_sales": {
                "某某中心一部": ["陈霞"],
                "某某中心二部": ["于占国"],
            }
        }
    )
    name, why = lookups.resolve_sales("某某中心", "2026-08-01", "10", box)
    assert name == ""
    assert "斯佳" in why


def test_records_to_lookups_merges_lines_and_receipts():
    got = zhiyun_api.records_to_lookups(
        [
            {"客户": "甲科技有限公司", "业务线": "ICT", "销售": "于占国"},
            {"客户": "甲科技有限公司", "业务线": "游戏综合本地化", "销售": "陈霞"},
        ],
        [
            {"客户": "甲科技有限公司", "到账日期": "2026-08-01", "到账金额/本币": "10", "销售": "于占国"},
        ],
    )
    assert got["customer_lines"]["甲科技有限公司"] == ["ICT", "游戏综合本地化"]
    assert got["order_sales"]["甲科技有限公司"] == ["于占国", "陈霞"]
    assert got["receipt_sales"][0]["sales"] == ["于占国"]


def test_records_to_lookups_keeps_invoice_customer():
    got = zhiyun_api.records_to_lookups(
        [{"客户": "集团甲", "开票客户": "子公司乙", "销售": "于占国"}],
        [
            {
                "客户": "集团甲",
                "开票客户": "子公司乙",
                "到账日期": "2026-08-01",
                "到账金额/本币": "10",
                "销售": "于占国",
            }
        ],
    )
    assert got["order_sales"]["集团甲"] == ["于占国"]
    assert got["order_sales"]["子公司乙"] == ["于占国"]
    names = {item["customer"] for item in got["receipt_sales"]}
    assert names == {"集团甲", "子公司乙"}


def test_zhiyun_missing_credentials(tmp_path, monkeypatch):
    monkeypatch.setenv("ZHIYUN_LOCAL_JSON", str(tmp_path / "nope.json"))
    loaded = zhiyun_api.try_load_lookups()
    assert loaded["ok"] is False
    assert loaded["missing_credentials"] is True


def _assist(*rows):
    return lookups.assist_rows_from_dicts(list(rows))


def test_pick_assist_unique_1131_books_income_tail():
    rows = _assist(
        {
            "period": "202608",
            "customer_code": "4146",
            "customer_name": "甲",
            "account": "113105",
            "ending_debit": "1",
            "ytd_debit": "1",
        }
    )
    ar, rev, why = lookups.pick_assist_account("4146", "2026-09-07", rows)
    assert (ar, rev, why) == ("113105", "510105", "")


def test_pick_assist_two_ending_holds_no_silent_max():
    rows = _assist(
        {
            "period": "202608",
            "customer_code": "1",
            "customer_name": "甲",
            "account": "113103",
            "ending_debit": "80",
            "ytd_debit": "80",
        },
        {
            "period": "202608",
            "customer_code": "1",
            "customer_name": "甲",
            "account": "113102",
            "ending_debit": "20",
            "ytd_debit": "20",
        },
    )
    ar, _rev, why = lookups.pick_assist_account("1", "2026-09-07", rows)
    assert ar == ""
    assert "多条" in why
    assert "本期借方" not in why
    assert lookups.list_assist_accounts("1", "2026-09-07", rows) == ["113103", "113102"]


def test_pick_assist_preferred_tencent_uses_07():
    rows = _assist(
        {
            "period": "202608",
            "customer_code": "3843",
            "customer_name": "腾讯科技（深圳）有限公司",
            "account": "113101",
            "ending_debit": "10",
        },
        {
            "period": "202608",
            "customer_code": "3843",
            "customer_name": "腾讯科技（深圳）有限公司",
            "account": "113103",
            "ending_debit": "20",
        },
        {
            "period": "202608",
            "customer_code": "3843",
            "customer_name": "腾讯科技（深圳）有限公司",
            "account": "113107",
            "ending_debit": "80",
        },
    )
    ar, rev, why = lookups.pick_assist_account("3843", "2026-09-03", rows, preferred_ar="07")
    assert (ar, rev, why) == ("113107", "510107", "")
    ar2, _rev, why2 = lookups.pick_assist_account("3843", "2026-09-03", rows)
    assert ar2 == ""
    assert "多条" in why2


def test_pick_assist_preferred_missing_from_table_still_holds():
    rows = _assist(
        {"period": "202608", "customer_code": "1", "account": "113103", "ending_debit": "1"},
        {"period": "202608", "customer_code": "1", "account": "113101", "ending_debit": "1"},
    )
    ar, _rev, why = lookups.pick_assist_account("1", "2026-09-07", rows, preferred_ar="113107")
    assert ar == ""
    assert "多条" in why


def test_pick_assist_unique_ignores_preferred():
    rows = _assist(
        {"period": "202608", "customer_code": "1", "account": "113103", "ending_debit": "1"},
    )
    ar, rev, why = lookups.pick_assist_account("1", "2026-09-07", rows, preferred_ar="07")
    assert (ar, rev, why) == ("113103", "510103", "")


def test_pick_assist_missing_asks_new():
    ar, _rev, why = lookups.pick_assist_account("1", "2026-09-07", [])
    assert ar == ""
    assert "申请人科目" in why


def test_pick_assist_empty_uses_preferred_for_first_invoice():
    ar, rev, why = lookups.pick_assist_account("4957", "2026-09-03", [], preferred_ar="03")
    assert (ar, rev, why) == ("113103", "510103", "")


def test_pick_assist_empty_ending_unique_ytd():
    rows = _assist(
        {"period": "202608", "customer_code": "1", "name": "甲", "account": "113103", "ytd_debit": "9"},
        {"period": "202608", "customer_code": "1", "name": "甲", "account": "113102"},
    )
    ar, rev, why = lookups.pick_assist_account("1", "2026-09-07", rows)
    assert (ar, rev, why) == ("113103", "510103", "")


def test_pick_assist_uses_prev_completed_month_not_future():
    rows = _assist(
        {"period": "202607", "customer_code": "1", "name": "甲", "account": "113103", "ending_debit": "1"},
        {"period": "202609", "customer_code": "1", "name": "甲", "account": "113102", "ending_debit": "9"},
        {"period": "202609", "customer_code": "1", "name": "甲", "account": "113103", "ending_debit": "9"},
    )
    ar, rev, why = lookups.pick_assist_account("1", "2026-09-07", rows)
    assert (ar, rev, why) == ("113103", "510103", "")
    assert lookups.prev_completed_month("2026-09-07") == "202608"
