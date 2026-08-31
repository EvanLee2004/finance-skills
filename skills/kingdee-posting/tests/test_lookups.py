#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""预处理查找：业务线、回款优先、双销售 hold。不访问网络。"""
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
    line_accounts = {
        "ICT": "113103",
        "游戏综合本地化": "113102",
        "视频字幕": "113101",
    }
    return lookups.box_from_dict(raw, line_accounts, {"于占国": "15"})


def test_single_line_to_accounts():
    box = _box({"customer_lines": {"甲科技有限公司": ["ICT"]}})
    ar, rev, why = lookups.resolve_ar("甲科技有限公司", "2026-08-01", "1001", box)
    assert (ar, rev, why) == ("113103", "510103", "")


def test_unknown_customer_has_no_line():
    box = _box({"customer_lines": {}})
    ar, rev, why = lookups.resolve_ar("甲科技有限公司", "2026-08-01", "1001", box)
    assert ar == ""
    assert why == "客户无业务线"


def test_multi_line_picks_larger_period_debit():
    box = _box(
        {
            "customer_lines": {"甲科技有限公司": ["ICT", "游戏综合本地化"]},
            "period_debit": [
                {"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "80"},
                {"customer_code": "1001", "account": "113102", "period": "2026-08", "debit": "20"},
            ],
        }
    )
    ar, rev, why = lookups.resolve_ar("甲科技有限公司", "2026-08-15", "1001", box)
    assert (ar, rev, why) == ("113103", "510103", "")


def test_multi_line_without_debit_holds():
    box = _box({"customer_lines": {"甲科技有限公司": ["ICT", "游戏综合本地化"]}})
    ar, _rev, why = lookups.resolve_ar("甲科技有限公司", "2026-08-01", "1001", box)
    assert ar == ""
    assert why == "多业务线且本期借方不可用"


def test_multi_line_equal_debit_holds():
    box = _box(
        {
            "customer_lines": {"甲科技有限公司": ["ICT", "游戏综合本地化"]},
            "period_debit": [
                {"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "10"},
                {"customer_code": "1001", "account": "113102", "period": "2026-08", "debit": "10"},
            ],
        }
    )
    ar, _rev, why = lookups.resolve_ar("甲科技有限公司", "2026-08-01", "1001", box)
    assert ar == ""
    assert why == "多业务线本期借方不唯一"


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


def test_zhiyun_missing_credentials(tmp_path, monkeypatch):
    monkeypatch.setenv("ZHIYUN_LOCAL_JSON", str(tmp_path / "nope.json"))
    loaded = zhiyun_api.try_load_lookups()
    assert loaded["ok"] is False
    assert loaded["missing_credentials"] is True
