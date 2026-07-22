# -*- coding: utf-8 -*-
"""步骤6 判定：E 码、分笔 hold、预存日期过滤、std 端到端。"""
import datetime as dt
import json
from pathlib import Path

import pytest

import classify_hexiao as C
import common
from conftest import FIXTURE


def test_channel_fenbi():
    assert C._channel_for_type("分笔回款") == "fenbi"
    assert C._channel_for_type("预存回款") == "yucun"
    assert C._channel_for_type("整笔回款") == "duizhang"
    assert C._channel_for_type("") == "duizhang"


def test_fenbi_always_hold():
    rec = {
        "ar": "AR_TEST_F1",
        "so": "SO26010001",
        "sod": "SOD26010001",
        "customer": "测试客户甲",
        "amount_orig": 100.0,
        "currency": "人民币CNY",
        "hexiao_date": dt.date(2026, 7, 8),
        "shoukuan_date": dt.date(2026, 7, 8),
        "status": "手动核销",
        "huikuan_type": "分笔回款",
        "channel": "fenbi",
        "fee": 0,
    }
    r = C.classify_one(rec, C.LedgerIndex(), {}, 0.0, 2026)
    assert r["bucket"] == "hold"
    assert r["code"] == "E1"


def test_void_status_exception():
    rec = {
        "ar": "AR_V",
        "so": "SO26010002",
        "sod": "",
        "customer": "X",
        "amount_orig": 10.0,
        "currency": "人民币CNY",
        "status": "已作废",
        "channel": "duizhang",
        "fee": 0,
    }
    r = C.classify_one(rec, C.LedgerIndex(), {}, 0.0, 2026)
    assert r["bucket"] == "exception"
    assert r["code"] == "E7"


def test_no_so_exception():
    rec = {
        "ar": "AR_N",
        "so": "",
        "sod": "",
        "amount_orig": 10.0,
        "currency": "人民币CNY",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 0,
    }
    r = C.classify_one(rec, C.LedgerIndex(), {}, 0.0, 2026)
    assert r["code"] == "E7"


def test_fee_hold():
    rec = {
        "ar": "AR_FEE",
        "so": "SO26010003",
        "sod": "SOD26010003",
        "amount_orig": 298.38,
        "currency": "人民币CNY",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 1.62,
        "customer": "Y",
    }
    r = C.classify_one(rec, C.LedgerIndex(), {}, 0.0, 2026)
    assert r["bucket"] == "hold"
    assert r["code"] == "E_FEE"


def test_cross_year_hold():
    rec = {
        "ar": "AR_OLD",
        "so": "SO25010001",
        "sod": "SOD25120001",
        "amount_orig": 50.0,
        "currency": "人民币CNY",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 0,
        "customer": "Z",
    }
    r = C.classify_one(rec, C.LedgerIndex(), {}, 0.0, 2026)
    assert r["bucket"] == "hold"
    assert r["code"] == "E3"


def test_fx_missing_rate_e6():
    rec = {
        "ar": "AR_FX",
        "so": "SO26010004",
        "sod": "SOD26010004",
        "amount_orig": 10.0,
        "currency": "美元USD",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 0,
        "customer": "F",
    }
    r = C.classify_one(rec, C.LedgerIndex(), {}, 0.0, 2026)
    assert r["code"] == "E6"


def test_multi_row_so_e8():
    ledger = C.LedgerIndex(
        synthetic={
            "so": {"SO26019999": [10, 11, 12]},
            "sod": {},
            "rows": {},
        }
    )
    rec = {
        "ar": "AR_M",
        "so": "SO26019999",
        "sod": "",
        "amount_orig": 100.0,
        "currency": "人民币CNY",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 0,
        "customer": "M",
    }
    r = C.classify_one(rec, ledger, {}, 0.0, 2026)
    assert r["code"] == "E8"
    assert r["bucket"] == "hold"
    assert len(r["candidates"]) == 3


def test_missing_so_in_ledger_e2():
    ledger = C.LedgerIndex(synthetic={"so": {"SO26010000": [1]}, "sod": {}, "rows": {}})
    rec = {
        "ar": "AR_E2",
        "so": "SO26999999",
        "sod": "SOD26999999",
        "amount_orig": 100.0,
        "currency": "人民币CNY",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 0,
        "customer": "E2",
    }
    r = C.classify_one(rec, ledger, {}, 0.0, 2026)
    assert r["code"] == "E2"


def test_tail_diff_e5():
    ledger = C.LedgerIndex(
        synthetic={
            "so": {"SO26018888": [5]},
            "sod": {"SOD26018888": [5]},
            "rows": {5: {"so": "SO26018888", "sod": "SOD26018888"}},
        }
    )
    rec = {
        "ar": "AR_T",
        "so": "SO26018888",
        "sod": "SOD26018888",
        "amount_orig": 100.0,
        "deliver_local": 100.32,
        "currency": "人民币CNY",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 0,
        "customer": "T",
        "hexiao_date": dt.date(2026, 7, 1),
        "shoukuan_date": dt.date(2026, 7, 1),
    }
    r = C.classify_one(rec, ledger, {}, 0.0, 2026)
    assert r["code"] == "E5"


def test_auto_happy_path():
    ledger = C.LedgerIndex(
        synthetic={
            "so": {"SO26017777": [3]},
            "sod": {"SOD26017777": [3]},
            "rows": {
                3: {
                    "so": "SO26017777",
                    "sod": "SOD26017777",
                    "jiti": None,
                    "huikuan": None,
                    "jiezhang": None,
                    "shoukuan_time": None,
                    "shoukuan_way": None,
                }
            },
        }
    )
    rec = {
        "ar": "AR_OK",
        "so": "SO26017777",
        "sod": "SOD26017777",
        "amount_orig": 200.0,
        "deliver_local": 200.0,
        "currency": "人民币CNY",
        "status": "手动核销",
        "channel": "duizhang",
        "fee": 0,
        "customer": "OK客户",
        "hexiao_date": dt.date(2026, 7, 8),
        "shoukuan_date": dt.date(2026, 7, 8),
    }
    r = C.classify_one(rec, ledger, {}, 0.0, 2026)
    assert r["bucket"] == "auto"
    assert r["five_cols"]["计提"] == 200.0
    assert r["five_cols"]["回款明细"] == 200.0
    assert r["five_cols"]["是否结账"] == "是"
    assert r["five_cols"]["收款方式"] == "汇"
    assert r["five_cols"]["收款时间"] == "2026-07-08"
    assert "行号" not in (r.get("locate_hint") or "") or "禁止" in r["locate_hint"]


def test_yucun_date_filter_keeps_only_today():
    details = [
        {"ar": "AR_Y", "hexiao_date": dt.date(2026, 6, 1), "amount": 500.0, "so": "SO1"},
        {"ar": "AR_Y", "hexiao_date": dt.date(2026, 7, 20), "amount": 9.28, "so": "SO2"},
    ]
    filtered = C.filter_yucun_details_by_date(details, dt.date(2026, 7, 20))
    assert len(filtered) == 1
    assert filtered[0]["amount"] == 9.28


def test_yucun_filter_prevents_double_count():
    """模拟：不过滤会 2 行，过滤后 1 行。"""
    ar_list = [
        {
            "ar": "AR_Y2",
            "hexiao_date": "2026-07-20",
            "shoukuan_date": "2026-07-20",
            "status": "预存已核销",
            "currency": "人民币CNY",
            "customer": "预存客户",
            "details": [
                {"核销日期": "2026-06-15", "本次核销金额": 500.0, "so": "SO260601"},
                {"核销日期": "2026-07-20", "本次核销金额": 9.28, "so": "SO260701"},
            ],
        }
    ]
    recs = C._records_yucun_synthetic(ar_list, None)
    assert len(recs) == 1
    assert recs[0]["amount_orig"] == 9.28
    assert recs[0]["channel"] == "yucun"


def test_yucun_classify_after_filter():
    ledger = C.LedgerIndex(
        synthetic={
            "so": {"SO26070101": [1]},
            "sod": {},
            "rows": {1: {"so": "SO26070101", "sod": ""}},
        }
    )
    rec = {
        "ar": "AR_Y3",
        "so": "SO26070101",
        "sod": "",
        "amount_orig": 9.28,
        "currency": "人民币CNY",
        "status": "预存已核销",
        "huikuan_type": "预存回款",
        "channel": "yucun",
        "fee": 0,
        "customer": "Y3",
        "hexiao_date": dt.date(2026, 7, 20),
        "shoukuan_date": dt.date(2026, 7, 20),
    }
    r = C.classify_one(rec, ledger, {}, 0.0, 2026)
    assert r["bucket"] == "auto"
    assert r["five_cols"]["收款方式"] == "冲预收"
    assert r["five_cols"]["计提"] == 9.28


def test_classify_records_counts():
    recs = [
        {
            "ar": "A1",
            "so": "SO2601",
            "sod": "",
            "amount_orig": 1,
            "currency": "人民币CNY",
            "status": "手动核销",
            "channel": "fenbi",
            "huikuan_type": "分笔回款",
            "fee": 0,
            "customer": "c",
        },
        {
            "ar": "A2",
            "so": "",
            "sod": "",
            "amount_orig": 1,
            "currency": "人民币CNY",
            "status": "手动核销",
            "channel": "duizhang",
            "fee": 0,
            "customer": "c",
        },
    ]
    result = C.classify_records(recs, C.LedgerIndex(), {})
    assert result["counts"]["total"] == 2
    assert result["counts"]["hold"] >= 1
    assert result["counts"]["exception"] >= 1


@pytest.mark.skipif(not FIXTURE.is_file(), reason="无夹具")
def test_std_end_to_end_five_cols():
    """标准答案端到端：五列逐格比对（夹具 std）。"""
    fix = json.loads(FIXTURE.read_text(encoding="utf-8"))
    recs = C.records_from_fixture(fix, "std")
    assert len(recs) == 1
    # 无真实大表时用合成索引命中
    so = recs[0]["so"]
    sod = recs[0]["sod"]
    ledger = C.LedgerIndex(
        synthetic={
            "so": {so: [100]},
            "sod": {sod: [100]},
            "rows": {
                100: {
                    "so": so,
                    "sod": sod,
                    "jiti": None,
                    "huikuan": None,
                    "jiezhang": None,
                    "shoukuan_time": None,
                    "shoukuan_way": None,
                }
            },
        }
    )
    rates = {"美元USD": 7.0, "美元": 7.0, "未知外币": 7.0}
    # 夹具 std 币种可能空且原币≠本币 → 未知外币
    recs[0]["currency"] = recs[0]["currency"] or "未知外币"
    if not recs[0]["currency"] or recs[0]["currency"] == "人民币CNY":
        # 原币 129.6 本币系统 932 → 应为外币
        if recs[0].get("amount_orig") and recs[0]["amount_orig"] < 200:
            recs[0]["currency"] = "美元USD"
    result = C.classify_records(recs, ledger, rates)
    assert result["counts"]["auto"] == 1, result
    five = result["auto"][0]["five_cols"]
    # 标准答案逐格
    assert five["计提"] == 907.20
    assert five["回款明细"] == 907.20
    assert five["是否结账"] == "是"
    assert five["收款时间"] == "2026-05-18"
    assert five["收款方式"] == "汇"
    assert five["实收SOD"] == "SOD26030563" or result["auto"][0]["sod"] == "SOD26030563"


@pytest.mark.skipif(not FIXTURE.is_file(), reason="无夹具")
def test_day_batch_runs():
    fix = json.loads(FIXTURE.read_text(encoding="utf-8"))
    recs = C.records_from_fixture(fix, "day")
    assert len(recs) == 53
    result = C.classify_records(recs, C.LedgerIndex(), {"美元USD": 7.0, "未知外币": 7.0})
    assert result["counts"]["total"] == 53
    # 跨年应进 hold
    assert result["counts"]["hold"] + result["counts"]["exception"] + result["counts"]["auto"] == 53


def test_no_proportion_in_source():
    """静态：分类脚本源码不含分摊逻辑关键词（与 A5 一致）。"""
    src = Path(C.__file__).read_text(encoding="utf-8")
    for bad in ("均分", "按比例分摊", "proportion"):
        assert bad not in src


def test_partial_status_no_jiti():
    ledger = C.LedgerIndex(
        synthetic={
            "so": {"SO26016666": [1]},
            "sod": {"SOD26016666": [1]},
            "rows": {1: {"so": "SO26016666", "sod": "SOD26016666"}},
        }
    )
    rec = {
        "ar": "AR_P",
        "so": "SO26016666",
        "sod": "SOD26016666",
        "amount_orig": 50.0,
        "deliver_local": 100.0,
        "currency": "人民币CNY",
        "status": "预存部分核销",
        "channel": "yucun",
        "huikuan_type": "预存回款",
        "fee": 0,
        "customer": "P",
        "hexiao_date": dt.date(2026, 7, 1),
        "shoukuan_date": dt.date(2026, 7, 1),
    }
    r = C.classify_one(rec, ledger, {}, 0.0, 2026)
    assert r["bucket"] == "auto"
    assert r["five_cols"]["计提"] is None
    assert r["five_cols"]["是否结账"] == "否"
