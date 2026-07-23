# -*- coding: utf-8 -*-
"""common 工具与列匹配。"""
import datetime as dt
import pytest

import common


def test_to_number_basic():
    assert common.to_number("1,234.50") == 1234.50
    assert common.to_number(None) is None
    assert common.to_number("") is None


def test_norm_date():
    assert common.norm_date("2026-05-18") == dt.date(2026, 5, 18)
    assert common.norm_date(dt.datetime(2026, 1, 2, 3, 4)) == dt.date(2026, 1, 2)


def test_fuzzy_find_col():
    headers = ["日期", "回款记录ID", "回款类型", "到账金额/原币"]
    assert common.fuzzy_find_col(headers, ["回款类型"]) == 2
    assert common.fuzzy_find_col(headers, ["不存在的列"]) is None


def test_resolve_columns_ok():
    headers = ["回款记录ID", "核销日期", "回款类型", "开票客户"]
    cols = common.resolve_columns(headers, "回款记录", ["AR", "回款类型"])
    assert cols["AR"] == 0
    assert cols["回款类型"] == 2


def test_resolve_columns_missing_lists_headers():
    headers = ["AAA", "BBB"]
    with pytest.raises(ValueError) as ei:
        common.resolve_columns(headers, "回款记录", ["AR", "回款类型"])
    msg = str(ei.value)
    assert "实际表头" in msg
    assert "AAA" in msg


def test_is_cny():
    assert common.is_cny("人民币CNY")
    assert common.is_cny("CNY")
    assert not common.is_cny("美元USD")


def test_year_from_so():
    assert common.year_from_so("SO26030412") == 2026
    assert common.year_from_so("SOD25120001") == 2025


def test_receipt_time_same_month():
    a = dt.date(2026, 5, 18)
    h = dt.date(2026, 5, 27)
    assert common.receipt_time(a, h) == a


def test_receipt_time_cross_month():
    a = dt.date(2026, 4, 30)
    h = dt.date(2026, 5, 2)
    assert common.receipt_time(a, h) == h


def test_compute_local_cny():
    v, err = common.compute_local_amount(100.0, "人民币CNY", {})
    assert err is None and v == 100.0


def test_compute_local_fx_need_rate():
    v, err = common.compute_local_amount(129.60, "美元USD", {})
    assert v is None and err == "E6"


def test_compute_local_fx_with_rate():
    v, err = common.compute_local_amount(129.60, "美元USD", {"美元USD": 7.0})
    assert err is None and v == 907.20


def test_pay_way():
    """2026-07-23：只按「到账月 vs 核销月」判；回款类型不再参与（旧规则甲已证伪）。"""
    import datetime as _dt

    assert common.pay_way("手动核销") == "汇"
    assert common.pay_way("预存已核销") == "汇"  # 旧版这里错判冲预收，15 笔全错
    d = _dt.date(2026, 7, 22)
    assert common.pay_way("核销成功", _dt.date(2026, 7, 17), d) == "汇"
    assert common.pay_way("手动核销", _dt.date(2026, 6, 26), d) == "冲预收"


def test_mask_customer():
    m = common.mask_customer("北京多语信息技术有限公司")
    assert m[0] == "北"
    assert "*" in m
    assert "多语" not in m or m != "北京多语信息技术有限公司"


def test_tail_threshold_default():
    assert common.tail_threshold() == 0.0


def test_parse_rate_args():
    d = common.parse_rate_args(["美元USD=7.0", "EUR=7.5"])
    assert d["美元USD"] == 7.0
