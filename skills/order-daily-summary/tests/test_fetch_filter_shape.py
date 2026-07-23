"""离线测 filter body / date_window 形状，不连网。"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

_SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(_SCRIPTS))

from fetch_orders import (  # noqa: E402
    build_date_range_filter,
    client_filter_by_date,
    fetch_all_rows,
    parse_cell,
    rows_to_records,
)


def test_build_date_range_filter_shape():
    controls = [
        {"controlId": "cid_date", "controlName": "下单日期", "type": 15},
        {"controlId": "cid_amt", "controlName": "下单预估额/本币", "type": 6},
    ]
    fc = build_date_range_filter(controls, "下单日期", date(2026, 7, 22), date(2026, 7, 22))
    assert len(fc) == 1
    f = fc[0]
    assert f["controlId"] == "cid_date"
    assert f["filterType"] == 11
    assert f["dateRange"] == 18
    assert f["minValue"] == "2026-07-22"
    assert f["maxValue"] == "2026-07-22"
    assert f.get("spliceType") == 1


def test_build_date_range_filter_missing_col():
    assert build_date_range_filter([], "下单日期", "2026-01-01", "2026-01-02") == []


def test_client_filter_inclusive():
    records = [
        {"下单日期": "2026-07-21", "销售": "A"},
        {"下单日期": "2026-07-22", "销售": "B"},
        {"下单日期": "2026-07-23", "销售": "C"},
        {"下单日期": "bad", "销售": "D"},
    ]
    out = client_filter_by_date(records, "2026-07-22", "2026-07-22")
    assert len(out) == 1
    assert out[0]["销售"] == "B"


def test_fetch_all_rows_pagination():
    pages = {
        1: [{"id": i} for i in range(1000)],
        2: [{"id": i} for i in range(1000, 1050)],
    }

    def post(path, body):
        assert path == "Worksheet/GetFilterRows"
        assert body["pageSize"] == 1000
        idx = body["pageIndex"]
        return {"data": {"data": pages.get(idx, []), "count": 1050}}

    rows = fetch_all_rows(post, "ws", "app", filter_controls=[{"x": 1}])
    assert len(rows) == 1050


def test_rows_to_records_and_parse_cell():
    controls = [
        {"controlId": "c1", "controlName": "销售", "type": 2},
        {
            "controlId": "c2",
            "controlName": "选项",
            "type": 9,
            "options": [{"key": "k1", "value": "中文A"}],
        },
    ]
    rows = [{"c1": "张三", "c2": '["k1"]'}]
    recs = rows_to_records(rows, controls)
    assert recs[0]["销售"] == "张三"
    assert recs[0]["选项"] == "中文A"
    assert parse_cell(None, {}) == ""
