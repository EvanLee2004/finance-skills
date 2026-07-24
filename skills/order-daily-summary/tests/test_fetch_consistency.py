"""翻页一致性闸：服务端 count 与实抓行数不符 → 报错，绝不静默少数。"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

_SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(_SCRIPTS))

from fetch_orders import FetchError, fetch_all_rows  # noqa: E402


def _post_factory(pages: dict[int, list[dict]], count):
    def post(path, body):
        assert path == "Worksheet/GetFilterRows"
        idx = body["pageIndex"]
        rows = pages.get(idx, [])
        data = {"data": rows}
        # 真实智云：仅第 1 页(notGetTotal=False)带 count
        if not body.get("notGetTotal"):
            data["count"] = count
        return {"data": data}

    return post


def test_undercount_raises():
    """服务端说命中 11 行，实际每页只吐 10 → 必须 FetchError，不能静默返回 10。"""
    post = _post_factory({1: [{"id": i} for i in range(10)]}, count=11)
    with pytest.raises(FetchError) as ei:
        fetch_all_rows(post, "ws", "app", filter_controls=[{"x": 1}])
    assert "不一致" in str(ei.value)


def test_count_ok_short_page():
    """count=10、首页就吐满 10（<pageSize）→ 正常返回 10，不误报。"""
    post = _post_factory({1: [{"id": i} for i in range(10)]}, count=10)
    rows = fetch_all_rows(post, "ws", "app", filter_controls=[{"x": 1}])
    assert len(rows) == 10


def test_count_multipage_ok():
    """count=1050，跨两页拉满 → 返回 1050，无报错。"""
    post = _post_factory(
        {1: [{"id": i} for i in range(1000)], 2: [{"id": i} for i in range(1000, 1050)]},
        count=1050,
    )
    rows = fetch_all_rows(post, "ws", "app", filter_controls=[{"x": 1}])
    assert len(rows) == 1050


def test_no_count_falls_back_to_short_page():
    """服务端不给 count → 退回“短页即止”旧行为，不报错。"""

    def post(path, body):
        idx = body["pageIndex"]
        pages = {1: [{"id": i} for i in range(10)]}
        return {"data": {"data": pages.get(idx, [])}}  # 无 count

    rows = fetch_all_rows(post, "ws", "app", filter_controls=[{"x": 1}])
    assert len(rows) == 10
