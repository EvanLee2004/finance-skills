# -*- coding: utf-8 -*-
"""fetch_zhiyun 纯函数单测（不连网、不碰账密）。"""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import fetch_zhiyun as F


class SettlementClient:
    id_by_name = staticmethod(F.ZhiyunClient.id_by_name)

    def __init__(self, orders):
        self.orders = orders
        self.calls = []

    def datasource_of(self, worksheet, relation):
        assert (worksheet, relation) == (F.WS_HUIKUAN, F.REL_JIESUAN)
        return "settlements"

    def relation_rows(self, worksheet, row, control):
        self.calls.append((worksheet, row, control))
        assert worksheet == "settlements" and control == "orders"
        value = self.orders[row]
        if isinstance(value, Exception):
            raise value
        return value, [
            {"controlId": "so", "controlName": "SO"},
            {"controlId": "amount", "controlName": "交付额/本币"},
        ]


def test_settlement_nested_orders_keep_order_amounts_and_deduplicate():
    client = SettlementClient({
        "s1": [{"so": "SO26000001", "amount": 10},
               {"so": "SO26000002", "amount": 20}],
        "s2": [{"so": "SO26000001", "amount": 10}],
    })
    controls = [
        {"controlId": "orders", "controlName": "订单"},
        {"controlId": "amount", "controlName": "交付额/本币"},
    ]
    result = F.settlement_related_orders(
        client, [{"rowid": "s1", "amount": 999}, {"rowid": "s2"}], controls,
    )
    assert [v["so"] for v in result] == ["SO26000001", "SO26000002"]
    assert [v["deliver_local"] for v in result] == ["10", "20"]
    assert all(v["written_off"] == "" for v in result)
    assert all(v["source"] == "结算" for v in result)


def test_settlement_direct_order_does_not_query_nested_relation():
    client = SettlementClient({})
    rows = [{"so": "SO26000001", "amount": 10}]
    controls = [{"controlId": "so", "controlName": "SO"},
                {"controlId": "amount", "controlName": "交付额/本币"}]
    assert F.settlement_related_orders(client, rows, controls) == (
        F.extract_related_orders(rows, controls, F.REL_JIESUAN)
    )
    assert client.calls == []


@pytest.mark.parametrize("orders", [
    {"s1": [{"so": "SO26000001", "amount": 10}],
     "s2": [{"so": "SO26000001", "amount": 20}]},
    {"s1": [{"so": "SO26000001"}], "s2": [{"so": ""}]},
    {"s1": [{"so": "SO26000001"}], "s2": F.FetchError("read failed")},
])
def test_settlement_nested_conflict_or_incomplete_read_blocks(orders):
    with pytest.raises(F.FetchError):
        F.settlement_related_orders(
            SettlementClient(orders), [{"rowid": "s1"}, {"rowid": "s2"}],
            [{"controlId": "orders", "controlName": "订单"}],
        )


def test_settlement_without_order_relation_remains_empty():
    client = SettlementClient({})
    assert F.settlement_related_orders(client, [{"rowid": "s1"}], []) == []
    assert client.calls == []


@pytest.mark.parametrize("source", ["下单", "结算", "结算订单"])
def test_fetch_day_routes_settlement_fallback_and_exports_each_so(source, monkeypatch, tmp_path):
    order_controls = [
        {"controlId": "so", "controlName": "SO"},
        {"controlId": "amount", "controlName": "交付额/本币"},
        {"controlId": "date", "controlName": "项目交付日期"},
    ]
    orders = [{"so": "SO26000001", "amount": 10, "date": "2026-09-01"},
              {"so": "SO26000002", "amount": 20, "date": "2026-09-01"}]

    class Client(F.ZhiyunClient):
        def __init__(self):
            self.calls = []

        def controls(self, worksheet):
            if worksheet == F.WS_HUIKUAN:
                return [
                    {"controlId": name, "controlName": name, "dataSource": name}
                    for name in (F.REL_XIADAN, F.REL_JIESUAN, F.REL_HEXIAO_MINGXI)
                ]
            return order_controls

        def filter_rows_by_date(self, *args):
            return [{"rowid": "payment", F.F_HK["ar"]: "AR_TEST_001",
                     F.F_HK["hexiao_date"]: "2026-09-07"}], 1

        def relation_rows(self, worksheet, row, control):
            self.calls.append((worksheet, row, control))
            if control == F.REL_XIADAN:
                return (orders if source == "下单" else []), order_controls
            if control == F.REL_JIESUAN:
                assert source != "下单"
                if source == "结算":
                    return orders, order_controls
                return [{"rowid": "settlement"}], [
                    {"controlId": "orders", "controlName": "订单"},
                ]
            if control == "orders":
                assert (worksheet, row, source) == ("结算", "settlement", "结算订单")
                return orders, order_controls
            assert control == F.REL_HEXIAO_MINGXI
            return [], []

        def search_rows(self, worksheet, so):
            assert worksheet == F.REL_HEXIAO_MINGXI
            assert so in {"SO26000001", "SO26000002"}
            return []

    captured = {}

    def publish(_out, _tag, datasets, summary):
        captured.update({name: rows for name, _headers, rows in datasets})
        return summary

    monkeypatch.setattr(F, "publish_day_exports", publish)
    client = Client()
    summary = F.fetch_day(client, "2026-09-07", tmp_path)
    assert summary["无下单行的AR"] == []
    assert summary["从结算找回单号的AR数"] == (0 if source == "下单" else 1)
    assert [row[1] for row in captured["订单交付_20260907.xlsx"]] == [
        "SO26000001", "SO26000002",
    ]
    assert [row[5] for row in captured["订单交付_20260907.xlsx"]] == ["10", "20"]
    assert captured["核销明细_20260907.xlsx"] == []
    assert any(control == "orders" for _, _, control in client.calls) == (source == "结算订单")


def test_resolve_date_yesterday():
    d = F.resolve_date("yesterday")
    assert len(d) == 10 and d[4] == "-" and d[7] == "-"


def test_resolve_date_fixed():
    assert F.resolve_date("2026-07-21") == "2026-07-21"


def test_existing_exports_require_current_schema_version(tmp_path):
    day = "2026-07-27"
    stamp = "20260727"
    for role in ("回款记录", "订单交付", "核销明细", "订单明细"):
        (tmp_path / f"{role}_{stamp}.xlsx").write_bytes(b"fixture")

    # 旧四件套没有版本摘要时，默认禁止跳过重新取数。
    assert F.already_fetched(tmp_path, day) == []
    assert len(F.already_fetched(tmp_path, day, accept_unversioned=True)) == 4

    (tmp_path / f"取数摘要_{stamp}.json").write_text(
        '{"export_schema_version":"old"}',
        encoding="utf-8",
    )
    assert F.already_fetched(tmp_path, day) == []

    files = [f"{role}_{stamp}.xlsx" for role in ("回款记录", "订单交付", "核销明细", "订单明细")]
    (tmp_path / f"取数摘要_{stamp}.json").write_text(
        json.dumps({
            "export_schema_version": F.EXPORT_SCHEMA_VERSION,
            "file_sha256": {name: F._sha256(tmp_path / name) for name in files},
        }, ensure_ascii=False),
        encoding="utf-8",
    )
    assert len(F.already_fetched(tmp_path, day)) == 4


def test_plain_option_and_relation():
    opts = {"k1": "整笔回款"}
    assert F._plain('["k1"]', opts) == "整笔回款"
    assert F._plain('[{"name":"某某客户"}]') == "某某客户"
    assert F._plain(None) == ""


def test_settlement_relation_recovers_order_when_xiadan_is_empty():
    controls = [
        {"controlId": "order", "controlName": "结算订单"},
        {"controlId": "written", "controlName": "订单已核销金额"},
        {"controlId": "amount", "controlName": "交付额/原币"},
        {"controlId": "currency", "controlName": "结算币种"},
    ]
    rows = [{
        "order": '[{"name":"SO26000001"}]',
        "written": "120.47",
        "amount": "120.50",
        "currency": "人民币CNY",
    }]

    got = F.extract_related_orders(rows, controls, F.REL_JIESUAN)

    assert got == [{
        "so": "SO26000001",
        "written_off": "120.47",
        "written_off_local": "",
        "deliver": "120.50",
        "deliver_local": "",
        "rate": "",
        "currency": "人民币CNY",
        "name": "",
        "delivery_date": "",
        "delivery_date_status": "",
        "source": "结算",
    }]


def test_related_order_reads_delivery_amount_local():
    controls = [
        {"controlId": "order", "controlName": "SO"},
        {"controlId": "original", "controlName": "交付额/原币"},
        {"controlId": "local", "controlName": "交付额/本币"},
    ]
    rows = [{
        "order": "SO26060803",
        "original": "8844.00",
        "local": "63605.16",
    }]

    got = F.extract_related_orders(rows, controls, F.REL_XIADAN)

    assert got[0]["deliver"] == "8844.00"
    assert got[0]["deliver_local"] == "63605.16"


def test_related_order_reads_project_delivery_date_from_order_detail():
    controls = [
        {"controlId": "order", "controlName": "SO"},
        {"controlId": "date", "controlName": "项目交付日期"},
    ]
    rows = [{"order": "SO24100160", "date": "2025-08-13"}]

    got = F.extract_related_orders(rows, controls, F.REL_XIADAN)

    assert got[0]["delivery_date"] == "2025-08-13"
    assert got[0]["delivery_date_status"] == "关联下单明确值"


def test_lookup_order_delivery_date_requires_unique_explicit_date():
    controls = [
        {"controlId": "order", "controlName": "SO"},
        {"controlId": "date", "controlName": "项目交付日期"},
    ]

    class FakeClient:
        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        @staticmethod
        def search_rows(_worksheet_id, _so):
            return [
                {"order": "SO24100160", "date": "2025-08-13"},
                {"order": "SO24100160", "date": "2025-08-13"},
                {"order": "SO99999999", "date": "2024-01-01"},
            ]

    value, status = F.lookup_order_delivery_date(
        FakeClient(), "orders", controls, "SO24100160"
    )
    assert value == "2025-08-13"
    assert status == "订单详情明确值"


def test_no_credentials_in_source():
    src = Path(__file__).resolve().parents[1] / "scripts" / "fetch_zhiyun.py"
    text = src.read_text(encoding="utf-8")
    # 禁止真实账号/密码痕迹（允许文档里出现变量名 ZHIYUN_PASS）
    assert "sharon" not in text.lower()
    assert "sharon1234" not in text
    assert "getpass" not in text  # 核销任务不在取数中途交互询问
    assert "input(" not in text
    assert "自动任务不会在取数过程中弹出账号密码询问" in text
    # 禁止把真实密码字面量赋给环境示例
    assert "PASS='****'" not in text


def test_historical_writeoffs_for_sos_gets_cross_parent_history_only():
    names = [
        "核销记录NUM", "回款记录NUM", "订单NUM", "本次核销金额", "本次核销金额本币",
        "核销日期", "币种", "汇率", "订单名称", "是否已撤销",
    ]
    controls = [
        {"controlId": f"c{i}", "controlName": name}
        for i, name in enumerate(names)
    ]

    def row(**values):
        return {
            f"c{i}": values.get(name, "")
            for i, name in enumerate(names)
        }

    class FakeClient:
        def controls(self, _worksheet_id):
            return controls

        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        def search_rows(self, _worksheet_id, so):
            assert so == "SO26000001"
            return [
                row(
                    核销记录NUM="HX_OLD_001",
                    回款记录NUM='[{"name":"AR_OLD_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=30,
                    本次核销金额本币=30,
                    核销日期="2026-06-11",
                ),
                row(  # 目标日当前行由父回款关联子表负责，不在这里重复补。
                    核销记录NUM="HX_NOW_001",
                    回款记录NUM='[{"name":"AR_NOW_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=10,
                    本次核销金额本币=10,
                    核销日期="2026-07-24",
                ),
                row(  # 已撤销历史行不计。
                    核销记录NUM="HX_OLD_002",
                    回款记录NUM='[{"name":"AR_OLD_002"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=5,
                    本次核销金额本币=5,
                    核销日期="2026-05-01",
                    是否已撤销="是",
                ),
                row(  # 全文搜索误命中的其它 SO 必须精确排除。
                    核销记录NUM="HX_OTHER",
                    回款记录NUM='[{"name":"AR_OTHER"}]',
                    订单NUM='[{"name":"SO260000010"}]',
                    本次核销金额=99,
                    本次核销金额本币=99,
                    核销日期="2026-04-01",
                ),
            ]

    got = F.historical_writeoffs_for_sos(
        FakeClient(), "WS_MX", ["SO26000001"], "2026-07-24"
    )
    assert len(got) == 2  # 撤销记录也保留给分类器审计，但不参与金额。
    assert got[0][0] == "HX_OLD_001"
    assert got[0][2] == "AR_OLD_001"
    assert got[0][3] == "2026-06-11"
    assert got[0][8] == "SO26000001"
    assert got[1][10] == "是"


def test_historical_writeoffs_only_dedup_same_record_id_and_never_business_fields():
    names = [
        "核销记录NUM", "回款记录NUM", "订单NUM", "本次核销金额", "本次核销金额本币",
        "核销日期", "币种", "汇率", "订单名称", "是否已撤销",
    ]
    controls = [
        {"controlId": f"c{i}", "controlName": name}
        for i, name in enumerate(names)
    ]

    def row(record_id):
        values = {
            "核销记录NUM": record_id,
            "回款记录NUM": '[{"name":"AR26070001"}]',
            "订单NUM": '[{"name":"SO26000001"}]',
            "本次核销金额": 100,
            "本次核销金额本币": 100,
            "核销日期": "2026-07-01",
            "币种": "CNY",
        }
        return {
            **{f"c{i}": values.get(name, "") for i, name in enumerate(names)},
            "rowid": f"ROW-{record_id or 'NONE'}",
        }

    class FakeClient:
        def controls(self, _worksheet_id):
            return controls

        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        def search_rows(self, _worksheet_id, _so):
            return [row("HX1"), row("HX1"), row("HX2"), row(""), row("")]

    got = F.historical_writeoffs_for_sos(
        FakeClient(), "WS_MX", ["SO26000001"], "2026-07-31"
    )
    assert [item[0] for item in got] == ["HX1", "HX2", "", ""]
