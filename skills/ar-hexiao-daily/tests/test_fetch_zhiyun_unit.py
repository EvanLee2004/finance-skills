# -*- coding: utf-8 -*-
"""fetch_zhiyun 纯函数单测（不连网、不碰账密）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import fetch_zhiyun as F


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

    (tmp_path / f"取数摘要_{stamp}.json").write_text(
        '{"export_schema_version":"' + F.EXPORT_SCHEMA_VERSION + '"}',
        encoding="utf-8",
    )
    assert len(F.already_fetched(tmp_path, day)) == 4


def test_plain_option_and_relation():
    opts = {"k1": "整笔回款"}
    assert F._plain('["k1"]', opts) == "整笔回款"
    assert F._plain('[{"name":"某某客户"}]') == "某某客户"
    assert F._plain(None) == ""


def test_no_credentials_in_source():
    src = Path(__file__).resolve().parents[1] / "scripts" / "fetch_zhiyun.py"
    text = src.read_text(encoding="utf-8")
    # 禁止真实账号/密码痕迹（允许文档里出现变量名 ZHIYUN_PASS）
    assert "sharon" not in text.lower()
    assert "sharon1234" not in text
    assert "getpass" in text  # 必须支持交互输入
    # 禁止把真实密码字面量赋给环境示例
    assert "PASS='****'" not in text


def test_historical_writeoffs_for_sos_gets_cross_parent_history_only():
    names = [
        "回款记录NUM", "订单NUM", "本次核销金额", "本次核销金额本币",
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
                    回款记录NUM='[{"name":"AR_OLD_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=30,
                    本次核销金额本币=30,
                    核销日期="2026-06-11",
                ),
                row(  # 目标日当前行由父回款关联子表负责，不在这里重复补。
                    回款记录NUM='[{"name":"AR_NOW_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=10,
                    本次核销金额本币=10,
                    核销日期="2026-07-24",
                ),
                row(  # 已撤销历史行不计。
                    回款记录NUM='[{"name":"AR_OLD_002"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=5,
                    本次核销金额本币=5,
                    核销日期="2026-05-01",
                    是否已撤销="是",
                ),
                row(  # 全文搜索误命中的其它 SO 必须精确排除。
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
    assert len(got) == 1
    assert got[0][0] == "AR_OLD_001"
    assert got[0][1] == "2026-06-11"
    assert got[0][6] == "SO26000001"
