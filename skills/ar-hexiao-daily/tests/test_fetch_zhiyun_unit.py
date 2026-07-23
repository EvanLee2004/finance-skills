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


def test_plain_option_and_relation():
    opts = {"k1": "整笔回款"}
    assert F._plain('["k1"]', opts) == "整笔回款"
    assert F._plain('[{"name":"某某客户"}]') == "某某客户"
    assert F._plain(None) == ""


def test_no_credentials_in_source():
    src = Path(__file__).resolve().parents[1] / "scripts" / "fetch_zhiyun.py"
    text = src.read_text(encoding="utf-8")
    assert "sharon" not in text.lower()
    assert "besteasy.com" not in text
    assert "password =" not in text.lower() or "getpass" in text
    # 明确禁止硬编码密码赋值
    assert "ZHIYUN_PASS='" not in text
    assert 'ZHIYUN_PASS="' not in text
