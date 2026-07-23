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
    # 禁止真实账号/密码痕迹（允许文档里出现变量名 ZHIYUN_PASS）
    assert "sharon" not in text.lower()
    assert "sharon1234" not in text
    assert "getpass" in text  # 必须支持交互输入
    # 禁止把真实密码字面量赋给环境示例
    assert "PASS='****'" not in text
