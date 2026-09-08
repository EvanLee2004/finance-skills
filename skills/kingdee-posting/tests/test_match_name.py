#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from pathlib import Path
import importlib.util
import sys

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


m = _load("kingdee_match_name", SCRIPTS / "match_name.py")


def test_peel_limited_company_unique():
    recs = [("2002", "中国广播电影电视节目交易中心")]
    got = m.match_records("中国广播电影电视节目交易中心有限公司", recs)
    assert got.status == "ok"
    assert got.hit[0] == "2002"


def test_doubled_archive_name():
    recs = [("9001", "乙科技有限公司乙科技有限公司")]
    got = m.match_records("乙科技有限公司", recs)
    assert got.status == "ok"
    assert got.hit[0] == "9001"


def test_many_candidates_not_auto():
    recs = [("1", "甲科技有限公司"), ("2", "甲科技股份有限公司")]
    got = m.match_records("甲科技", recs)
    assert got.status == "many"
    assert got.hit is None


def test_person_and_police_flags():
    assert m.is_person_heading("刘芳") is True
    assert m.is_person_heading("咪咕数字传媒有限公司") is False
    assert m.maps_to_police_ministry("平度市公安局", "于占国") is True
    assert m.maps_to_police_ministry("北京市公安局海淀分局", "陈霞") is False
    assert m.is_haidian_police("北京市公安局海淀分局") is True
    assert m.is_haidian_police("平度市公安局") is False
    assert m.hang_employee("童睿智", {"童睿智": "郑瑞"}) == "郑瑞"
    assert m.hang_employee("路人甲", {"童睿智": "郑瑞"}) == "路人甲"
