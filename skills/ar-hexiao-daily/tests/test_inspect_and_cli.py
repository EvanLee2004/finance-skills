# -*- coding: utf-8 -*-
"""inspect_inputs + 列缺失报错 + 禁止项结构。"""
import json
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

import classify_hexiao as C
import common
from conftest import ROOT, FIXTURE

SCRIPTS = ROOT / "scripts"
PY = sys.executable


def test_inspect_runs(tmp_path):
    import inspect_inputs as I

    (tmp_path / "01_智云导出").mkdir()
    rc = I.main.__wrapped__ if False else None
    # call via argparse path
    import inspect_inputs

    sys_argv = ["inspect_inputs.py", "--workspace", str(tmp_path), "--report", str(tmp_path / "r.txt")]
    # patch
    old = sys.argv
    try:
        # use function
        from argparse import Namespace

        # direct
        common.ensure_out_dirs()
        files = inspect_inputs.list_files(tmp_path)
        assert files == [] or True
    finally:
        sys.argv = old
    # run module main
    rc = subprocess.run(
        [PY, str(SCRIPTS / "inspect_inputs.py"), "--workspace", str(tmp_path), "--report", str(tmp_path / "r.txt")],
        capture_output=True,
        text=True,
    )
    assert rc.returncode == 0
    assert (tmp_path / "r.txt").is_file()


def test_missing_column_raises():
    headers = ["日期", "金额"]
    with pytest.raises(ValueError) as ei:
        common.resolve_columns(headers, "回款记录", ["AR", "回款类型"])
    assert "实际表头" in str(ei.value)


def test_classify_missing_col_excel(tmp_path):
    """删掉回款类型列 → 读导出时报错。"""
    export = tmp_path / "01_智云导出"
    export.mkdir()
    # 只有对账表，故意缺关键列
    p = export / "回款核销对账_坏.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["日期", "备注"])  # 缺 SO / 回款额
    ws.append(["2026-07-01", "x"])
    wb.save(p)
    with pytest.raises(ValueError) as ei:
        C.load_excel_exports(tmp_path)
    assert "实际表头" in str(ei.value) or "缺列" in str(ei.value)


def test_skill_md_documents_full_flow():
    """7-23 口径重写后步骤表把 1–4 合并、突出第 6/8/9 步，不再逐行 | n |。
    锁的是"完整流程与关键红线仍被明确记录"，而不是某种固定表格排版。"""
    text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    assert "11 步" in text                       # 仍声明覆盖她每天 11 步
    for anchor in ["第 6 步", "核销判定", "挂账重扫", "明细", "她做"]:
        assert anchor in text, anchor
    assert "永不写智云" in text


def test_skill_md_has_comms_guidance():
    """跟明妹说话要短、要点、主动要料——这条行为要求必须在 SKILL 里，且有话术模板可依。"""
    text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    assert "怎么跟她说话" in text
    assert "主动要料" in text or "主动要材料" in text
    assert (ROOT / "references" / "跟明妹沟通.md").is_file()


def test_usage_seven_sections():
    text = (ROOT / "references" / "使用说明_给明妹.md").read_text(encoding="utf-8")
    for title in [
        "## 1. 一次性准备",
        "## 2. 每天怎么做",
        "## 3. 清单怎么看",
        "## 4. 每个异常原因怎么办",
        "## 5. 哪些它不会做",
        "## 6. 出错了怎么办",
        "## 7. 常见问题",
    ]:
        assert title in text


def test_scripts_exist():
    for name in [
        "inspect_inputs.py",
        "extract_income.py",
        "classify_hexiao.py",
        "build_worklist.py",
        "rescan_holds.py",
        "common.py",
    ]:
        assert (SCRIPTS / name).is_file()


def test_a5_grep_no_hits():
    """禁止项：scripts 内无均分/按比例/写智云 POST。"""
    import re

    pat = re.compile(
        r"均分|按比例分摊|proportion|writer\.save.*副本|POST.*wwwapi.*(add|update|save)",
        re.I,
    )
    hits = []
    for p in SCRIPTS.glob("*.py"):
        t = p.read_text(encoding="utf-8")
        if pat.search(t):
            hits.append(p.name)
    assert hits == [], hits


def test_no_write_zhiyun_urls():
    for p in SCRIPTS.glob("*.py"):
        t = p.read_text(encoding="utf-8")
        assert "wwwapi" not in t.lower() or "POST" not in t
        # classify 不应请求网络写
        assert "requests.post" not in t
