#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合成数据回归：表结构模板完整、脚本可跑、整挂规则、核对逻辑。不含真实金额/人名。"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
SCRIPTS = SKILL / "scripts"
CONFIG = SKILL / "config"


def test_structure_template_complete():
    st = json.loads((CONFIG / "表结构模板.json").read_text(encoding="utf-8"))
    assert len(st["entities"]) >= 3
    assert len(st["departments"]) >= 20
    assert len(st["account_rows"]) >= 100
    codes = [a["code"] for a in st["account_rows"] if a["code"]]
    assert "5101" in codes and "5401" in codes and "5502" in codes
    assert not any(c.endswith(".0") for c in codes)
    depts = [d["name"] for d in st["departments"]]
    assert "本地化" in depts and "财务中心" in depts and "运营保障中心" in depts


def test_business_rules_file():
    text = (CONFIG / "业务规则.md").read_text(encoding="utf-8")
    assert "按人拆" in text
    assert "540101" in text or "540103" in text or "项目总监" in text
    assert "所有主体合计" in text or "左" in text
    assert "5401" in text


def _make_synthetic_balance(path: Path, entity_tag: str):
    """最小 xlsx 余额表：5101 贷、5401/5402/5504 借。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "发生额及余额表"
    ws["A2"] = entity_tag
    ws["A3"] = "科目编码"
    ws["B3"] = "科目名称"
    ws["C3"] = "期初余额"
    ws["E3"] = "本期发生"
    ws["G3"] = "期末余额"
    ws["C4"] = "借方"
    ws["D4"] = "贷方"
    ws["E4"] = "借方"
    ws["F4"] = "贷方"
    # 5101 收入 贷 1000
    ws["A5"] = "5101"
    ws["B5"] = "主营业务收入"
    ws["F5"] = 1000
    # 540103 翻译语言服务 借 200 → 整挂项目总监及助理
    ws["A6"] = "540103"
    ws["B6"] = "翻译语言服务"
    ws["E6"] = 200
    # 5402 税金 借 50
    ws["A7"] = "5402"
    ws["B7"] = "税金及附加"
    ws["E7"] = 50
    # 5504 财务费用 借 10
    ws["A8"] = "5504"
    ws["B8"] = "财务费用"
    ws["E8"] = 10
    wb.save(path)


def test_allocate_synthetic_direct_hang():
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        inp = td / "input"
        inp.mkdir()
        # name must trigger entity discovery
        _make_synthetic_balance(inp / "文化_余额.xlsx", "文化")
        out = td / "out.xlsx"
        r = subprocess.run(
            [
                sys.executable,
                str(SCRIPTS / "allocate.py"),
                "--input-dir",
                str(inp),
                "--out",
                str(out),
            ],
            capture_output=True,
            text=True,
        )
        assert r.returncode == 0, r.stderr + r.stdout
        assert out.is_file()
        wb = openpyxl.load_workbook(out, data_only=True)
        assert "部门科目余额表" in wb.sheetnames
        assert "利润表" in wb.sheetnames
        ws = wb["部门科目余额表"]
        # find 5402 row and 财务中心 col
        headers = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1)}
        assert "财务中心" in headers
        assert "项目总监及助理" in headers
        fee_col = headers.get("费用合计")
        check_col = headers.get("核对")
        fin_col = headers["财务中心"]
        pm_col = headers["项目总监及助理"]
        found_5402 = found_5401 = False
        for r in range(3, ws.max_row + 1):
            code = str(ws.cell(r, 1).value)
            if code == "5402":
                found_5402 = True
                assert float(ws.cell(r, fin_col).value or 0) == 50
                assert float(ws.cell(r, fee_col).value or 0) == 50
                assert abs(float(ws.cell(r, check_col).value or 0)) < 0.02
            if code == "540103":
                found_5401 = True
                assert float(ws.cell(r, pm_col).value or 0) == 200
                assert abs(float(ws.cell(r, check_col).value or 0)) < 0.02
        assert found_5402 and found_5401


def test_inspect_runs():
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        inp = td / "01_主体余额表"
        inp.mkdir(parents=True)
        _make_synthetic_balance(inp / "上海.xlsx", "上海")
        r = subprocess.run(
            [sys.executable, str(SCRIPTS / "inspect_inputs.py"), "--input-dir", str(td)],
            capture_output=True,
            text=True,
        )
        # may return 2 if missing required — still should not crash
        assert r.returncode in (0, 2), r.stderr + r.stdout
        assert "盘点目录" in r.stdout or "盘点" in r.stdout


if __name__ == "__main__":
    test_structure_template_complete()
    test_business_rules_file()
    test_allocate_synthetic_direct_hang()
    test_inspect_runs()
    print("OK all tests passed")
