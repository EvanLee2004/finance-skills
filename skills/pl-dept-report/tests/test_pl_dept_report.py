#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
SCRIPTS = SKILL / "scripts"
CONFIG = SKILL / "config"
CONVERT = SCRIPTS / "convert.py"
POSTING_API = SKILL.parent / "kingdee-posting" / "scripts" / "kingdee_api.py"

for _name in ("inspect_inputs", "parse_export", "layout", "common", "formula_eval"):
    sys.modules.pop(_name, None)
sys.path.insert(0, str(SCRIPTS))
from formula_eval import eval_workbook  # noqa: E402
from layout import account_row_map, dept_col_letter, dept_columns, direct_children, load_layout  # noqa: E402


def _run(args: list[str], env: dict | None = None) -> subprocess.CompletedProcess:
    run_env = os.environ.copy()
    run_env["PL_DEPT_SKIP_API"] = "1"
    run_env["PL_DEPT_USE_CACHE"] = "0"
    run_env["KINGDEE_LOCAL_JSON"] = str(HERE / "_missing_kingdee.json")
    run_env["KINGDEE_PL_LOCAL_JSON"] = str(HERE / "_missing_pl.json")
    if env:
        run_env.update(env)
    return subprocess.run(
        [sys.executable, str(CONVERT), *args],
        capture_output=True,
        text=True,
        env=run_env,
    )


def _write_account(path: Path, company: str, rows: list[tuple[str, str, float | None, float | None]]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "科目余额表"
    ws["A1"] = "科目余额表"
    ws["A2"] = company
    ws["A3"] = "科目编码"
    ws["B3"] = "科目名称"
    ws["C3"] = "本期发生借方"
    ws["D3"] = "本期发生贷方"
    for i, (code, name, debit, credit) in enumerate(rows, start=4):
        ws.cell(i, 1).value = code
        ws.cell(i, 2).value = name
        ws.cell(i, 3).value = debit
        ws.cell(i, 4).value = credit
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _write_assist(path: Path, company: str, rows: list[tuple[str, str, str, float | None, float | None]]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "核算项目余额表"
    ws["A1"] = "核算项目余额表"
    ws["B1"] = "核算项目类别：部门"
    ws["A2"] = company
    ws["A3"] = "科目编码"
    ws["B3"] = "科目名称"
    ws["C3"] = "核算项目名称"
    ws["D3"] = "本期发生借方"
    ws["E3"] = "本期发生贷方"
    for i, (code, name, dept, debit, credit) in enumerate(rows, start=4):
        ws.cell(i, 1).value = code
        ws.cell(i, 2).value = name
        ws.cell(i, 3).value = dept
        ws.cell(i, 4).value = debit
        ws.cell(i, 5).value = credit
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _write_profit(path: Path, company: str, items: list[tuple[str, float]]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "利润表"
    ws["A1"] = "利润表"
    ws["A2"] = company
    ws["A3"] = "项目"
    ws["B3"] = "本月金额"
    for i, (name, amt) in enumerate(items, start=4):
        ws.cell(i, 1).value = name
        ws.cell(i, 2).value = amt
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def test_same_dept_across_books_is_summed(tmp_path: Path):
    _write_account(
        tmp_path / "wh.xlsx",
        "北京甲骨易文化传媒有限公司",
        [("540103", "翻译语言服务", 10.0, None)],
    )
    _write_assist(
        tmp_path / "wh_d.xlsx",
        "北京甲骨易文化传媒有限公司",
        [("540103", "翻译语言服务", "大客户", 10.0, None)],
    )
    _write_account(
        tmp_path / "sh.xlsx",
        "甲骨易智译（上海）科技有限公司",
        [("540103", "翻译语言服务", 15.0, None)],
    )
    _write_assist(
        tmp_path / "sh_d.xlsx",
        "甲骨易智译（上海）科技有限公司",
        [("540103", "翻译语言服务", "大客户", 15.0, None)],
    )
    out = tmp_path / "out.xlsx"
    _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    ws = openpyxl.load_workbook(out)["损益表"]
    layout = load_layout()
    row = account_row_map(layout)["540103"]
    ka = next(letter for name, letter in dept_columns(layout) if name == "KA")
    assert ws[f"{ka}{row}"].value == 25.0
    assert ws[f"E{row}"].value == 10
    assert ws[f"G{row}"].value == 15


def test_prior_period_profit_export_goes_to_last_month_not_current(tmp_path: Path):
    path = tmp_path / "wh_pl.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "利润表"
    ws["A1"] = "2026年7期利润表（月报）"
    ws["A2"] = "公司名称：北京甲骨易文化传媒有限公司"
    ws["A3"] = "项目"
    ws["B3"] = "本月金额"
    ws["A4"] = "一、营业收入"
    ws["B4"] = 40.0
    wb.save(path)
    _write_account(
        tmp_path / "wh.xlsx",
        "北京甲骨易文化传媒有限公司",
        [("5101", "主营业务收入", None, 10.0)],
    )
    out = tmp_path / "out.xlsx"
    _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    ws = openpyxl.load_workbook(out)["利润表"]
    layout = load_layout()
    labels = [row["label"] for row in layout["profit_rows"]]
    r = 2 + labels.index("收入")
    assert ws.cell(r, 3).value in (None, "")
    assert ws.cell(r, 11).value == 40.0


def test_filename_timestamp_is_not_report_period():
    from common import detect_period_text

    assert detect_period_text("期间：202608-202608") == "202608"
    assert detect_period_text("2026年7期利润表") == "202607"
    assert detect_period_text("核算项目余额表-20260901190115.xlsx") is None


def test_discover_input_dir_uses_explicit(tmp_path: Path):
    from common import discover_input_dir

    (tmp_path / "a.xlsx").write_bytes(b"not-xlsx")
    got = discover_input_dir(str(tmp_path))
    assert got == tmp_path


def test_leaf_code_folds_to_layout_parent(tmp_path: Path):
    _write_account(
        tmp_path / "wh.xlsx",
        "北京甲骨易文化传媒有限公司",
        [("510103", "主营业务收入_翻译语言服务", None, 30.0)],
    )
    out = tmp_path / "out.xlsx"
    _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    wb = openpyxl.load_workbook(out, data_only=False)
    layout = load_layout()
    r = account_row_map(layout)["5101"]
    assert wb["损益表"][f"F{r}"].value == 30.0
    report = out.with_name(out.stem + "_运行报告.txt").read_text(encoding="utf-8")
    assert "表外科目=510103" not in report


def test_xingchen_two_row_period_header(tmp_path: Path):
    path = tmp_path / "xingchen_assist.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "核算项目余额表"
    ws["A1"] = "核算项目余额表"
    ws["A2"] = "公司名称：北京甲骨易文化传媒有限公司"
    ws["A4"] = "期间"
    ws["B4"] = "部门编码"
    ws["C4"] = "部门名称"
    ws["D4"] = "科目编码"
    ws["E4"] = "科目名称"
    ws["H4"] = "本期发生额"
    ws["H5"] = "借方"
    ws["I5"] = "贷方"
    ws["A6"] = "202608"
    ws["C6"] = "本公司"
    ws["D6"] = "510103"
    ws["E6"] = "翻译语言服务"
    ws["H6"] = 12.5
    ws["I6"] = 30.0
    wb.save(path)
    sys.modules.pop("inspect_inputs", None)
    sys.modules.pop("parse_export", None)
    if str(SCRIPTS) in sys.path:
        sys.path.remove(str(SCRIPTS))
    sys.path.insert(0, str(SCRIPTS))
    from inspect_inputs import inspect_dir, header_map
    from layout import load_export_aliases
    from parse_export import parse_inspected

    aliases = load_export_aliases()
    headers = header_map(openpyxl.load_workbook(path).active, aliases)
    assert headers["period_debit"] == (5, 8)
    assert headers["period_credit"] == (5, 9)
    found = inspect_dir(tmp_path)
    assert found and found[0]["kind"] == "assist"
    assert found[0]["entity"] == "文化"
    parsed = parse_inspected(found)
    rows = [r for r in parsed["depts"] if r["code"] == "510103"]
    assert rows
    assert rows[0]["debit"] == Decimal("12.5")
    assert rows[0]["credit"] == Decimal("30.0")


def test_layout_json_has_structure_no_gold_amounts():
    layout = json.loads((CONFIG / "版式.json").read_text(encoding="utf-8"))
    assert layout["freeze"] == "C3"
    assert layout["entities"][:3] == ["甲骨易", "文化", "上海"]
    assert "山东分公司" in layout["no_xingchen"]
    codes = [a["code"] for a in layout["accounts"]]
    assert "5101" in codes and "5401" in codes and "5502" in codes
    text = (CONFIG / "版式.json").read_text(encoding="utf-8")
    assert "8166971" not in text
    kids = direct_children(layout)
    assert "540204" in kids["5402"]
    assert "540205" in kids["5402"]
    assert "550122" in kids["5501"]


def test_two_sheets_freeze_order_and_no_confirm(tmp_path: Path):
    _write_account(
        tmp_path / "foo.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [("5101", "主营业务收入", None, 100.0), ("540103", "翻译语言服务", 40.0, None)],
    )
    _write_assist(
        tmp_path / "bar.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [
            ("5101", "主营业务收入", "大客户", None, 100.0),
            ("540103", "翻译语言服务", "大客户", 40.0, None),
        ],
    )
    out = tmp_path / "out.xlsx"
    r = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    assert out.is_file(), r.stdout + r.stderr
    wb = openpyxl.load_workbook(out)
    assert set(wb.sheetnames) == {"损益表", "利润表"}
    assert "确认情况" not in wb.sheetnames
    ws = wb["损益表"]
    assert ws.freeze_panes == "C3"
    assert ws["C1"].value == "甲骨易"
    assert ws["E1"].value == "文化"
    assert ws["I1"].value == "山东分公司"
    assert ws["Q1"].value == "济南子公司"
    depts = [ws.cell(2, c).value for c in range(21, 47)]
    assert depts[0] == "营销总监及助理"
    assert depts[1] == "KA"
    assert depts[3] == "本地化"
    assert depts[8] == "本地化"


def test_shandong_sichuan_jinan_empty_not_zero(tmp_path: Path):
    _write_account(
        tmp_path / "hq.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [("540103", "翻译语言服务", 10.0, None)],
    )
    out = tmp_path / "out.xlsx"
    _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    ws = openpyxl.load_workbook(out)["损益表"]
    layout = load_layout()
    r = account_row_map(layout)["540103"]
    assert ws[f"C{r}"].value == 10
    assert ws[f"I{r}"].value is None
    assert ws[f"O{r}"].value is None
    assert ws[f"Q{r}"].value is None
    report = (tmp_path / "out_运行报告.txt").read_text(encoding="utf-8")
    assert "山东分公司" in report
    assert "四川分公司" in report
    assert "济南子公司" in report


def test_map_dakehu_to_ka_and_unmapped_stays_out(tmp_path: Path):
    _write_account(
        tmp_path / "a.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [("540103", "翻译语言服务", 80.0, None)],
    )
    _write_assist(
        tmp_path / "b.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [
            ("540103", "翻译语言服务", "大客户", 50.0, None),
            ("540103", "翻译语言服务", "神秘事业部", 30.0, None),
        ],
    )
    out = tmp_path / "out.xlsx"
    r = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    ws = openpyxl.load_workbook(out)["损益表"]
    layout = load_layout()
    row = account_row_map(layout)["540103"]
    ka_col = None
    for name, letter in dept_columns(layout):
        if name == "KA":
            ka_col = letter
            break
    assert ka_col
    assert ws[f"{ka_col}{row}"].value == 50
    for name, letter in dept_columns(layout):
        val = ws[f"{letter}{row}"].value
        if name != "KA":
            assert val in (None, "") or (isinstance(val, str) and val.startswith("="))
            if val not in (None, "") and not str(val).startswith("="):
                raise AssertionError(name)
    report = (tmp_path / "out_运行报告.txt").read_text(encoding="utf-8")
    assert "神秘事业部" in report
    assert "50" not in r.stdout
    assert "80" not in r.stdout
    assert "30.00" not in r.stdout


def test_stdout_has_no_amount_or_secret(tmp_path: Path):
    secret = "SUPERSECRET_TOKEN_XYZ"
    fake = tmp_path / "kingdee.local.json"
    fake.write_text(json.dumps({"client_id": "1", "client_secret": secret, "app_key": "k", "app_secret": secret}), encoding="utf-8")
    _write_account(
        tmp_path / "a.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [("540103", "翻译语言服务", 12345.67, None)],
    )
    out = tmp_path / "out.xlsx"
    r = _run(
        ["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"],
        env={"KINGDEE_LOCAL_JSON": str(fake)},
    )
    assert "12345.67" not in r.stdout
    assert secret not in r.stdout
    assert "期间=202608" in r.stdout
    assert "产物=" in r.stdout


def test_company_name_line_beats_related_party_in_rows(tmp_path: Path):
    path = tmp_path / "hunan_zgs.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "科目余额表"
    ws["A1"] = "科目余额表"
    ws["A2"] = "公司名称：甲骨易（湖南）科技有限公司"
    ws["A3"] = "科目编码"
    ws["B3"] = "科目名称"
    ws["C3"] = "本期发生借方"
    ws["D3"] = "本期发生贷方"
    ws["A4"] = "3001"
    ws["B4"] = "甲骨易（北京）语言科技股份有限公司"
    ws["C4"] = 1.0
    wb.save(path)
    sys.modules.pop("inspect_inputs", None)
    if str(SCRIPTS) in sys.path:
        sys.path.remove(str(SCRIPTS))
    sys.path.insert(0, str(SCRIPTS))
    from inspect_inputs import inspect_dir

    found = inspect_dir(tmp_path)
    assert found
    assert found[0]["entity"] == "湖南子公司"


def test_fetched_book_without_pl_codes_counts_as_source(tmp_path: Path):
    path = tmp_path / "hunan.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "科目余额表"
    ws["A1"] = "科目余额表"
    ws["A2"] = "公司名称：甲骨易（北京）语言科技股份有限公司湖南分公司"
    ws["A3"] = "科目编码"
    ws["B3"] = "科目名称"
    ws["C3"] = "本期发生借方"
    ws["D3"] = "本期发生贷方"
    ws["A4"] = "1002"
    ws["B4"] = "银行存款"
    ws["C4"] = 1.0
    wb.save(path)
    out = tmp_path / "out.xlsx"
    r = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    assert "有源账套=" in r.stdout
    assert "湖南分公司" in r.stdout.split("有源账套=")[1].split("\n")[0]
    assert "已取但无损益科目=湖南分公司" in (tmp_path / "out_运行报告.txt").read_text(encoding="utf-8")


def test_no_key_with_export_still_builds(tmp_path: Path):
    _write_account(
        tmp_path / "wenhua.xlsx",
        "北京甲骨易文化传媒有限公司",
        [("5101", "主营业务收入", None, 20.0)],
    )
    out = tmp_path / "out.xlsx"
    r = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    assert out.is_file()
    assert set(openpyxl.load_workbook(out).sheetnames) == {"损益表", "利润表"}
    assert "Client" not in r.stdout
    assert "secret" not in r.stdout.lower() or "[secret]" in r.stdout.lower()


def test_formulas_not_dead_numbers_balanced_and_unbalanced(tmp_path: Path):
    _write_account(
        tmp_path / "bal.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [("540103", "翻译语言服务", 80.0, None), ("540109", "工资", 20.0, None)],
    )
    _write_assist(
        tmp_path / "bal_d.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [
            ("540103", "翻译语言服务", "KA", 80.0, None),
            ("540109", "工资", "KA", 20.0, None),
        ],
    )
    out = tmp_path / "bal_out.xlsx"
    _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    wb = openpyxl.load_workbook(out, data_only=False)
    layout = load_layout()
    row = account_row_map(layout)["540103"]
    parent = account_row_map(layout)["5401"]
    assert str(wb["损益表"][f"S{row}"].value).startswith("=")
    assert str(wb["损益表"][f"AU{row}"].value).startswith("=")
    assert str(wb["损益表"][f"AV{row}"].value).startswith("=")
    assert str(wb["损益表"][f"U{parent}"].value).startswith("=")
    assert str(wb["利润表"]["R2"].value).startswith("=")
    book, _ = eval_workbook(wb)
    av = book.cell_value("损益表", row, 48)
    assert abs(Decimal(str(av))) <= Decimal("0.05")
    parent_u = book.cell_value("损益表", parent, 21)
    child_u = book.cell_value("损益表", row, 21)
    wage = account_row_map(layout)["540109"]
    wage_u = book.cell_value("损益表", wage, 21)
    # parent 5401 U = sum of direct children; KA is column V=22 not U
    ka_parent = book.cell_value("损益表", parent, 22)
    ka_child = book.cell_value("损益表", row, 22)
    ka_wage = book.cell_value("损益表", wage, 22)
    assert Decimal(str(ka_parent)) == Decimal(str(ka_child or 0)) + Decimal(str(ka_wage or 0))

    bad = tmp_path / "bad"
    bad.mkdir()
    _write_account(bad / "a.xlsx", "甲骨易（北京）语言科技股份有限公司", [("540103", "翻译语言服务", 80.0, None)])
    _write_assist(bad / "d.xlsx", "甲骨易（北京）语言科技股份有限公司", [("540103", "翻译语言服务", "KA", 50.0, None)])
    out2 = bad / "out.xlsx"
    _run(["--period", "202608", "--input-dir", str(bad), "--out", str(out2), "--no-api"])
    wb2 = openpyxl.load_workbook(out2, data_only=False)
    book2, _ = eval_workbook(wb2)
    av2 = book2.cell_value("损益表", account_row_map(layout)["540103"], 48)
    assert abs(Decimal(str(av2))) > Decimal("0.05")


def test_change_rate_formula_not_gold_wrong_ref(tmp_path: Path):
    _write_account(tmp_path / "a.xlsx", "甲骨易（北京）语言科技股份有限公司", [("5101", "主营业务收入", None, 10.0)])
    out = tmp_path / "out.xlsx"
    _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    ws = openpyxl.load_workbook(out)["利润表"]
    layout = load_layout()
    labels = [row["label"] for row in layout["profit_rows"]]
    r = 2 + labels.index("营业收入变动率")
    formula = str(ws.cell(r, 2).value)
    assert formula.startswith("=")
    assert "J22" not in formula
    assert "J2" in formula


def test_parent_includes_children_gold_missed():
    layout = load_layout()
    kids = direct_children(layout)
    assert "550122" in kids["5501"]
    src = (SCRIPTS / "convert.py").read_text(encoding="utf-8")
    assert "direct_children" in src


def test_does_not_import_posting_picker():
    text = (SCRIPTS / "kingdee_client.py").read_text(encoding="utf-8")
    assert "kingdee-posting" not in text
    assert "kingdee_api" not in text
    posting = POSTING_API.read_text(encoding="utf-8")
    assert "多账套同时授权时只认本机已绑的总部账套" in posting


def test_profit_label_fuzzy_match():
    from fetch_reports import _match_profit_label, map_profit_rows
    from layout import load_layout

    layout = load_layout()
    aliases = layout.get("profit_item_aliases") or {}
    assert _match_profit_label("一、营业收入", aliases) == "收入"
    assert _match_profit_label("投资收益（损失以“-”号填列）", aliases) == "+投资收益"
    assert _match_profit_label("减：营业成本", aliases) == "成本"
    rows = [
        {"item_name": "一、营业收入", "current_amount": 10},
        {"item_name": "公允价值变动收益（损失以“-”号填列）", "current_amount": 1},
    ]
    mapped = map_profit_rows(rows, layout)
    assert mapped["收入"] is not None
    assert mapped["+公允价值变动收益"] is not None


def test_default_period_august_when_september():
    from common import default_period
    from datetime import date

    assert default_period(date(2026, 9, 1)) == "202608"


def test_parent_formula_includes_gold_missed_children(tmp_path: Path):
    _write_account(tmp_path / "a.xlsx", "甲骨易（北京）语言科技股份有限公司", [("540204", "印花税", 1.0, None)])
    out = tmp_path / "out.xlsx"
    _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    ws = openpyxl.load_workbook(out)["损益表"]
    layout = load_layout()
    parent = account_row_map(layout)["5402"]
    child = account_row_map(layout)["540204"]
    formula = str(ws[f"U{parent}"].value)
    assert formula.startswith("=")
    assert f"U{child}" in formula
    extra = account_row_map(layout)["540205"]
    assert f"U{extra}" in formula


def test_pl_picker_skips_empty_book():
    from kingdee_client import pick_pl_authorize_rows

    rows = [
        {"status": 1, "accountId": "1783670301378479516", "serviceId": "7914379139221", "appSecret": "empty"},
        {"status": 1, "accountId": "1783803326505631821", "serviceId": "795589109148", "appSecret": "hq"},
    ]
    got = pick_pl_authorize_rows(rows, {"account_id": "1783803326505631821"})
    assert len(got) == 1
    assert got[0]["serviceId"] == "795589109148"


def test_fetch_balance_400_falls_back_to_voucher(monkeypatch):
    import fetch_reports
    import kingdee_client

    class FakeResp:
        def __init__(self, status, payload):
            self.status_code = status
            self._payload = payload

        def json(self):
            return self._payload

    def fake_request(method, path, creds, params=None, extra=None, timeout=30, retries=4):
        if path.endswith("profit_report"):
            return FakeResp(200, {"data": [{"item_name": "营业收入", "current_amount": 1}]})
        if path.endswith("account_balance_report"):
            return FakeResp(400, {"errcode": 400})
        if path.endswith("/fi/voucher"):
            return FakeResp(200, {"data": {"count": 0, "rows": []}})
        if path.endswith("department"):
            return FakeResp(200, {"data": []})
        return FakeResp(404, {})

    monkeypatch.setattr(kingdee_client, "request", fake_request)
    monkeypatch.setattr(fetch_reports, "request", fake_request)
    monkeypatch.setattr(fetch_reports, "get_app_token", lambda creds: ("tok", "https://example"))
    monkeypatch.setenv("PL_DEPT_USE_CACHE", "0")
    got = fetch_reports.fetch_hq("202608", {"client_id": "x"})
    assert any("400" in n for n in got["notes"])
    assert any("voucher_count" in n for n in got["notes"])
    assert "profit_rows=1" in got["notes"]


def test_parent_and_leaf_same_amount_does_not_double_count(tmp_path: Path):
    _write_account(
        tmp_path / "wh.xlsx",
        "北京甲骨易文化传媒有限公司",
        [
            ("5101", "主营业务收入", None, 249313.21),
            ("510103", "主营业务收入_翻译语言服务", None, 249313.21),
        ],
    )
    out = tmp_path / "out.xlsx"
    ran = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    wb = openpyxl.load_workbook(out, data_only=False)
    layout = load_layout()
    r = account_row_map(layout)["5101"]
    assert wb["损益表"][f"F{r}"].value == 249313.21
    report = out.with_name(out.stem + "_运行报告.txt").read_text(encoding="utf-8")
    assert "表外科目=510103" not in report
    assert "249313.21" not in ran.stdout


def test_same_code_different_name_does_not_write_layout_row(tmp_path: Path):
    _write_account(
        tmp_path / "wh.xlsx",
        "北京甲骨易文化传媒有限公司",
        [
            ("540103", "住房公积金", 350.0, None),
            ("550104", "工资", 100.0, None),
            ("550204", "房租", 200.0, None),
            ("5503", "财务费用", 10.0, None),
        ],
    )
    _write_account(
        tmp_path / "sh.xlsx",
        "甲骨易智译（上海）科技有限公司",
        [("5504", "销售费用", 50.0, None)],
    )
    out = tmp_path / "out.xlsx"
    r = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    assert out.is_file(), r.stdout + r.stderr
    ws = openpyxl.load_workbook(out)["损益表"]
    layout = load_layout()
    row = account_row_map(layout)
    assert ws[f"E{row['540103']}"].value in (None, "")
    assert ws[f"E{row['540112']}"].value == 350
    assert ws[f"E{row['550104']}"].value in (None, "")
    assert ws[f"E{row['550101']}"].value == 100
    assert ws[f"E{row['550204']}"].value in (None, "")
    assert ws[f"E{row['550212']}"].value == 200
    assert ws[f"E{row['5503']}"].value in (None, "")
    assert ws[f"E{row['5504']}"].value == 10
    assert ws[f"G{row['5504']}"].value in (None, "")
    assert ws[f"G{row['5501']}"].value == 50
    report = out.with_name(out.stem + "_运行报告.txt").read_text(encoding="utf-8")
    assert "同码不同名=540103->540112" in report
    assert "350" not in r.stdout
    assert "249313.21" not in r.stdout


def _yaml_description(skill_md: Path) -> str:
    text = skill_md.read_text(encoding="utf-8")
    parts = text.split("---", 2)
    assert len(parts) >= 3, skill_md
    return parts[1]


def test_two_trigger_phrases_do_not_steal_dept_expense_alloc():
    pl_md = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    pl_yaml = _yaml_description(SKILL / "SKILL.md")
    other_md = (SKILL.parent / "dept-expense-alloc" / "SKILL.md").read_text(encoding="utf-8")
    other_yaml = _yaml_description(SKILL.parent / "dept-expense-alloc" / "SKILL.md")
    for phrase in ("月度损益表", "科目余额表"):
        assert phrase in pl_yaml
        assert phrase not in other_yaml
    assert "金蝶" in pl_yaml
    assert "用友" not in pl_yaml
    assert "pl-dept-report" in other_md
    assert "dept-expense-alloc" in pl_md
    assert "部门费用归集" in other_yaml


def test_skill_forbids_adhoc_openpyxl_and_default_skips_input_dir():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    assert "python -c" in text
    assert "import openpyxl" in text
    assert "convert.py" in text
    assert "不要**加 `--input-dir`" in text or "不要加 `--input-dir`" in text
    assert "hq.xlsx" in text
    yaml = _yaml_description(SKILL / "SKILL.md")
    assert "月度损益表" in yaml and "科目余额表" in yaml


def test_shanghai_and_wenhua_bengongsi_special_mapping(tmp_path: Path):
    _write_account(
        tmp_path / "sh_acc.xlsx",
        "甲骨易智译（上海）科技有限公司",
        [("550101", "工资", 40.0, None), ("5101", "主营业务收入", None, 80.0), ("550212", "房租", 9.0, None)],
    )
    _write_assist(
        tmp_path / "sh_dept.xlsx",
        "甲骨易智译（上海）科技有限公司",
        [
            ("550101", "工资", "本公司", 40.0, None),
            ("5101", "主营业务收入", "本公司", None, 80.0),
            ("550212", "房租", "本公司", 9.0, None),
        ],
    )
    _write_account(
        tmp_path / "wh_acc.xlsx",
        "北京甲骨易文化传媒有限公司",
        [
            ("550111", "差旅费", 12.0, None),
            ("550201", "工资", 30.0, None),
            ("540109", "工资", 15.0, None),
        ],
    )
    _write_assist(
        tmp_path / "wh_dept.xlsx",
        "北京甲骨易文化传媒有限公司",
        [
            ("550111", "差旅费", "本公司", 12.0, None),
            ("550201", "工资", "本公司", 30.0, None),
            ("540109", "工资", "本公司", 15.0, None),
        ],
    )
    _write_assist(
        tmp_path / "hq_dept.xlsx",
        "甲骨易（北京）语言科技股份有限公司",
        [("550101", "工资", "大客户", 7.0, None)],
    )
    out = tmp_path / "out.xlsx"
    r = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    assert out.is_file(), r.stdout + r.stderr
    ws = openpyxl.load_workbook(out)["损益表"]
    layout = load_layout()
    row = account_row_map(layout)
    local = dept_col_letter(layout, "本地化", 1)
    second_local = dept_col_letter(layout, "本地化", 2)
    shi = dept_col_letter(layout, "视听", 1)
    yun = dept_col_letter(layout, "运营保障中心", 1)
    fan = dept_col_letter(layout, "翻译中心", 1)
    ka = dept_col_letter(layout, "KA", 1)
    assert local and shi and yun and fan and ka and second_local
    assert ws[f"{local}{row['550101']}"].value == 40
    assert ws[f"{second_local}{row['550101']}"].value in (None, "")
    assert ws[f"{shi}{row['550111']}"].value == 12
    assert ws[f"{yun}{row['550201']}"].value == 30
    assert ws[f"{fan}{row['540109']}"].value == 15
    assert ws[f"{ka}{row['550101']}"].value == 7
    assert ws[f"{local}{row['5101']}"].value in (None, "")
    assert ws[f"{local}{row['550212']}"].value in (None, "")
    report = out.with_name(out.stem + "_运行报告.txt").read_text(encoding="utf-8")
    assert "上海:本公司" in report or "本公司" in report


def test_monthly_excel_fills_shandong(tmp_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "损益表"
    ws["C1"] = "甲骨易"
    ws["E1"] = "文化"
    ws["G1"] = "上海"
    ws["I1"] = "山东分公司"
    ws["K1"] = "湖南分公司"
    ws["M1"] = "湖南子公司"
    ws["O1"] = "四川分公司"
    ws["Q1"] = "济南子公司"
    ws["A2"] = "科目编码"
    ws["B2"] = "科目名称"
    ws["I2"] = "本期发生借方"
    ws["J2"] = "本期发生贷方"
    ws["O2"] = "本期发生借方"
    ws["A3"] = "540109"
    ws["B3"] = "工资"
    ws["I3"] = 88.0
    ws["O3"] = 22.0
    profit = wb.create_sheet("利润表")
    profit["A1"] = "项目"
    profit["E1"] = "山东26年 8月"
    profit["H1"] = "四川26年 8月"
    profit["A2"] = "成本"
    profit["E2"] = 88.0
    profit["H2"] = 22.0
    wb.save(tmp_path / "2026年8月损益类部门科目余额表.xlsx")
    wb.close()
    _write_account(tmp_path / "hq.xlsx", "甲骨易（北京）语言科技股份有限公司", [("5101", "主营业务收入", None, 1.0)])
    out = tmp_path / "out.xlsx"
    r = _run(["--period", "202608", "--input-dir", str(tmp_path), "--out", str(out), "--no-api"])
    assert out.is_file(), r.stdout + r.stderr
    ws = openpyxl.load_workbook(out)["损益表"]
    layout = load_layout()
    row = account_row_map(layout)
    assert ws[f"I{row['540109']}"].value == 88
    assert ws[f"O{row['540109']}"].value == 22
    report = out.with_name(out.stem + "_运行报告.txt").read_text(encoding="utf-8")
    assert "山东分公司" in report.split("有源账套=")[1].split("\n")[0]


def test_default_desktop_dir_helper(tmp_path: Path):
    from datetime import date
    from common import default_desktop_dir

    desktop = tmp_path / "Desktop"
    desktop.mkdir()
    got = default_desktop_dir("月度损益表", today=date(2026, 9, 2), home=tmp_path)
    assert got == desktop / "月度损益表_20260902"
