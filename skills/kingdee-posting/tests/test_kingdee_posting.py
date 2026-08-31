#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金蝶入账 · 合成回归。真表不进仓。"""
from __future__ import annotations

import json
import sys
import time
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

import importlib.util

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
SCRIPTS = SKILL / "scripts"


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


convert = _load("kingdee_posting_convert", SCRIPTS / "convert.py")
inspect_inputs = _load("kingdee_posting_inspect", SCRIPTS / "inspect_inputs.py")
kingdee_api = _load("kingdee_posting_api", SCRIPTS / "kingdee_api.py")


def _master():
    return {
        "employee": [{"code": "103", "name": "于占国"}, {"code": "011", "name": "陈霞"}, {"code": "113", "name": "项目总监"}],
        "department": [{"code": "15", "name": "本地化事业部"}, {"code": "0405", "name": "项目总监及助理"}],
        "customer": [
            {"code": "1001", "name": "甲科技有限公司"},
            {"code": "2001", "name": "国广国际在线网络（北京）有限公司"},
            {"code": "2002", "name": "中国广播电影电视交易中心"},
            {"code": "2003", "name": "商务部培训中心（商务部国际商务官员研修学院）"},
        ],
        "supplier": [{"code": "8001", "name": "北京某翻译店"}, {"code": "9999", "name": "其他供应商"}],
    }


def _lookups(extra_customers=None, customer_lines=None, receipt_sales=None, order_sales=None, period_debit=None):
    lines = customer_lines if customer_lines is not None else {
        "甲科技有限公司": ["ICT"],
        "国广国际在线网络（北京）有限公司": ["ICT"],
        "中国广播电影电视交易中心": ["ICT"],
        "商务部培训中心（商务部国际商务官员研修学院）": ["ICT"],
    }
    for name in extra_customers or []:
        lines.setdefault(name, ["ICT"])
    return {
        "customer_lines": lines,
        "receipt_sales": receipt_sales
        if receipt_sales is not None
        else [
            {"customer": "甲科技有限公司", "date": "2026-08-01", "amount": "10.00", "sales": ["于占国"]},
            {"customer": "国广国际在线网络（北京）有限公司", "date": "2026-08-01", "amount": "20.00", "sales": ["没有这个人"]},
            {"customer": "中国广播电影电视交易中心", "date": "2026-08-01", "amount": "10.00", "sales": ["于占国"]},
        ],
        "order_sales": order_sales if order_sales is not None else {"甲科技有限公司": ["于占国"]},
        "period_debit": period_debit or [],
    }


def _run(tmp_path, scene, master=None, lookups=None):
    return convert.run_dir(tmp_path, scene, "2026-08-27", _master() if master is None else master, lookups if lookups is not None else _lookups())


def _dummy_pdf(path: Path):
    path.write_bytes(b"%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n")


def _write_sales(path: Path, rows, org=None, with_org=True, headers=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "发票"
    ws.append(
        headers
        or [
            "日期",
            "发票类型",
            "发票号",
            "单位名称",
            "价税合计",
            "金额",
            "税额",
            "申请人",
            "部门编码",
            "应收账款编码",
            "主营业务收入编码",
        ]
    )
    for r in rows:
        ws.append(r)
    if with_org:
        org_ws = wb.create_sheet("组织架构")
        org_ws.append(["姓名", "部门编码"])
        for pair in org or [("于占国", "15"), ("陈霞", "15")]:
            org_ws.append(list(pair))
    wb.save(path)
    wb.close()


def _write_pay(path: Path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "付款"
    ws.append(["供应商", "应付金额本币", "开户名"])
    for r in rows:
        ws.append(r)
    wb.save(path)
    wb.close()


def _write_receipt(path: Path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "收款"
    ws.append(["日期", "客户名称", "借方（增加）", "销售", "部门编码", "应收账款编码"])
    for r in rows:
        ws.append(r)
    wb.save(path)
    wb.close()


def _ok_sales(name="甲科技有限公司", app="于占国", tot=1060, amt=1000, tax=60, ar="113103", rev="510103", typ="专票", inv="1", day="2026-08-01", dept="15"):
    return [day, typ, inv, name, tot, amt, tax, app, dept, ar, rev]


def _voucher_nums(path, n_entries):
    kd = load_workbook(path)
    ws = kd[convert.KINGDEE_SHEET]
    nums = []
    for row in ws.iter_rows(min_row=4, max_row=3 + n_entries, max_col=3, values_only=True):
        nums.append(row[2])
    kd.close()
    return nums


def test_inspect_empty(tmp_path):
    report = inspect_inputs.inspect_dir(tmp_path)
    assert report["ready"] is False
    assert inspect_inputs.main(["--input-dir", str(tmp_path)]) == 2


def test_inspect_mixed_asks(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()])
    _write_receipt(tmp_path / "收款.xlsx", [["2026-08-01", "甲", 10, "于占国", "15", "113101"]])
    report = inspect_inputs.inspect_dir(tmp_path)
    assert report["ready"] is False
    assert report["mixed"] is True


def test_inspect_payment_missing_pdf(tmp_path):
    _write_pay(tmp_path / "付款.xlsx", [["北京某翻译店", 100, "北京某翻译店"]])
    report = inspect_inputs.inspect_dir(tmp_path, "付款")
    assert report["ready"] is False
    assert any("发票" in m for m in report["missing"])


def test_sales_no_org_uses_applicant_dept(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    assert str(ws.cell(4, 22).value) == "15"
    kd.close()


def test_sales_no_invoice_no_column(tmp_path):
    headers = [
        "日期",
        "发票类型",
        "单位名称",
        "价税合计",
        "金额",
        "税额",
        "申请人",
        "部门编码",
        "应收账款编码",
        "主营业务收入编码",
    ]
    row = ["2026-08-01", "专票", "甲科技有限公司", 1060, 1000, 60, "于占国", "15", "113103", "510103"]
    _write_sales(tmp_path / "发票.xlsx", [row], headers=headers, with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1


def test_sales_no_business_line_holds(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()])
    result = _run(tmp_path, "销项发票", lookups=_lookups(customer_lines={}))
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1


def test_sales_unknown_customer_holds(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(name="不存在客户甲有限公司")])
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1


def test_sales_tax_no_aux_and_balance(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()])
    result = _run(tmp_path, "销项发票")
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    lines = list(ws.iter_rows(min_row=4, max_row=6, max_col=25, values_only=True))
    debit = sum(Decimal(str(r[15] or 0)) for r in lines)
    credit = sum(Decimal(str(r[16] or 0)) for r in lines)
    assert debit == credit == Decimal("1060.00")
    assert lines[2][6] == "21710105"
    assert lines[2][17] in (None, "")
    assert str(lines[0][0]) == "2026-08-27"
    kd.close()


def test_sales_pack_keeps_consecutive(tmp_path):
    rows = [_ok_sales(name=f"散户{i}", inv=str(i)) for i in range(5)]
    rows += [_ok_sales(name="连号甲", inv=f"a{i}") for i in range(3)]
    rows += [_ok_sales(name="连号乙", inv=f"b{i}") for i in range(4)]
    _write_sales(tmp_path / "发票.xlsx", rows)
    master = _master()
    names = list(dict.fromkeys(row[3] for row in rows))
    master["customer"] += [{"code": str(3000 + i), "name": name} for i, name in enumerate(names)]
    result = convert.run_dir(tmp_path, "销项发票", "2026-08-27", master, _lookups(extra_customers=names))
    nums = _voucher_nums(result["kingdee_path"], 12 * 3)
    assert nums.count(1) == 15
    assert nums.count(2) == 21
    assert 3 not in nums


def test_sales_no_master_holds(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    result = convert.run_dir(tmp_path, "销项发票", "2026-08-27", None, _lookups())
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1


def test_payment_special_and_normal(tmp_path, monkeypatch):
    _write_pay(tmp_path / "付款.xlsx", [["北京某翻译店", 106, "北京某翻译店"], ["无档店", 200, "无档店"]])
    a = tmp_path / "北京某翻译店"
    b = tmp_path / "无档店"
    a.mkdir()
    b.mkdir()
    _dummy_pdf(a / "a.pdf")
    _dummy_pdf(b / "b.pdf")

    def fake_parse(path: Path):
        if path.parent.name == "北京某翻译店":
            return {"kind": "专票", "seller": "北京某翻译店", "total": Decimal("106.00"), "tax": Decimal("6.00")}
        return {"kind": "普票", "seller": "无档店", "total": Decimal("200.00"), "tax": None}

    monkeypatch.setattr(convert, "parse_invoice_pdf", fake_parse)
    result = _run(tmp_path, "付款")
    assert result["bookable_count"] == 2
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    accounts = [ws.cell(r, 7).value for r in range(4, 9)]
    assert "540103" in accounts
    assert "21710101" in accounts
    assert "100206" in accounts
    # 无档 → 9999
    sup_codes = [ws.cell(r, 20).value for r in range(4, 9)]
    assert "9999" in sup_codes
    # 业务规则只给查找键；实际辅助核算值必须来自已核验的档案。
    dept_codes = [ws.cell(r, 22).value for r in range(4, 9)]
    employee_codes = [ws.cell(r, 24).value for r in range(4, 9)]
    assert "0405" in dept_codes
    assert "113" in employee_codes
    kd.close()


def test_payment_unclear_type_holds(tmp_path, monkeypatch):
    _write_pay(tmp_path / "付款.xlsx", [["北京某翻译店", 100, "北京某翻译店"]])
    folder = tmp_path / "北京某翻译店"
    folder.mkdir()
    _dummy_pdf(folder / "a.pdf")
    monkeypatch.setattr(convert, "parse_invoice_pdf", lambda p: {"kind": "", "seller": "x", "total": None, "tax": None})
    result = _run(tmp_path, "付款")
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1


def test_payment_ticket_less_than_payable_holds(tmp_path, monkeypatch):
    _write_pay(tmp_path / "付款.xlsx", [["北京某翻译店", 200, "北京某翻译店"]])
    folder = tmp_path / "北京某翻译店"
    folder.mkdir()
    _dummy_pdf(folder / "a.pdf")
    monkeypatch.setattr(
        convert,
        "parse_invoice_pdf",
        lambda p: {"kind": "普票", "seller": "北京某翻译店", "total": Decimal("100.00"), "tax": None},
    )
    result = _run(tmp_path, "付款")
    assert result["hold_count"] == 1


def test_receipt_pack_alias_pingdu_and_empty_emp(tmp_path):
    rows = []
    for i in range(10):
        rows.append(["2026-08-01", "甲科技有限公司", 10, "表内销售应忽略", "15", "113101"])
    rows.append(["2026-08-01", "国广国际在线网络（北京）有限公司陕西分公司", 20, "没有这个人", "15", "113102"])
    rows.append(["2026-08-01", "平度市公安局", 30, "于占国", "15", "113101"])
    _write_receipt(tmp_path / "收款.xlsx", rows)
    result = _run(tmp_path, "收款")
    assert result["hold_count"] == 1
    assert result["bookable_count"] == 11
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    nums = _voucher_nums(result["kingdee_path"], 22)
    assert nums.count(1) == 20
    names = [ws.cell(r, 19).value for r in range(4, 30)]
    assert "国广国际在线网络（北京）有限公司" in names
    assert "平度市公安局" not in names
    emp_codes = [ws.cell(r, 24).value for r in range(4, 30)]
    assert "103" in emp_codes
    guang_rows = [r for r in range(4, 30) if ws.cell(r, 19).value == "国广国际在线网络（北京）有限公司"]
    assert guang_rows
    assert all(ws.cell(r, 24).value in (None, "") for r in guang_rows)
    kd.close()


def test_receipt_no_master_holds(tmp_path):
    _write_receipt(tmp_path / "收款.xlsx", [["2026-08-01", "甲科技有限公司", 10, "于占国", "15", "113101"]])
    result = convert.run_dir(tmp_path, "收款", "2026-08-27", None, _lookups())
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1


def test_receipt_accepts_word_spelling_without_programme(tmp_path):
    _write_receipt(
        tmp_path / "收款.xlsx",
        [["2026-08-01", "中国广播电影电视交易中心有限公司", 10, "于占国", "15", "113101"]],
    )
    result = _run(tmp_path, "收款")
    assert result["bookable_count"] == 1
    assert result["hold_count"] == 0


def test_pick_code_uses_cached_value_not_formula():
    formula = "=VLOOKUP(D:D,[1]组织架构!A$1:B$65536,2,0)"
    assert convert.pick_code(formula, "0302") == "0302"
    assert convert.pick_code(formula, None) == ""
    assert convert.pick_code("15", None) == "15"


def test_parse_invoice_text_special():
    text = "电子发票（专用发票）\n销售方名称：北京某翻译店\n价税合计（大写）壹佰圆\n（小写）¥106.00\n税额 6.00"
    got = convert.parse_invoice_text(text)
    assert got["kind"] == "专票"
    assert got["seller"] == "北京某翻译店"
    assert got["total"] == Decimal("106.00")
    assert got["tax"] == Decimal("6.00")


def test_kingdee_api_missing_credentials(tmp_path, monkeypatch):
    monkeypatch.setenv("KINGDEE_LOCAL_JSON", str(tmp_path / "nope.json"))
    monkeypatch.setenv("KINGDEE_MASTER_CACHE", str(tmp_path / "no-cache.json"))
    loaded = kingdee_api.try_load_master()
    assert loaded["missing_credentials"] is True
    assert loaded["ok"] is False


def test_fetch_list_uses_official_api_host_not_token_domain(monkeypatch):
    seen = []

    class Response:
        status_code = 200

        @staticmethod
        def json():
            return {"data": {"rows": [], "count": "0"}}

    def fake_request(method, url, creds, path, params=None, extra_headers=None, timeout=30):
        seen.append((url, params))
        return Response()

    monkeypatch.setattr(kingdee_api, "_request", fake_request)
    got = kingdee_api.fetch_list(
        {"client_id": "test", "client_secret": "test"},
        "token",
        "https://tf.jdy.com",
        "/jdy/v2/bd/customer",
    )
    assert got == []
    assert seen == [
        (
            "https://api.kingdee.com/jdy/v2/bd/customer",
            {"page": "1", "page_size": "2000"},
        )
    ]


def test_master_uses_fresh_local_cache_before_network(tmp_path, monkeypatch):
    cache = tmp_path / "kingdee-master.json"
    cache.write_text(
        json.dumps(
            {
                "cached_at": time.time(),
                "data": {"customer": [], "employee": [], "supplier": [], "department": []},
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv("KINGDEE_MASTER_CACHE", str(cache))
    monkeypatch.setattr(kingdee_api, "load_local", lambda: (_ for _ in ()).throw(AssertionError("不应联网")))
    loaded = kingdee_api.try_load_master()
    assert loaded["ok"] is True
    assert loaded["source"] == "cache"
    assert loaded["data"] == {"customer": [], "employee": [], "supplier": [], "department": []}


def test_sign_plain_path_encoding():
    plain = kingdee_api.sign_plain(
        "GET",
        "/jdyconnector/app_management/kingdee_auth_token",
        {"app_key": "bVZgAZOv1", "app_signature": "abc=="},
        "4427456950",
        "1670305063559",
    )
    assert plain.startswith("GET\n%2Fjdyconnector%2Fapp_management%2Fkingdee_auth_token\n")
    assert "%253D%253D" in plain
    assert plain.endswith("\n")


def test_cli_inspect_then_convert(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    assert convert.main(["--inspect", "--input-dir", str(tmp_path), "--scene", "销项发票"]) == 0


def test_cli_refuses_to_create_unverified_auxiliaries_without_api(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    assert convert.main(["--input-dir", str(tmp_path), "--scene", "销项发票", "--no-api"]) == 2
    assert not (tmp_path / "凭证引入_结果.xlsx").exists()


def test_inspect_receipt_without_sales_or_ar(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "收款"
    ws.append(["日期", "客户名称", "借方（增加）", "部门编码"])
    ws.append(["2026-08-01", "甲科技有限公司", 10, "15"])
    wb.save(tmp_path / "收款.xlsx")
    wb.close()
    report = inspect_inputs.inspect_dir(tmp_path, "收款")
    assert report["ready"] is True


def test_inspect_sales_without_account_columns(tmp_path):
    headers = ["日期", "发票类型", "单位名称", "价税合计", "金额", "税额", "申请人"]
    row = ["2026-08-01", "专票", "甲科技有限公司", 1060, 1000, 60, "于占国"]
    _write_sales(tmp_path / "发票.xlsx", [row], headers=headers, with_org=False)
    report = inspect_inputs.inspect_dir(tmp_path, "销项发票")
    assert report["ready"] is True


def test_sales_ignores_order_no_and_table_accounts(tmp_path):
    headers = [
        "日期",
        "发票类型",
        "单位名称",
        "价税合计",
        "金额",
        "税额",
        "申请人",
        "下单号",
        "合同号",
        "应收账款编码",
        "主营业务收入编码",
    ]
    row = ["2026-08-01", "专票", "甲科技有限公司", 1060, 1000, 60, "于占国", "SO-WRONG", "HT-WRONG", "113101", "510101"]
    _write_sales(tmp_path / "发票.xlsx", [row], headers=headers, with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    accounts = [ws.cell(r, 7).value for r in range(4, 7)]
    assert "113103" in accounts
    assert "510103" in accounts
    assert "113101" not in accounts
    kd.close()


def test_sales_applicant_not_in_dept_table_holds(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(app="路人甲")], with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1


def test_sales_multi_line_uses_period_debit(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    result = _run(
        tmp_path,
        "销项发票",
        lookups=_lookups(
            customer_lines={"甲科技有限公司": ["ICT", "游戏综合本地化"]},
            period_debit=[
                {"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "90"},
                {"customer_code": "1001", "account": "113102", "period": "2026-08", "debit": "10"},
            ],
        ),
    )
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    accounts = [ws.cell(r, 7).value for r in range(4, 7)]
    assert "113103" in accounts
    assert "113102" not in accounts
    kd.close()


def test_receipt_uses_lookups_not_table_ar(tmp_path):
    _write_receipt(tmp_path / "收款.xlsx", [["2026-08-01", "甲科技有限公司", 10, "表内销售", "15", "113101"]])
    result = _run(tmp_path, "收款")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    accounts = [ws.cell(r, 7).value for r in range(4, 6)]
    assert "113103" in accounts
    assert "113101" not in accounts
    kd.close()


def test_receipt_order_fallback_and_dual_sales(tmp_path):
    _write_receipt(tmp_path / "收款.xlsx", [["2026-08-01", "甲科技有限公司", 10, "表内销售", "15", "113101"]])
    fallback = _run(
        tmp_path,
        "收款",
        lookups=_lookups(
            receipt_sales=[],
            order_sales={"甲科技有限公司": ["陈霞"]},
        ),
    )
    assert fallback["bookable_count"] == 1
    kd = load_workbook(fallback["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    assert "011" in [ws.cell(r, 24).value for r in range(4, 6)]
    kd.close()
    hold = convert.run_dir(
        tmp_path,
        "收款",
        "2026-08-27",
        _master(),
        _lookups(receipt_sales=[], order_sales={"甲科技有限公司": ["于占国", "陈霞"]}),
    )
    assert hold["bookable_count"] == 0
    assert hold["hold_count"] == 1
    detail = load_workbook(hold["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "斯佳" in reason


def test_cli_refuses_without_zhiyun_lookups(tmp_path, monkeypatch):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    master = tmp_path / "master.json"
    master.write_text(json.dumps(_master()), encoding="utf-8")
    monkeypatch.setattr(
        convert.zhiyun_api,
        "try_load_lookups",
        lambda: {"ok": False, "missing_credentials": True, "data": None},
    )
    assert convert.main(
        ["--input-dir", str(tmp_path), "--scene", "销项发票", "--master", str(master)]
    ) == 2
    assert not (tmp_path / "凭证引入_结果.xlsx").exists()


def test_cli_lookups_file_converts(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    master = tmp_path / "master.json"
    lookups = tmp_path / "lookups.json"
    master.write_text(json.dumps(_master()), encoding="utf-8")
    lookups.write_text(json.dumps(_lookups()), encoding="utf-8")
    assert (
        convert.main(
            [
                "--input-dir",
                str(tmp_path),
                "--scene",
                "销项发票",
                "--master",
                str(master),
                "--lookups",
                str(lookups),
                "--date",
                "2026-08-27",
            ]
        )
        == 0
    )
    assert (tmp_path / "凭证引入_结果.xlsx").exists()
