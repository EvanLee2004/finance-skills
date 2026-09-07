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
        "employee": [
            {"code": "103", "name": "于占国"},
            {"code": "011", "name": "陈霞"},
            {"code": "113", "name": "项目总监"},
            {"code": "205", "name": "郑瑞"},
        ],
        "department": [
            {"code": "15", "name": "本地化事业部"},
            {"code": "0405", "name": "项目总监及助理"},
            {"code": "0308", "name": "商务中心"},
        ],
        "customer": [
            {"code": "1001", "name": "甲科技有限公司"},
            {"code": "2001", "name": "国广国际在线网络（北京）有限公司"},
            {"code": "2002", "name": "中国广播电影电视节目交易中心"},
            {"code": "2003", "name": "商务部培训中心（商务部国际商务官员研修学院）"},
            {"code": "0582", "name": "个人"},
            {"code": "0386", "name": "公安部"},
            {"code": "1940", "name": "北京市公安局海淀分局"},
            {"code": "9001", "name": "乙科技有限公司乙科技有限公司"},
        ],
        "supplier": [{"code": "8001", "name": "北京某翻译店"}, {"code": "9999", "name": "其他供应商"}],
    }


def _lookups(extra_customers=None, customer_lines=None, receipt_sales=None, order_sales=None, period_debit=None, ar_balance=None, assist_rows=None, include_assist=True):
    lines = customer_lines if customer_lines is not None else {}
    for name in extra_customers or []:
        lines.setdefault(name, [])
    balances = ar_balance
    if balances is None:
        balances = [
            {"customer_code": "1001", "account": "113103", "balance": "1"},
            {"customer_code": "2001", "account": "113103", "balance": "1"},
            {"customer_code": "2002", "account": "113103", "balance": "1"},
            {"customer_code": "2003", "account": "113103", "balance": "1"},
            {"customer_code": "0582", "account": "113103", "balance": "1"},
            {"customer_code": "0386", "account": "113103", "balance": "1"},
            {"customer_code": "1940", "account": "113103", "balance": "1"},
            {"customer_code": "9001", "account": "113103", "balance": "1"},
        ]
        for i, _name in enumerate(extra_customers or []):
            balances.append({"customer_code": str(3000 + i), "account": "113103", "balance": "1"})
    payload = {
        "ar_accounts": ["113101", "113102", "113103", "113105", "113107"],
        "ar_balance": balances,
        "customer_lines": lines,
        "receipt_sales": receipt_sales
        if receipt_sales is not None
        else [
            {"customer": "甲科技有限公司", "date": "2026-08-01", "amount": "10.00", "sales": ["于占国"]},
            {"customer": "国广国际在线网络（北京）有限公司", "date": "2026-08-01", "amount": "20.00", "sales": ["没有这个人"]},
            {"customer": "中国广播电影电视交易中心", "date": "2026-08-01", "amount": "10.00", "sales": ["于占国"]},
            {"customer": "中国广播电影电视节目交易中心", "date": "2026-08-01", "amount": "10.00", "sales": ["于占国"]},
            {"customer": "公安部", "date": "2026-08-01", "amount": "30.00", "sales": ["于占国"]},
            {"customer": "平度市公安局", "date": "2026-08-01", "amount": "30.00", "sales": ["于占国"]},
            {"customer": "个人", "date": "2026-08-01", "amount": "10.00", "sales": ["于占国"]},
            {"customer": "北京市公安局海淀分局", "date": "2026-08-01", "amount": "10.00", "sales": ["陈霞"]},
        ],
        "order_sales": order_sales if order_sales is not None else {"甲科技有限公司": ["于占国"]},
        "period_debit": period_debit or [],
    }
    if include_assist:
        if assist_rows is not None:
            payload["assist_rows"] = assist_rows
        else:
            names_by_code = {c["code"]: c["name"] for c in _master()["customer"]}
            for i, name in enumerate(extra_customers or []):
                names_by_code.setdefault(str(3000 + i), name)
            built = []
            for b in balances:
                code = str(b.get("customer_code") or "")
                built.append(
                    {
                        "period": "202607",
                        "customer_code": code,
                        "customer_name": names_by_code.get(code, code),
                        "account": b.get("account"),
                        "ending_debit": b.get("balance"),
                        "ending_credit": None,
                        "ytd_debit": b.get("balance"),
                        "ytd_credit": None,
                    }
                )
            payload["assist_rows"] = built
    return payload


def _run(tmp_path, scene, master=None, lookups=None, start_voucher_no=1, out_dir=None):
    return convert.run_dir(
        tmp_path,
        scene,
        "2026-08-27",
        _master() if master is None else master,
        lookups if lookups is not None else _lookups(),
        start_voucher_no=start_voucher_no,
        out_dir=out_dir or tmp_path,
    )


def _dummy_pdf(path: Path):
    path.write_bytes(b"%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n")


def _write_assist(path: Path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws.cell(1, 1, "核算项目余额表")
    ws.cell(2, 1, "公司名称：测试")
    ws.cell(2, 8, "期间：202601-202612")
    headers = ["期间", "客户编码", "客户名称", "科目编码", "科目名称", "期初", "期初", "本期发生额", "本期发生额", "本年累计", "本年累计", "期末", "期末"]
    sub = ["期间", "客户编码", "客户名称", "科目编码", "科目名称", "借方", "贷方", "借方", "贷方", "借方", "贷方", "借方", "贷方"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    for i, h in enumerate(sub, 1):
        ws.cell(5, i, h)
    r = 6
    for item in rows:
        ws.cell(r, 1, item.get("period") or "202607")
        ws.cell(r, 2, item.get("customer_code"))
        ws.cell(r, 3, item.get("customer_name"))
        ws.cell(r, 4, item.get("account"))
        ws.cell(r, 10, item.get("ytd_debit"))
        ws.cell(r, 11, item.get("ytd_credit"))
        ws.cell(r, 12, item.get("ending_debit"))
        ws.cell(r, 13, item.get("ending_credit"))
        r += 1
    wb.save(path)
    wb.close()


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
    result = _run(tmp_path, "销项发票", lookups=_lookups(ar_balance=[]))
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1
    detail = load_workbook(result["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "是否新建" in reason
    assert "无业务线" not in reason


def test_sales_unknown_customer_holds(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(name="不存在客户甲有限公司")])
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1
    detail = load_workbook(result["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "请斯佳确认是否新建" in reason


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
    result = convert.run_dir(tmp_path, "销项发票", "2026-08-27", master, _lookups(extra_customers=names), out_dir=tmp_path)
    nums = _voucher_nums(result["kingdee_path"], 12 * 3)
    assert nums.count(1) == 15
    assert nums.count(2) == 21
    assert 3 not in nums


def test_sales_no_master_holds(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    result = convert.run_dir(tmp_path, "销项发票", "2026-08-27", None, _lookups(), out_dir=tmp_path)
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
    assert result["hold_count"] == 0
    assert result["bookable_count"] == 12
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    nums = _voucher_nums(result["kingdee_path"], 24)
    assert nums.count(1) == 20
    names = [ws.cell(r, 19).value for r in range(4, 40)]
    assert "国广国际在线网络（北京）有限公司" in names
    assert "公安部" in names
    assert "平度市公安局" not in names
    emp_codes = [ws.cell(r, 24).value for r in range(4, 30)]
    assert "103" in emp_codes
    guang_rows = [r for r in range(4, 30) if ws.cell(r, 19).value == "国广国际在线网络（北京）有限公司"]
    assert guang_rows
    assert all(ws.cell(r, 24).value in (None, "") for r in guang_rows)
    kd.close()


def test_receipt_no_master_holds(tmp_path):
    _write_receipt(tmp_path / "收款.xlsx", [["2026-08-01", "甲科技有限公司", 10, "于占国", "15", "113101"]])
    result = convert.run_dir(tmp_path, "收款", "2026-08-27", None, _lookups(), out_dir=tmp_path)
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


class _FakeResp:
    def __init__(self, payload, status=200):
        self.status_code = status
        self._payload = payload

    def json(self):
        return self._payload


def test_pick_authorize_row_keeps_hq_not_empty_books():
    rows = [
        {
            "status": 1,
            "accountId": "1783670301378479516",
            "serviceId": "7914379139221",
            "appKey": "8uVREFRV",
            "appSecret": "empty-secret",
        },
        {
            "status": 1,
            "accountId": "1783803326505631821",
            "serviceId": "795589109148",
            "appKey": "KZbQMo3T",
            "appSecret": "hq-secret",
        },
    ]
    creds = {
        "account_id": "1783803326505631821",
        "service_id": "795589109148",
        "app_key": "KZbQMo3T",
    }
    got = kingdee_api.pick_authorize_row(rows, creds)
    assert got is not None
    assert got["appKey"] == "KZbQMo3T"
    assert got["serviceId"] == "795589109148"


def test_get_app_token_refreshes_stale_app_secret(tmp_path, monkeypatch):
    local = tmp_path / "kingdee.local.json"
    creds = {
        "client_id": "357164",
        "client_secret": "x" * 32,
        "app_key": "KZbQMo3T",
        "app_secret": "old-secret",
        "account_id": "1783803326505631821",
        "service_id": "795589109148",
        "outer_instance_id": "572594141763080192",
    }
    local.write_text(json.dumps(creds), encoding="utf-8")
    monkeypatch.setenv("KINGDEE_LOCAL_JSON", str(local))
    paths = []

    def fake_request(method, url, used, path, params=None, extra_headers=None, timeout=30):
        paths.append((method, path))
        if path == kingdee_api.AUTH_PATH:
            if used.get("app_secret") == "new-hq-secret":
                return _FakeResp({"data": {"app-token": "tok", "domain": "https://tf.jdy.com"}})
            return _FakeResp({"errcode": 1030002006, "description": "授权密钥校验失败", "data": None})
        if path == kingdee_api.AUTHORIZE_PATH:
            return _FakeResp(
                {
                    "code": 200,
                    "data": [
                        {
                            "status": 1,
                            "accountId": "1783670301378479516",
                            "serviceId": "7914379139221",
                            "appKey": "8uVREFRV",
                            "appSecret": "empty-secret",
                            "accessToken": "do-not-save",
                        },
                        {
                            "status": 1,
                            "accountId": "1783803326505631821",
                            "serviceId": "795589109148",
                            "appKey": "KZbQMo3T",
                            "appSecret": "new-hq-secret",
                            "accountName": "总部",
                            "agreementCompanyName": "总部",
                            "domain": "https://tf.jdy.com",
                            "outerInstanceId": "572594141763080192",
                            "groupName": "ns-t33w",
                            "accessToken": "do-not-save",
                            "appToken": "do-not-save",
                        },
                    ],
                }
            )
        raise AssertionError(path)

    monkeypatch.setattr(kingdee_api, "_request", fake_request)
    token, domain = kingdee_api.get_app_token(creds)
    assert token == "tok"
    assert domain == "https://tf.jdy.com"
    saved = json.loads(local.read_text(encoding="utf-8"))
    assert saved["app_secret"] == "new-hq-secret"
    assert saved["account_id"] == "1783803326505631821"
    assert saved["service_id"] == "795589109148"
    assert "accessToken" not in saved
    assert "appToken" not in saved
    assert ( "POST", kingdee_api.AUTHORIZE_PATH) in paths
    assert paths[-1] == ("GET", kingdee_api.AUTH_PATH)


def test_voucher_ar_reads_customer_assist(tmp_path, monkeypatch):
    from decimal import Decimal

    local = tmp_path / "kingdee.local.json"
    local.write_text(
        json.dumps(
            {
                "client_id": "357164",
                "client_secret": "x" * 32,
                "app_key": "KZbQMo3T",
                "app_secret": "s" * 40,
                "account_id": "1783803326505631821",
                "service_id": "795589109148",
                "outer_instance_id": "572594141763080192",
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv("KINGDEE_LOCAL_JSON", str(local))
    monkeypatch.setenv("KINGDEE_AR_CACHE", str(tmp_path / "ar.json"))
    monkeypatch.setattr(kingdee_api, "get_app_token", lambda creds: ("tok", "https://tf.jdy.com"))

    def fake_request(method, url, used, path, params=None, extra_headers=None, timeout=30):
        if path == "/jdy/v2/fi/voucher":
            return _FakeResp({"errcode": 0, "data": {"rows": [{"id": "v1", "period": "202608"}], "count": 1}})
        if path == "/jdy/v2/fi/voucher_detail":
            return _FakeResp(
                {
                    "errcode": 0,
                    "data": {
                        "period": "202608",
                        "entry_list": [
                            {
                                "account_number": "113103",
                                "debit_amount": "80",
                                "credit_amount": "0",
                                "assist": [{"type": "bd_customer", "number": "1001"}],
                            },
                            {
                                "account_number": "113102",
                                "debit_amount": "10",
                                "credit_amount": "0",
                                "assist": [{"type": "bd_customer", "number": "1001"}],
                            },
                        ],
                    },
                }
            )
        raise AssertionError(path)

    monkeypatch.setattr(kingdee_api, "_request", fake_request)
    got = kingdee_api.try_fetch_customer_ar("1001", ["113101", "113102", "113103"], "2026-08")
    assert got["ok"] is True
    assert got["balances"][("1001", "113103")] == Decimal("80.00")
    assert got["period_debit"][("1001", "113103", "2026-08")] == Decimal("80.00")
    empty = kingdee_api.try_fetch_customer_ar("9999", ["113103"], "2026-08")
    assert empty["ok"] is True
    assert empty["balances"] == {}


def test_ar_windows_recent_then_year_start():
    assert kingdee_api.ar_windows("2026-09") == {"recent": ("202608", "202609"), "ytd": ("202601", "202609")}
    assert kingdee_api.ar_windows("2026-01") == {"recent": ("202601", "202601"), "ytd": ("202601", "202601")}
    assert kingdee_api.ar_windows("2026-09-07") == {"recent": ("202608", "202609"), "ytd": ("202601", "202609")}


def test_fetch_uses_recent_window_when_customer_present(tmp_path, monkeypatch):
    from decimal import Decimal

    local = tmp_path / "kingdee.local.json"
    local.write_text(
        json.dumps(
            {
                "client_id": "357164",
                "client_secret": "x" * 32,
                "app_key": "KZbQMo3T",
                "app_secret": "s" * 40,
                "account_id": "1783803326505631821",
                "service_id": "795589109148",
                "outer_instance_id": "572594141763080192",
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv("KINGDEE_LOCAL_JSON", str(local))
    monkeypatch.setenv("KINGDEE_AR_CACHE", str(tmp_path / "ar.json"))
    monkeypatch.setattr(kingdee_api, "get_app_token", lambda creds: ("tok", "https://tf.jdy.com"))
    windows = []

    def fake_request(method, url, used, path, params=None, extra_headers=None, timeout=30):
        if path == "/jdy/v2/fi/voucher":
            windows.append((params.get("start_period"), params.get("end_period")))
            return _FakeResp({"errcode": 0, "data": {"rows": [{"id": "v-aug", "period": "202608"}], "count": 1}})
        if path == "/jdy/v2/fi/voucher_detail":
            return _FakeResp(
                {
                    "errcode": 0,
                    "data": {
                        "period": "202608",
                        "entry_list": [
                            {
                                "account_number": "113103",
                                "debit_amount": "80",
                                "credit_amount": "0",
                                "assist": [{"type": "bd_customer", "number": "1001"}],
                            }
                        ],
                    },
                }
            )
        raise AssertionError(path)

    monkeypatch.setattr(kingdee_api, "_request", fake_request)
    got = kingdee_api.try_fetch_customer_ar("1001", ["113103"], "2026-09")
    assert got["ok"] is True
    assert got["balances"][("1001", "113103")] == Decimal("80.00")
    assert got["period_debit"] == {}
    assert windows == [("202608", "202609")]


def test_fetch_extends_to_year_start_when_recent_empty(tmp_path, monkeypatch):
    from decimal import Decimal

    local = tmp_path / "kingdee.local.json"
    local.write_text(
        json.dumps(
            {
                "client_id": "357164",
                "client_secret": "x" * 32,
                "app_key": "KZbQMo3T",
                "app_secret": "s" * 40,
                "account_id": "1783803326505631821",
                "service_id": "795589109148",
                "outer_instance_id": "572594141763080192",
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv("KINGDEE_LOCAL_JSON", str(local))
    monkeypatch.setenv("KINGDEE_AR_CACHE", str(tmp_path / "ar.json"))
    monkeypatch.setattr(kingdee_api, "get_app_token", lambda creds: ("tok", "https://tf.jdy.com"))
    windows = []

    def fake_request(method, url, used, path, params=None, extra_headers=None, timeout=30):
        if path == "/jdy/v2/fi/voucher":
            start = str((params or {}).get("start_period") or "")
            end = str((params or {}).get("end_period") or "")
            windows.append((start, end))
            if start == "202608":
                return _FakeResp({"errcode": 0, "data": {"rows": [], "count": 0}})
            if start == "202601":
                return _FakeResp({"errcode": 0, "data": {"rows": [{"id": "v-mar", "period": "202603"}], "count": 1}})
            raise AssertionError((start, end))
        if path == "/jdy/v2/fi/voucher_detail":
            return _FakeResp(
                {
                    "errcode": 0,
                    "data": {
                        "period": "202603",
                        "entry_list": [
                            {
                                "account_number": "113103",
                                "debit_amount": "50",
                                "credit_amount": "0",
                                "assist": [{"type": "bd_customer", "number": "1001"}],
                            }
                        ],
                    },
                }
            )
        raise AssertionError(path)

    monkeypatch.setattr(kingdee_api, "_request", fake_request)
    got = kingdee_api.try_fetch_customer_ar("1001", ["113103"], "2026-09")
    assert got["ok"] is True
    assert got["balances"][("1001", "113103")] == Decimal("50.00")
    assert got["period_debit"] == {}
    assert windows == [("202608", "202609"), ("202601", "202609")]


def test_repo_venv_python_exists():
    got = convert.repo_venv_python()
    assert got is not None
    assert got.is_file()
    assert Path(sys.prefix).resolve() == got.parent.parent.resolve()


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
            ar_balance=[
                {"customer_code": "1001", "account": "113103", "balance": "1"},
                {"customer_code": "1001", "account": "113102", "balance": "1"},
            ],
            period_debit=[
                {"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "90"},
                {"customer_code": "1001", "account": "113102", "period": "2026-08", "debit": "10"},
            ],
        ),
    )
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1
    detail = load_workbook(result["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "多条" in reason
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    accounts = [ws.cell(r, 7).value for r in range(4, 7)]
    kd.close()
    assert "113103" not in accounts
    assert "113102" not in accounts


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
        out_dir=tmp_path,
    )
    assert hold["bookable_count"] == 0
    assert hold["hold_count"] == 1
    detail = load_workbook(hold["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "斯佳" in reason


def test_cli_sales_runs_without_zhiyun_lookups(tmp_path, monkeypatch):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    master = tmp_path / "master.json"
    lookups = tmp_path / "lookups.json"
    master.write_text(json.dumps(_master()), encoding="utf-8")
    lookups.write_text(json.dumps(_lookups()), encoding="utf-8")
    monkeypatch.setattr(
        convert.zhiyun_api,
        "try_load_lookups",
        lambda: (_ for _ in ()).throw(AssertionError("销项不应访问智云")),
    )
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
                "--start-voucher-no",
                "1",
                "--out-dir",
                str(tmp_path),
            ]
        )
        == 0
    )
    assert (tmp_path / "凭证引入_结果.xlsx").exists()


def test_cli_receipt_refuses_without_zhiyun_lookups(tmp_path, monkeypatch):
    _write_receipt(tmp_path / "收款.xlsx", [["2026-08-01", "甲科技有限公司", 10, "于占国", "15", "113101"]])
    master = tmp_path / "master.json"
    master.write_text(json.dumps(_master()), encoding="utf-8")
    monkeypatch.setattr(
        convert.zhiyun_api,
        "try_load_lookups",
        lambda: {"ok": False, "missing_credentials": True, "data": None},
    )
    assert convert.main(
        ["--input-dir", str(tmp_path), "--scene", "收款", "--master", str(master)]
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
                "--start-voucher-no",
                "1",
                "--out-dir",
                str(tmp_path),
            ]
        )
        == 0
    )
    assert (tmp_path / "凭证引入_结果.xlsx").exists()


def test_sales_programme_peels_limited_company(tmp_path):
    _write_sales(
        tmp_path / "发票.xlsx",
        [_ok_sales(name="中国广播电影电视节目交易中心有限公司")],
        with_org=False,
    )
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    names = [ws.cell(r, 19).value for r in range(4, 7)]
    assert "中国广播电影电视节目交易中心" in names
    kd.close()


def test_sales_doubled_archive_name(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(name="乙科技有限公司")], with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1


def test_sales_person_heading_uses_personal_customer(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(name="刘芳")], with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    names = [ws.cell(r, 19).value for r in range(4, 7)]
    assert "个人" in names
    kd.close()


def test_sales_police_maps_ministry_except_haidian_chenxia(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(name="平度市公安局")], with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    names = [ws.cell(r, 19).value for r in range(4, 7)]
    assert "公安部" in names
    kd.close()
    _write_sales(
        tmp_path / "海淀.xlsx",
        [_ok_sales(name="北京市公安局海淀分局", app="陈霞")],
        with_org=False,
    )
    # mixed folder would confuse inspect; use dedicated dir via rewriting same 发票.xlsx
    _write_sales(
        tmp_path / "发票.xlsx",
        [_ok_sales(name="北京市公安局海淀分局", app="陈霞")],
        with_org=False,
    )
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    names = [ws.cell(r, 19).value for r in range(4, 7)]
    assert "北京市公安局海淀分局" in names
    kd.close()


def test_sales_hang_tong_zhao_to_zheng(tmp_path):
    _write_sales(
        tmp_path / "发票.xlsx",
        [_ok_sales(app="童睿智"), _ok_sales(app="赵贺斌", inv="2")],
        with_org=False,
    )
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 2
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    emps = [str(ws.cell(r, 24).value or "") for r in range(4, 10)]
    assert emps.count("205") >= 2
    kd.close()


def test_receipt_person_and_police(tmp_path):
    _write_receipt(
        tmp_path / "收款.xlsx",
        [
            ["2026-08-01", "刘芳", 10, "于占国", "15", "113101"],
            ["2026-08-01", "平度市公安局", 30, "于占国", "15", "113101"],
        ],
    )
    result = _run(tmp_path, "收款")
    assert result["bookable_count"] == 2
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    names = [ws.cell(r, 19).value for r in range(4, 10)]
    assert "个人" in names
    assert "公安部" in names
    kd.close()


def test_payment_peels_limited_company(tmp_path, monkeypatch):
    vendor = "北京某翻译店有限公司"
    _write_pay(tmp_path / "付款.xlsx", [[vendor, 106, vendor]])
    folder = tmp_path / vendor
    folder.mkdir()
    _dummy_pdf(folder / "a.pdf")
    monkeypatch.setattr(
        convert,
        "parse_invoice_pdf",
        lambda p: {"kind": "专票", "seller": vendor, "total": Decimal("106.00"), "tax": Decimal("6.00")},
    )
    result = _run(tmp_path, "付款")
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    codes = [ws.cell(r, 20).value for r in range(4, 8)]
    kd.close()
    assert "8001" in codes


def test_sales_unmapped_employee_holds(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(app="杨利宏")], with_org=False)
    result = _run(tmp_path, "销项发票")
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1
    detail = load_workbook(result["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "职员" in reason


def test_expired_cache_not_used_when_auth_fails(tmp_path, monkeypatch):
    cache = tmp_path / "kingdee-master.json"
    cache.write_text(
        json.dumps(
            {
                "cached_at": time.time() - 3600,
                "data": {
                    "customer": [{"code": "1", "name": "缓存客户"}],
                    "employee": [],
                    "supplier": [],
                    "department": [],
                },
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv("KINGDEE_MASTER_CACHE", str(cache))
    monkeypatch.setenv("KINGDEE_MASTER_CACHE_TTL_SECONDS", "900")
    monkeypatch.setattr(
        kingdee_api,
        "load_local",
        lambda: {"client_id": "x", "client_secret": "x", "app_key": "x", "app_secret": "x"},
    )
    monkeypatch.setattr(
        kingdee_api,
        "get_app_token",
        lambda creds: (_ for _ in ()).throw(RuntimeError("auth missing data errcode=1030002006")),
    )
    loaded = kingdee_api.try_load_master()
    assert loaded["ok"] is False
    assert loaded.get("source") != "cache"


def test_sales_zero_balance_books_period_debit_account(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    lookups = _lookups(
        ar_balance=[{"customer_code": "1001", "account": "113103", "balance": "0"}],
        period_debit=[{"customer_code": "1001", "account": "113103", "period": "2026-08", "debit": "80"}],
    )
    result = _run(tmp_path, "销项发票", lookups=lookups)
    assert result["bookable_count"] == 1
    assert result["hold_count"] == 0
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    codes = []
    for r in range(4, 7):
        codes.extend(str(ws.cell(r, c).value or "") for c in range(1, 40))
    kd.close()
    assert "113103" in codes
    assert "510103" in codes


def test_sales_zero_balance_without_debit_holds_kingdee_ar(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    lookups = _lookups(ar_balance=[], period_debit=[])
    result = _run(tmp_path, "销项发票", lookups=lookups)
    assert result["bookable_count"] == 0
    assert result["hold_count"] == 1
    detail = load_workbook(result["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "是否新建" in reason
    assert "业务线" not in reason


def test_period_debit_fetch_is_called_when_injected_missing(tmp_path):
    called = {}

    def fetch(box, cus_code, day, accounts):
        called["ok"] = True
        box.ar_balance[(cus_code, "113103")] = Decimal("80")
        box.ar_balance[(cus_code, "113102")] = Decimal("10")
        box.period_debit[(cus_code, "113103", "2026-08")] = Decimal("80")
        box.period_debit[(cus_code, "113102", "2026-08")] = Decimal("10")

    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    result = convert.run_dir(
        tmp_path,
        "销项发票",
        "2026-08-27",
        _master(),
        _lookups(ar_balance=[]),
        period_fetch=fetch,
        out_dir=tmp_path,
    )
    assert called.get("ok") is not True
    assert result["bookable_count"] == 0
    detail = load_workbook(result["detail_path"])
    reason = str(detail.active.cell(2, 2).value or "")
    detail.close()
    assert "是否新建" in reason


def test_start_voucher_no_shifts_payment_batches(tmp_path, monkeypatch):
    _write_pay(tmp_path / "付款.xlsx", [["北京某翻译店", 100, "北京某翻译店"], ["无档店", 200, "无档店"]])
    a = tmp_path / "北京某翻译店"
    b = tmp_path / "无档店"
    a.mkdir()
    b.mkdir()
    _dummy_pdf(a / "a.pdf")
    _dummy_pdf(b / "b.pdf")
    monkeypatch.setattr(
        convert,
        "parse_invoice_pdf",
        lambda p: {"kind": "普票", "seller": p.parent.name, "total": Decimal("100.00") if p.parent.name == "北京某翻译店" else Decimal("200.00"), "tax": None},
    )
    result = _run(tmp_path, "付款", start_voucher_no=20)
    nums = [n for n in _voucher_nums(result["kingdee_path"], 8) if n]
    assert set(nums) == {20, 21}


def test_default_desktop_dir_helper(tmp_path):
    desktop = tmp_path / "Desktop"
    desktop.mkdir()
    from datetime import date

    got = convert.default_desktop_dir("金蝶入账", today=date(2026, 9, 2), home=tmp_path)
    assert got == desktop / "金蝶入账_20260902"


def test_ar_accounts_include_113105():
    assert "113105" in convert.load_ar_accounts()


def test_sales_convert_does_not_pick_by_period_debit():
    text = (SCRIPTS / "convert.py").read_text(encoding="utf-8")
    start = text.index("def resolve_sales_party")
    end = text.index("def parse_invoice_text")
    body = text[start:end]
    assert "pick_assist_account" in body
    assert "resolve_sales_party" in body
    assert "resolve_ar" not in body
    assert "period_debit" not in body


def test_assist_export_filters_match_probe():
    assist = _load("kingdee_posting_assist_xlsx", SCRIPTS / "assist_xlsx.py")
    assert assist.EXPORT_FILTERS["assist_type"] == "客户"
    assert assist.EXPORT_FILTERS["account"] == "1131"
    assert assist.EXPORT_FILTERS["period"] == "本年"
    assert assist.EXPORT_FILTERS["hide_zero_balance"] is False
    assert "gl_rpt_assistbalance" in assist.ASSIST_FORM


def test_kingdee_api_has_no_assist_balance_openapi_path():
    text = (SCRIPTS / "kingdee_api.py").read_text(encoding="utf-8")
    assert "gl_rpt_assistbalance" not in text
    assert "/jdy/v2/fi/voucher" in text


def test_sales_reads_assist_xlsx_from_input_dir(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    _write_assist(
        tmp_path / "核算项目余额表_客户_1131_本年.xlsx",
        [
            {
                "period": "202607",
                "customer_code": "1001",
                "customer_name": "甲科技有限公司",
                "account": "113105",
                "ending_debit": "1",
                "ytd_debit": "1",
            }
        ],
    )
    result = _run(tmp_path, "销项发票", lookups=_lookups(include_assist=False))
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    accounts = [ws.cell(r, 7).value for r in range(4, 7)]
    kd.close()
    assert "113105" in accounts
    assert "510105" in accounts


def test_sales_police_books_0386_even_if_heading_absent_from_assist(tmp_path):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales(name="抚顺市公安局")], with_org=False)
    result = _run(
        tmp_path,
        "销项发票",
        lookups=_lookups(
            assist_rows=[
                {
                    "period": "202607",
                    "customer_code": "0386",
                    "customer_name": "公安部",
                    "account": "113103",
                    "ending_debit": "1",
                    "ytd_debit": "1",
                }
            ]
        ),
    )
    assert result["bookable_count"] == 1
    kd = load_workbook(result["kingdee_path"])
    ws = kd[convert.KINGDEE_SHEET]
    names = [ws.cell(r, 19).value for r in range(4, 7)]
    codes = [ws.cell(r, 18).value for r in range(4, 7)]
    kd.close()
    assert "公安部" in names
    assert "0386" in [str(c) for c in codes]


def test_sales_default_voucher_no_follows_month_max(tmp_path, monkeypatch):
    _write_sales(tmp_path / "发票.xlsx", [_ok_sales()], with_org=False)
    master = tmp_path / "master.json"
    lookups = tmp_path / "lookups.json"
    master.write_text(json.dumps(_master()), encoding="utf-8")
    lookups.write_text(json.dumps(_lookups()), encoding="utf-8")

    def fake_request(method, url, used, path, params=None, extra_headers=None, timeout=30):
        if path == "/jdy/v2/fi/voucher":
            return _FakeResp(
                {
                    "errcode": 0,
                    "data": {
                        "rows": [
                            {"id": "a", "number": 25},
                            {"id": "b", "number": 18},
                        ],
                        "count": 2,
                    },
                }
            )
        raise AssertionError(path)

    monkeypatch.setattr(convert.kingdee_api, "load_local", lambda: {
        "client_id": "x",
        "client_secret": "x" * 32,
        "app_key": "k",
        "app_secret": "s" * 40,
    })
    monkeypatch.setattr(convert.kingdee_api, "get_app_token", lambda creds: ("tok", "https://tf.jdy.com"))
    monkeypatch.setattr(convert.kingdee_api, "_request", fake_request)
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
                "--out-dir",
                str(tmp_path),
            ]
        )
        == 0
    )
    nums = [n for n in _voucher_nums(tmp_path / "凭证引入_结果.xlsx", 9) if n]
    assert min(nums) == 26
    assert 1 not in nums
