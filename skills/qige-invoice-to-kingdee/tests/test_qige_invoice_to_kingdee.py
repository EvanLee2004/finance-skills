#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""琪哥发票入金蝶 · 合成回归 + 本机 30 张金标（有文件才跑）。"""
from __future__ import annotations

import json
import sys
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

import convert  # noqa: E402

GOLD_INVOICE = Path(
    "/Users/evanlee/Documents/甲骨易实习/项目/长期项目/财务部skills"
    "/技能/琪哥发票入金蝶/原始素材/20260814_改样发票/发票明昊.xlsx"
)
GOLD_MASTER = Path(
    "/Users/evanlee/Documents/甲骨易实习/项目/长期项目/自动化记账（金蝶）"
    "/原始素材/实操批次/20260803_金蝶补辅助核算_七月前半/调研/金蝶基础资料_master.json"
)


def _master():
    return {
        "employee": [
            {"code": "103", "name": "于占国"},
            {"code": "011", "name": "陈霞"},
            {"code": "036", "name": "梁玲玲"},
        ],
        "department": [
            {"code": "15", "name": "本地化事业部"},
            {"code": "0302", "name": "营销二部"},
        ],
        "customer": [
            {"code": "1001", "name": "甲科技有限公司"},
            {"code": "1002", "name": "乙事务所"},
            {"code": "2460", "name": "乙中心"},
            {"code": "3843", "name": "丙科技（深圳）有限公司"},
        ],
    }


def _write_invoice(path: Path, rows, org=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "发票"
    ws.append(
        [
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
    org_ws = wb.create_sheet("组织架构")
    org_ws.append(["姓名", "部门编码"])
    for pair in org or [("于占国", "15"), ("陈霞", "15"), ("梁玲玲", "0302"), ("不存在职员甲", "15")]:
        org_ws.append(list(pair))
    wb.save(path)
    wb.close()


def _write_master(path: Path, data=None):
    path.write_text(json.dumps(data or _master(), ensure_ascii=False), encoding="utf-8")


def _run(tmp_path: Path, rows, org=None, master=None, booking="2026-08-14"):
    inv = tmp_path / "发票.xlsx"
    mas = tmp_path / "master.json"
    out = tmp_path / "out"
    out.mkdir()
    _write_invoice(inv, rows, org=org)
    _write_master(mas, master)
    return convert.convert(
        invoice_path=inv,
        master_path=mas,
        out_dir=out,
        booking_date=booking,
    )


def _ok_row(
    name="甲科技有限公司",
    app="于占国",
    tot=1060,
    amt=1000,
    tax=60,
    ar="113103",
    rev="510103",
    typ="专票",
    inv="1",
    day="2026-08-01",
):
    return [day, typ, inv, name, tot, amt, tax, app, None, ar, rev]


def test_skip_empty_unit_name(tmp_path):
    rows = [_ok_row(), [None, None, None, None, None, None, None, None, None, None, None]]
    result = _run(tmp_path, rows)
    assert result.source_count == 1
    assert result.bookable_count == 1
    assert result.hold_count == 0


def test_hold_missing_account(tmp_path):
    rows = [_ok_row(ar=None, rev="510103")]
    result = _run(tmp_path, rows)
    assert result.bookable_count == 0
    assert result.hold_count == 1
    assert "科目" in result.holds[0]["原因"]


def test_hold_missing_customer(tmp_path):
    rows = [_ok_row(name="不存在客户甲有限公司")]
    result = _run(tmp_path, rows)
    assert result.bookable_count == 0
    assert "客户" in result.holds[0]["原因"]


def test_hold_customer_suffix_mismatch(tmp_path):
    rows = [_ok_row(name="乙中心有限公司")]
    result = _run(tmp_path, rows)
    assert result.bookable_count == 0
    assert "客户" in result.holds[0]["原因"]


def test_hold_missing_employee(tmp_path):
    rows = [_ok_row(app="不存在职员甲")]
    result = _run(tmp_path, rows)
    assert result.bookable_count == 0
    assert "职员" in result.holds[0]["原因"]


def test_normalize_brackets_and_spaces_match(tmp_path):
    rows = [
        _ok_row(name="丙科技(深圳)有限公司", inv="a"),
        _ok_row(name="甲科技有限公司 ", inv="b"),
    ]
    result = _run(tmp_path, rows)
    assert result.bookable_count == 2
    codes = {r["客户编码"] for r in result.bookable}
    assert codes == {"3843", "1001"}


def test_three_lines_balance_and_tax_no_aux(tmp_path):
    result = _run(tmp_path, [_ok_row()])
    kd = load_workbook(result.kingdee_path)
    ws = kd["sheet1（名称勿改）"]
    lines = []
    for row in ws.iter_rows(min_row=4, max_row=6, max_col=25, values_only=True):
        lines.append(row)
    assert len(lines) == 3
    debit = sum(Decimal(str(r[15] or 0)) for r in lines)
    credit = sum(Decimal(str(r[16] or 0)) for r in lines)
    assert debit == credit == Decimal("1060.00")
    assert lines[0][6] == "113103"
    assert lines[1][6] == "510103"
    assert lines[2][6] == "21710105"
    assert lines[2][17] in (None, "")
    assert lines[2][21] in (None, "")
    assert lines[2][23] in (None, "")
    assert lines[0][17] == "1001"
    assert str(lines[0][21]) == "15"
    assert str(lines[0][23]) == "103"
    assert lines[0][5] == "专票：甲科技有限公司"
    assert str(lines[0][0]) == "2026-08-14"
    assert lines[0][1] == "记"
    kd.close()


def test_pack_five_per_voucher(tmp_path):
    rows = [_ok_row(inv=str(i), tot=1060, amt=1000, tax=60) for i in range(12)]
    result = _run(tmp_path, rows)
    assert result.bookable_count == 12
    kd = load_workbook(result.kingdee_path)
    ws = kd["sheet1（名称勿改）"]
    nums = []
    for row in ws.iter_rows(min_row=4, max_row=3 + 12 * 3, max_col=3, values_only=True):
        nums.append(row[2])
    assert nums.count(1) == 15
    assert nums.count(2) == 15
    assert nums.count(3) == 6
    kd.close()


def test_booking_date_not_invoice_date(tmp_path):
    result = _run(tmp_path, [_ok_row(day="2026-08-01")], booking="2026-08-14")
    kd = load_workbook(result.kingdee_path)
    ws = kd["sheet1（名称勿改）"]
    assert str(ws.cell(4, 1).value) == "2026-08-14"
    kd.close()


def test_red_invoice_negative_and_summary(tmp_path):
    rows = [_ok_row(typ="普票", tot=-1060, amt=-1000, tax=-60)]
    result = _run(tmp_path, rows)
    kd = load_workbook(result.kingdee_path)
    ws = kd["sheet1（名称勿改）"]
    assert ws.cell(4, 6).value == "红字普票：甲科技有限公司"
    assert Decimal(str(ws.cell(4, 16).value)) == Decimal("-1060.00")
    kd.close()


def test_detail_covers_every_source_row(tmp_path):
    rows = [
        _ok_row(inv="ok"),
        _ok_row(name="不存在客户甲有限公司", inv="hold"),
    ]
    result = _run(tmp_path, rows)
    det = load_workbook(result.detail_path)
    ws = det.active
    statuses = [ws.cell(i, 1).value for i in range(2, ws.max_row + 1)]
    assert len(statuses) == 2
    assert set(statuses) == {"可入账", "待确认"}
    det.close()


def test_missing_master_stops(tmp_path):
    inv = tmp_path / "发票.xlsx"
    _write_invoice(inv, [_ok_row()])
    with pytest.raises(SystemExit):
        convert.convert(
            invoice_path=inv,
            master_path=tmp_path / "nope.json",
            out_dir=tmp_path / "out",
            booking_date="2026-08-14",
        )


@pytest.mark.skipif(not GOLD_INVOICE.is_file() or not GOLD_MASTER.is_file(), reason="本机金标文件不在")
def test_gold_minghao_30(tmp_path):
    out = tmp_path / "gold"
    out.mkdir()
    result = convert.convert(
        invoice_path=GOLD_INVOICE,
        master_path=GOLD_MASTER,
        out_dir=out,
        booking_date="2026-08-14",
    )
    assert result.source_count == 30
    assert result.hold_count == 8
    assert result.bookable_count == 22
    assert all(h.get("原因") for h in result.holds)
    kd = load_workbook(result.kingdee_path)
    ws = kd["sheet1（名称勿改）"]
    n_lines = 0
    debit = Decimal("0")
    credit = Decimal("0")
    for row in ws.iter_rows(min_row=4, max_col=17, values_only=True):
        if not row[6]:
            continue
        n_lines += 1
        debit += Decimal(str(row[15] or 0))
        credit += Decimal(str(row[16] or 0))
    assert n_lines == 66
    assert debit == credit
    kd.close()
