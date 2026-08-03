# -*- coding: utf-8 -*-
"""真实误判回归：本币、累计核销、手续费、分位尾差和多 SO 安全边界。"""
import datetime as dt
import zipfile

import openpyxl

import classify_hexiao as C
import xlsx_patch


def _payment(amount=100.0, orders=None, writeoffs=None, **extra):
    value = {
        "ar": "AR_TEST",
        "hexiao_date": dt.date(2026, 7, 27),
        "arrival_date": dt.date(2026, 7, 24),
        "amount_orig": amount,
        "amount_local": amount,
        "fee": 0.0,
        "currency": "人民币CNY",
        "huikuan_type": "",
        "status": "核销成功",
        "customer": "测试客户",
        "orders": orders if orders is not None else [{"so": "SO1", "deliver": amount}],
        "writeoffs": writeoffs or {},
        "writeoffs_local": {},
        "cumulative_writeoffs": {},
        "cumulative_writeoffs_local": {},
        "sod_lines": {},
    }
    value.update(extra)
    return value


def _ledger(rows):
    so_index, sod_index = {}, {}
    for row, snap in rows.items():
        if snap.get("so"):
            so_index.setdefault(snap["so"], []).append(row)
        if snap.get("sod"):
            sod_index.setdefault(snap["sod"], []).append(row)
    return C.LedgerIndex(
        synthetic={"so": so_index, "sod": sod_index, "rows": rows}
    )


def test_foreign_currency_uses_zhiyun_local_amount_without_rate():
    """SO26030487 类：智云已有原币/本币金额，不应再报缺汇率。"""
    p = _payment(
        amount=2500.0,
        amount_local=17979.75,
        currency="美元USD",
        orders=[{"so": "SO1", "deliver": 2500.0, "rate": None}],
        writeoffs={"SO1": 2500.0},
        writeoffs_local={"SO1": 17979.75},
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 2500.0}]},
    )
    rec = C.expand_payment(p, {})[0]
    assert rec.get("forced_code") is None
    assert rec["amount_local"] == 17979.75
    assert rec["deliver_local"] == 17979.75

    result = C.classify_one(
        rec,
        _ledger({1: {"so": "SO1", "sod": "SOD1", "yingshou": 2500.0}}),
        {},
        0.0,
        2026,
    )
    assert result["bucket"] == "auto"
    assert result["code"] != "E6"
    assert result["five_cols"]["回款明细"] == 17979.75


def test_foreign_full_writeoff_without_detail_uses_order_rate_as_cny():
    """无逐SO核销明细时，USD交付额不能把原币数直接写进人民币盈亏表。"""
    p = _payment(
        amount=2500.0,
        amount_local=17979.75,
        currency="美元USD",
        orders=[{"so": "SO1", "deliver": 2500.0, "rate": 7.1919, "currency": "美元USD"}],
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 2500.0}]},
    )
    rec = C.expand_payment(p, {})[0]
    assert rec["amount_orig"] == 2500.0
    assert rec["amount_local"] == 17979.75
    assert rec["deliver_local"] == 17979.75


def test_unique_sod_uses_cumulative_writeoff_to_recognize_final_settlement():
    """SO26050128 类：本次只回尾款，但历史+本次已达到交付额，应结清而非再拆行。"""
    p = _payment(
        amount=1030.47,
        orders=[{"so": "SO1", "deliver": 4083.60}],
        writeoffs={"SO1": 1030.47},
        cumulative_writeoffs={"SO1": 4083.60},
        cumulative_writeoffs_local={"SO1": 4083.60},
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 4083.60}]},
    )
    rec = C.expand_payment(p, {})[0]
    assert rec["amount_orig"] == 1030.47
    assert rec["cumulative_received_local"] == 4083.60

    result = C.classify_one(
        rec,
        _ledger({1: {"so": "SO1", "sod": "SOD1", "yingshou": 4083.60}}),
        {},
        0.0,
        2026,
    )
    assert result["bucket"] == "auto"
    assert "row_operation" not in result
    assert result["five_cols"]["回款明细"] == 1030.47
    assert result["five_cols"]["计提"] == 4083.60
    assert result["five_cols"]["是否结账"] == "是"


def test_split_rows_use_aggregate_receivable_baseline_for_business_difference():
    """历史拆成两行后，B0 是两行应收合计；合计等于 D 时不得再造业务差异。"""
    p = _payment(
        amount=30.0,
        orders=[{"so": "SO1", "deliver": 100.0}],
        writeoffs={"SO1": 30.0},
        cumulative_writeoffs={"SO1": 100.0},
        cumulative_writeoffs_local={"SO1": 100.0},
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 100.0}]},
    )
    rec = C.expand_payment(p, {})[0]
    result = C.classify_one(
        rec,
        _ledger({
            1: {
                "so": "SO1", "sod": "SOD1", "yingshou": 30.0,
                "huikuan": None, "jiezhang": "否",
            },
            2: {
                "so": "SO1", "sod": "SOD1", "yingshou": 70.0,
                "huikuan": 70.0, "jiezhang": "是",
            },
        }),
        {},
        0.0,
        2026,
    )
    assert result["bucket"] == "auto"
    assert result["five_cols"]["计提"] == 100.0
    assert "derived_cols" not in result


def test_whole_payment_cross_month_uses_writeoff_date_and_advance_offset():
    """回款类型不再例外：整笔回款跨月也必须填核销日 + 冲预收。"""
    p = _payment(
        amount=100.0,
        huikuan_type="整笔回款",
        arrival_date=dt.date(2026, 4, 27),
        hexiao_date=dt.date(2026, 7, 27),
        orders=[{"so": "SO1", "deliver": 100.0}],
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 100.0}]},
    )
    rec = C.expand_payment(p, {})[0]
    result = C.classify_one(
        rec,
        _ledger({1: {"so": "SO1", "sod": "SOD1", "yingshou": 100.0}}),
        {},
        0.0,
        2026,
    )
    assert result["bucket"] == "auto"
    assert result["five_cols"]["收款时间"] == "2026-07-27"
    assert result["five_cols"]["收款方式"] == "冲预收"


def test_whole_payment_same_month_uses_arrival_date_and_hui():
    """回款类型不再例外：整笔回款同月按到账日 + 汇。"""
    p = _payment(
        amount=100.0,
        huikuan_type="整笔回款",
        arrival_date=dt.date(2026, 7, 24),
        hexiao_date=dt.date(2026, 7, 27),
        orders=[{"so": "SO1", "deliver": 100.0}],
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 100.0}]},
    )
    rec = C.expand_payment(p, {})[0]
    result = C.classify_one(
        rec,
        _ledger({1: {"so": "SO1", "sod": "SOD1", "yingshou": 100.0}}),
        {},
        0.0,
        2026,
    )
    assert result["bucket"] == "auto"
    assert result["five_cols"]["收款时间"] == "2026-07-24"
    assert result["five_cols"]["收款方式"] == "汇"


def test_fee_is_ignored_when_explicit_so_writeoffs_exist():
    """手续费不参与判定；逐 SO 本次核销金额原样进入业务字段。"""
    p = _payment(
        amount=300.0,
        amount_local=None,
        fee=99.0,
        currency="美元USD",
        orders=[{"so": "SO1", "deliver": 100.0}, {"so": "SO2", "deliver": 200.0}],
        writeoffs={"SO1": 100.0, "SO2": 200.0},
        sod_lines={
            "SO1": [{"sod": "SOD1", "deliver": 100.0}],
            "SO2": [{"sod": "SOD2", "deliver": 200.0}],
        },
    )
    recs = C.expand_payment(p, {})
    assert {r["so"] for r in recs} == {"SO1", "SO2"}
    assert all(r.get("forced_code") is None for r in recs)
    assert round(sum(r["amount_orig"] for r in recs), 2) == 300.0
    assert all(r["fee"] == 0.0 for r in recs)
    assert all("手续费忽略" in r["match_basis"] for r in recs)


def test_fee_never_reduces_or_allocates_writeoff_details():
    """即使逐 SO 合计看起来包含手续费，也不得扣减或重新分配。"""
    p = _payment(
        amount=300.0,
        fee=1.0,
        orders=[{"so": "SO1", "deliver": 100.0}, {"so": "SO2", "deliver": 201.0}],
        writeoffs={"SO1": 100.0, "SO2": 201.0},
        cumulative_writeoffs={"SO1": 100.0, "SO2": 201.0},
        sod_lines={
            "SO1": [{"sod": "SOD1", "deliver": 100.0}],
            "SO2": [{"sod": "SOD2", "deliver": 201.0}],
        },
    )
    recs = C.expand_payment(p, {})
    assert {r["so"] for r in recs} == {"SO1", "SO2"}
    assert all(r.get("forced_code") is None for r in recs)
    assert round(sum(r["amount_orig"] for r in recs), 2) == 301.0
    assert round(sum(r["deliver_local"] for r in recs), 2) == 301.0


def test_fee_does_not_block_or_rewrite_historical_cumulative_writeoff():
    p = _payment(
        amount=30.0,
        amount_local=None,
        fee=1.0,
        currency="美元USD",
        orders=[{"so": "SO1", "deliver": 80.0, "currency": "人民币CNY"}],
        writeoffs={"SO1": 30.0},
        cumulative_writeoffs={"SO1": 80.0},
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 80.0, "currency": "人民币CNY"}]},
    )
    rec = C.expand_payment(p, {})[0]
    assert rec.get("forced_code") is None
    assert rec["amount_local"] == 30.0
    assert rec["cumulative_received_local"] == 80.0


def test_no_writeoff_details_means_full_delivery_without_fee_allocation():
    """智云没有逐 SO 子明细时按全额核销，手续费仍不参与金额重写。"""
    p = _payment(
        amount=300.0,
        fee=1.0,
        orders=[{"so": "SO1", "deliver": 100.0}, {"so": "SO2", "deliver": 201.0}],
        sod_lines={
            "SO1": [{"sod": "SOD1", "deliver": 100.0}],
            "SO2": [{"sod": "SOD2", "deliver": 201.0}],
        },
    )
    recs = C.expand_payment(p, {})
    assert {r["so"] for r in recs} == {"SO1", "SO2"}
    assert all(r.get("forced_code") is None for r in recs)
    assert round(sum(r["amount_orig"] for r in recs), 2) == 301.0
    assert all("手续费忽略" in r["match_basis"] for r in recs)


def test_writeoff_without_local_amount_or_rate_uses_writeoff_directly():
    """有本次核销金额就直接跑，不因父回款外币或缺汇率报 E6。"""
    p = _payment(
        amount=90.0,
        amount_local=None,
        fee=1.0,
        currency="美元USD",
        orders=[{"so": "SO1", "deliver": 100.0, "currency": "人民币CNY", "rate": None}],
        writeoffs={"SO1": 90.0},
        cumulative_writeoffs={"SO1": 90.0},
        sod_lines={"SO1": [{"sod": "SOD1", "deliver": 100.0, "currency": "人民币CNY"}]},
    )
    rec = C.expand_payment(p, {})[0]
    assert rec.get("forced_code") is None
    assert rec["amount_local"] == 90.0
    assert rec["deliver_local"] == 100.0
    assert rec["cumulative_received_local"] == 90.0


def test_one_cent_multi_so_tail_is_rounding_not_partial_payment():
    """SO26030090/SO26040561 类：父回款与多 SO 合计差 0.01，不制造假部分回款。"""
    p = _payment(
        amount=299.99,
        orders=[{"so": "SO1", "deliver": 100.0}, {"so": "SO2", "deliver": 200.0}],
        sod_lines={
            "SO1": [{"sod": "SOD1", "deliver": 100.0}],
            "SO2": [{"sod": "SOD2", "deliver": 200.0}],
        },
    )
    recs = C.expand_payment(p, {})
    assert len(recs) == 2
    assert all(r.get("forced_code") is None for r in recs)
    assert round(sum(r["amount_orig"] for r in recs), 2) == 300.0


def test_no_writeoff_details_follow_zhiyun_full_settlement_semantics():
    """没有逐 SO 子明细即全额核销，不再拿父到账额制造假部分回款。"""
    p = _payment(
        amount=250.0,
        orders=[{"so": "SO1", "deliver": 100.0}, {"so": "SO2", "deliver": 200.0}],
        sod_lines={
            "SO1": [{"sod": "SOD1", "deliver": 100.0}],
            "SO2": [{"sod": "SOD2", "deliver": 200.0}],
        },
    )
    recs = C.expand_payment(p, {})
    assert {r["so"] for r in recs} == {"SO1", "SO2"}
    assert all(r.get("forced_code") is None for r in recs)
    assert round(sum(r["amount_local"] for r in recs), 2) == 300.0


def test_xlsx_patch_forces_excel_formula_recalculation(tmp_path):
    """写表后要求 Excel 重算，避免公式存在但缓存值仍是旧值。"""
    src = tmp_path / "source.xlsx"
    out = tmp_path / "output.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    ws["A1"] = 1
    ws["B1"] = "=A1*2"
    wb.save(src)

    xlsx_patch.patch_cells(src, out, "明细", [(1, 1, 3)])
    with zipfile.ZipFile(out) as zf:
        workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")
    assert 'calcMode="auto"' in workbook_xml
    assert 'fullCalcOnLoad="1"' in workbook_xml
    assert 'forceFullCalc="1"' in workbook_xml


def test_xlsx_patch_cleanly_invalidates_stale_calc_chain(tmp_path):
    """业务格或行坐标变化后，旧计算链必须连同关系声明一起移除。"""
    src = tmp_path / "source_with_chain.xlsx"
    rebuilt = tmp_path / "source_with_chain_rebuilt.xlsx"
    out = tmp_path / "output.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    ws["A1"] = 1
    ws["B1"] = "=A1*2"
    wb.save(src)

    with zipfile.ZipFile(src) as zin:
        payload = {name: zin.read(name) for name in zin.namelist()}
    rels_name = "xl/_rels/workbook.xml.rels"
    rels = payload[rels_name].decode("utf-8")
    rels = rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdCalc" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" '
        'Target="calcChain.xml"/></Relationships>',
    )
    payload[rels_name] = rels.encode("utf-8")
    content_types = payload["[Content_Types].xml"].decode("utf-8")
    content_types = content_types.replace(
        "</Types>",
        '<Override PartName="/xl/calcChain.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
        "</Types>",
    )
    payload["[Content_Types].xml"] = content_types.encode("utf-8")
    payload["xl/calcChain.xml"] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<c r="B1" i="1"/></calcChain>'
    )
    with zipfile.ZipFile(rebuilt, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in payload.items():
            zout.writestr(name, data)
    rebuilt.replace(src)

    xlsx_patch.patch_cells(src, out, "明细", [(1, 1, 3)])

    with zipfile.ZipFile(out) as zf:
        names = set(zf.namelist())
        rels = zf.read(rels_name).decode("utf-8")
        content_types = zf.read("[Content_Types].xml").decode("utf-8")
    assert "xl/calcChain.xml" not in names
    assert "calcChain" not in rels
    assert "calcChain" not in content_types
    assert xlsx_patch.parts_diff(src, out) == []
