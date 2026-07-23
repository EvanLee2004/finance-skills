# -*- coding: utf-8 -*-
"""
步骤6 判定 v2（单入口 · SOD 级）：展开、E 码、覆盖率硬校验、SOD 子集与整段对齐。
"""
import datetime as dt

import pytest

import classify_hexiao as C
from conftest import GOLD_DIR


# ── 小工具 ────────────────────────────────────────────────
def _pay(ar="AR_T", amount=100.0, orders=None, writeoffs=None, **kw):
    p = {
        "ar": ar,
        "hexiao_date": dt.date(2026, 7, 22),
        "arrival_date": dt.date(2026, 7, 21),
        "amount_orig": amount,
        "amount_local": amount,
        "fee": 0.0,
        "currency": "人民币CNY",
        "huikuan_type": "",
        "status": "手动核销",
        "customer": "测试客户甲",
        "orders": orders if orders is not None else [{"so": "SO26010001", "deliver": amount}],
        "writeoffs": writeoffs or {},
        "sod_lines": {},
    }
    p.update(kw)
    return p


def _led(rows: dict):
    """rows = {行号: {"so","sod","yingshou", …}} → 合成索引。"""
    so, sod = {}, {}
    for r, snap in rows.items():
        if snap.get("so"):
            so.setdefault(snap["so"], []).append(r)
        if snap.get("sod"):
            sod.setdefault(snap["sod"], []).append(r)
    return C.LedgerIndex(synthetic={"so": so, "sod": sod, "rows": rows})


# ══════════════════════════════════════════════════════════
# 一、subset_sum_unique
# ══════════════════════════════════════════════════════════
def test_subset_unique_hit():
    lines = [{"sod": "A", "deliver": 10.0}, {"sod": "B", "deliver": 25.5}, {"sod": "C", "deliver": 4.5}]
    got = C.subset_sum_unique(lines, 30.0)  # 25.5+4.5
    assert sorted(x["sod"] for x in got) == ["B", "C"]


def test_subset_multiple_solutions_returns_none():
    """两组都能凑出 → 不许随便挑一个。"""
    lines = [{"sod": "A", "deliver": 10.0}, {"sod": "B", "deliver": 10.0}, {"sod": "C", "deliver": 20.0}]
    assert C.subset_sum_unique(lines, 20.0) is None


def test_subset_no_solution_returns_none():
    lines = [{"sod": "A", "deliver": 10.0}, {"sod": "B", "deliver": 25.0}]
    assert C.subset_sum_unique(lines, 12.0) is None


def test_subset_too_many_lines_gives_up():
    lines = [{"sod": f"S{i}", "deliver": float(i + 1)} for i in range(40)]
    assert C.subset_sum_unique(lines, 3.0) is None


def test_subset_uses_cents_not_float():
    """0.1+0.2 这种浮点坑不许把命中判成不命中。"""
    lines = [{"sod": "A", "deliver": 0.1}, {"sod": "B", "deliver": 0.2}]
    assert C.subset_sum_unique(lines, 0.3) is not None


# ══════════════════════════════════════════════════════════
# 二、展开：一笔到账 → SOD 级 records
# ══════════════════════════════════════════════════════════
def test_no_writeoff_means_full_settle():
    """核销明细 0 行 = 全额核销（明妹原话：没有这张就说明到账=交付）。"""
    p = _pay(amount=5800.0, orders=[
        {"so": "SO1", "deliver": 1450.0}, {"so": "SO2", "deliver": 1740.0}, {"so": "SO3", "deliver": 2610.0},
    ])
    recs = C.expand_payment(p, {})
    assert len(recs) == 3
    assert sum(r["amount_orig"] for r in recs) == 5800.0
    assert all("全额核销" in r["match_basis"] for r in recs)


def test_writeoff_overrides_deliver():
    """有核销明细就以它为准（部分核销场景）。"""
    p = _pay(amount=1000.0,
             orders=[{"so": "SO1", "deliver": 900.0}, {"so": "SO2", "deliver": 500.0}],
             writeoffs={"SO1": 700.0, "SO2": 300.0})
    recs = C.expand_payment(p, {})
    assert sum(r["amount_orig"] for r in recs) == 1000.0
    assert all("核销明细" in r["match_basis"] for r in recs)


def test_sod_expansion_full():
    """一个 SO 拆 N 个 SOD，合计=核销额 → N 行全出（旧版误判成 E8 歧义的那类）。"""
    p = _pay(amount=300.0, orders=[{"so": "SO1", "deliver": 300.0}])
    p["sod_lines"] = {"SO1": [
        {"sod": "SOD3", "deliver": 100.0}, {"sod": "SOD2", "deliver": 120.0}, {"sod": "SOD1", "deliver": 80.0},
    ]}
    recs = C.expand_payment(p, {})
    assert len(recs) == 3
    assert {r["sod"] for r in recs} == {"SOD1", "SOD2", "SOD3"}
    assert sum(r["amount_orig"] for r in recs) == 300.0


def test_sod_expansion_subset():
    p = _pay(amount=180.0, orders=[{"so": "SO1", "deliver": 300.0}], writeoffs={"SO1": 180.0})
    p["sod_lines"] = {"SO1": [
        {"sod": "SOD1", "deliver": 100.0}, {"sod": "SOD2", "deliver": 120.0}, {"sod": "SOD3", "deliver": 80.0},
    ]}
    recs = C.expand_payment(p, {})
    assert {r["sod"] for r in recs} == {"SOD1", "SOD3"}


def test_sod_ambiguous_holds_e5():
    p = _pay(amount=100.0, orders=[{"so": "SO1", "deliver": 300.0}], writeoffs={"SO1": 100.0})
    p["sod_lines"] = {"SO1": [
        {"sod": "SOD1", "deliver": 100.0}, {"sod": "SOD2", "deliver": 100.0}, {"sod": "SOD3", "deliver": 100.0},
    ]}
    recs = C.expand_payment(p, {})
    assert len(recs) == 1 and recs[0]["forced_code"] == "E5"
    assert "SOD1" in recs[0]["forced_reason"]  # 候选要摆出来给她挑


def test_no_sod_falls_back_to_so():
    """订单明细查不到 SOD → 退化按 SO 匹配，仍然可判，不是丢单。"""
    p = _pay(amount=50.0, orders=[{"so": "SO1", "deliver": 50.0}])
    recs = C.expand_payment(p, {})
    assert len(recs) == 1 and recs[0]["sod"] == "" and recs[0]["amount_orig"] == 50.0


def test_fenbi_always_hold_e1():
    p = _pay(huikuan_type="分笔回款")
    recs = C.expand_payment(p, {})
    assert len(recs) == 1 and recs[0]["forced_code"] == "E1"


def test_fee_holds_whole_payment():
    p = _pay(fee=1.62)
    recs = C.expand_payment(p, {})
    assert recs[0]["forced_code"] == "E_FEE"


def test_no_orders_is_e7_not_silent_drop():
    p = _pay(orders=[])
    recs = C.expand_payment(p, {})
    assert len(recs) == 1 and recs[0]["forced_code"] == "E7"


def test_writeoff_over_arrival_is_e4():
    p = _pay(amount=100.0, orders=[{"so": "SO1", "deliver": 500.0}], writeoffs={"SO1": 500.0})
    recs = C.expand_payment(p, {})
    assert recs[0]["forced_code"] == "E4"


def test_deliver_over_arrival_is_e5_split_row():
    """AR26070105 那类：整单关联但钱没到齐 → 提示拆行，不是超额核销。"""
    p = _pay(amount=8729.35, orders=[{"so": "SO1", "deliver": 8946.89}])
    recs = C.expand_payment(p, {})
    assert recs[0]["forced_code"] == "E5"
    assert "插一行" in recs[0]["forced_reason"]


def test_fx_missing_rate_e6():
    p = _pay(currency="美元USD", amount_local=None)
    recs = C.expand_payment(p, {})
    assert recs[0]["forced_code"] == "E6"


# ══════════════════════════════════════════════════════════
# 三、AR 覆盖率硬校验（2026-07-22 静默丢 3 笔的防回归）
# ══════════════════════════════════════════════════════════
def test_every_payment_produces_at_least_one_record():
    payments = [_pay(ar="A1"), _pay(ar="A2", orders=[]), _pay(ar="A3", huikuan_type="分笔回款")]
    recs = C.expand_payments(payments, {})
    assert {r["ar"] for r in recs} == {"A1", "A2", "A3"}


def test_coverage_error_raised_when_ar_lost(monkeypatch):
    """人为制造丢单 → 必须炸，绝不静默放行。"""
    def _drop(p, rates):
        return [] if p["ar"] == "A2" else [{"ar": p["ar"], "so": "SO1", "sod": "", "amount_orig": 1.0}]

    monkeypatch.setattr(C, "expand_payment", _drop)
    with pytest.raises(C.CoverageError) as e:
        C.expand_payments([_pay(ar="A1"), _pay(ar="A2")], {})
    assert "A2" in str(e.value)


# ══════════════════════════════════════════════════════════
# 四、盈亏表定位
# ══════════════════════════════════════════════════════════
def test_match_prefers_so_plus_amount():
    """主键 = SO+应收金额：回填前后都成立。"""
    led = _led({
        1: {"so": "SO1", "sod": "", "yingshou": 100.0},
        2: {"so": "SO1", "sod": "", "yingshou": 200.0},
    })
    assert led.match("SO1", "", 200.0)[0] == 2
    assert led.match("SO1", "", 100.0)[0] == 1


def test_match_falls_back_to_sod_then_so():
    led = _led({5: {"so": "SO1", "sod": "SODX", "yingshou": None}})
    assert led.match("SO1", "SODX", 99.0)[1] == "SOD"
    led2 = _led({7: {"so": "SO2", "sod": "", "yingshou": None}})
    assert led2.match("SO2", "", None)[1] == "SO"


def test_match_multi_same_amount_is_e8():
    led = _led({
        1: {"so": "SO1", "sod": "", "yingshou": 18.7},
        2: {"so": "SO1", "sod": "", "yingshou": 18.7},
    })
    row, how, cands = led.match("SO1", "", 18.7)
    assert how == "E8" and row is None and len(cands) == 2


def test_positional_alignment_resolves_equal_amounts():
    """整段逐位对齐（SOD 降序 ↔ 行号升序）能严格消掉等额歧义。"""
    led = _led({
        10: {"so": "SO1", "sod": "", "yingshou": 72.98},
        11: {"so": "SO1", "sod": "", "yingshou": 18.70},
        12: {"so": "SO1", "sod": "", "yingshou": 18.70},
        13: {"so": "SO1", "sod": "", "yingshou": 25.38},
    })
    lines = [
        {"sod": "SOD9", "deliver": 72.98}, {"sod": "SOD8", "deliver": 18.70},
        {"sod": "SOD7", "deliver": 18.70}, {"sod": "SOD6", "deliver": 25.38},
    ]
    assert led.positional_row("SO1", "SOD8", lines) == (11, "exact", None)
    assert led.positional_row("SO1", "SOD7", lines) == (12, "exact", None)


def test_positional_alignment_accepts_systematic_ratio():
    """
    智云交付额与她表里应收整段差**同一个比例**（实测 SO26040322 = 0.977433）→
    是口径差不是行错位，可以对齐；行错位凑不出同一个比值。
    """
    led = _led({
        1: {"so": "SO1", "sod": "", "yingshou": 1240.14},
        2: {"so": "SO1", "sod": "", "yingshou": 514.35},
        3: {"so": "SO1", "sod": "", "yingshou": 488.64},
        4: {"so": "SO1", "sod": "", "yingshou": 676.58},
    })
    lines = [
        {"sod": "SOD27", "deliver": 1240.14}, {"sod": "SOD26", "deliver": 514.35},
        {"sod": "SOD25", "deliver": 477.61}, {"sod": "SOD24", "deliver": 661.31},
    ]
    row, kind, ratio = led.positional_row("SO1", "SOD25", lines)
    assert (row, kind) == (3, "ratio")
    assert abs(ratio - 0.977433) < 1e-5


def test_positional_alignment_refuses_inconsistent_ratios():
    """比例各不相同 → 更像行错位，必须拒绝。"""
    led = _led({
        1: {"so": "SO1", "sod": "", "yingshou": 100.0},
        2: {"so": "SO1", "sod": "", "yingshou": 200.0},
    })
    lines = [{"sod": "SOD2", "deliver": 90.0}, {"sod": "SOD1", "deliver": 150.0}]
    assert led.positional_row("SO1", "SOD2", lines) is None


def test_positional_alignment_refuses_when_sequence_mismatch():
    """她把某个 SOD 拆成了两行 → 对不齐 → 返回 None，老实挂起。"""
    led = _led({
        10: {"so": "SO1", "sod": "", "yingshou": 202.48},
        11: {"so": "SO1", "sod": "", "yingshou": 199.78},
        12: {"so": "SO1", "sod": "", "yingshou": 450.00},
    })
    lines = [{"sod": "SOD80", "deliver": 402.26}, {"sod": "SOD79", "deliver": 450.00}]
    assert led.positional_row("SO1", "SOD80", lines) is None


# ══════════════════════════════════════════════════════════
# 五、单条判定 E 码
# ══════════════════════════════════════════════════════════
def _rec(so="SO26010001", sod="SOD26010001", amount=100.0, **kw):
    r = {
        "ar": "AR_T", "so": so, "sod": sod, "amount_orig": amount,
        "currency": "人民币CNY", "status": "手动核销", "fee": 0,
        "customer": "测试客户甲",
        "hexiao_date": dt.date(2026, 7, 22), "shoukuan_date": dt.date(2026, 7, 21),
    }
    r.update(kw)
    return r


def test_auto_happy_path():
    led = _led({3: {"so": "SO26017777", "sod": "", "yingshou": 200.0}})
    r = C.classify_one(_rec("SO26017777", "SOD26017777", 200.0), led, {}, 0.0, 2026)
    assert r["bucket"] == "auto"
    assert r["five_cols"] == {
        "计提": 200.0, "回款明细": 200.0, "是否结账": "是",
        "收款时间": "2026-07-21", "收款方式": "汇", "实收SOD": "SOD26017777",
    }
    assert "禁止用行号" in r["locate_hint"]


def test_receipt_time_is_arrival_date_same_month():
    """她填的收款时间 = 到账日（同月）。7-17 到账、7-22 核销 → 填 7-17。"""
    led = _led({1: {"so": "SO1", "sod": "", "yingshou": 10.0}})
    r = C.classify_one(
        _rec("SO1", "SOD1", 10.0, shoukuan_date=dt.date(2026, 7, 17), hexiao_date=dt.date(2026, 7, 22)),
        led, {}, 0.0, 2026,
    )
    assert r["five_cols"]["收款时间"] == "2026-07-17"
    assert r["five_cols"]["收款方式"] == "汇"


def test_cross_month_is_chongyushou():
    led = _led({1: {"so": "SO1", "sod": "", "yingshou": 10.0}})
    r = C.classify_one(
        _rec("SO1", "SOD1", 10.0, shoukuan_date=dt.date(2026, 6, 26), hexiao_date=dt.date(2026, 7, 8)),
        led, {}, 0.0, 2026,
    )
    assert r["five_cols"]["收款方式"] == "冲预收"


def test_prepaid_type_no_longer_forces_chongyushou():
    """2026-07-23 口径修正：预存回款同月核销 → 「汇」（旧版错填冲预收，15 笔全错）。"""
    led = _led({1: {"so": "SO1", "sod": "", "yingshou": 10.0}})
    r = C.classify_one(
        _rec("SO1", "SOD1", 10.0, huikuan_type="预存回款", status="核销成功"),
        led, {}, 0.0, 2026,
    )
    assert r["five_cols"]["收款方式"] == "汇"


def test_cross_year_only_after_ledger_miss():
    """2025 的单**在表里有行**就正常填；只有表里找不到才判 E3。"""
    led = _led({1: {"so": "SO25120734", "sod": "", "yingshou": 52200.0}})
    r = C.classify_one(_rec("SO25120734", "SOD25121039", 52200.0), led, {}, 0.0, 2026)
    assert r["bucket"] == "auto", r
    r2 = C.classify_one(_rec("SO25080089", "SOD25080128", 1.0), led, {}, 0.0, 2026)
    assert r2["code"] == "E3" and r2["bucket"] == "hold"


def test_missing_so_is_e2():
    led = _led({1: {"so": "SO26010000", "sod": "", "yingshou": 1.0}})
    r = C.classify_one(_rec("SO26999999", "SOD26999999", 10.0), led, {}, 0.0, 2026)
    assert r["code"] == "E2" and r["bucket"] == "hold"


def test_void_status_e7():
    r = C.classify_one(_rec(status="已作废"), C.LedgerIndex(), {}, 0.0, 2026)
    assert r["code"] == "E7" and r["bucket"] == "exception"


def test_excess_over_yingshou_e4():
    led = _led({1: {"so": "SO1", "sod": "", "yingshou": 100.0}})
    r = C.classify_one(_rec("SO1", "SOD1", 100.0), led, {}, 0.0, 2026)
    assert r["bucket"] == "auto"
    led2 = _led({1: {"so": "SO1", "sod": "SOD1", "yingshou": 50.0}})
    r2 = C.classify_one(_rec("SO1", "SOD1", 500.0), led2, {}, 0.0, 2026)
    assert r2["code"] == "E4"


def test_flow_signals():
    led = _led({1: {"so": "SO1", "sod": "", "yingshou": 100.0}})
    assert C.classify_one(_rec("SO1", "SOD1", flow_hits=0), led, {}, 0.0, 2026)["code"] == "E0"
    assert C.classify_one(_rec("SO1", "SOD1", flow_hits=3), led, {}, 0.0, 2026)["code"] == "E12"
    assert C.classify_one(
        _rec("SO1", "SOD1", customer_archive_failed=True), led, {}, 0.0, 2026
    )["code"] == "E10"


def test_no_ledger_never_auto():
    r = C.classify_one(_rec(), None, {}, 0.0, 2026)
    assert r["bucket"] == "hold" and r["code"] == "E2"
    res = C.classify_records([_rec()], None, {})
    assert res["counts"]["auto"] == 0


def test_same_row_hit_twice_both_held():
    """两条计划命中同一行 → 都不自动写。"""
    led = _led({1: {"so": "SO1", "sod": "", "yingshou": 100.0}})
    res = C.classify_records([_rec("SO1", "SODA", 100.0), _rec("SO1", "SODB", 100.0)], led, {})
    assert res["counts"]["auto"] == 0
    assert all(h["code"] == "E8" for h in res["hold"])


def test_case_id_is_sod_level():
    led = _led({1: {"so": "SO1", "sod": "", "yingshou": 100.0}})
    r = C.classify_one(_rec("SO1", "SOD1", 100.0), led, {}, 0.0, 2026)
    assert r["case_id"] == "AR_T|SO1|SOD1"


def test_no_proportion_in_source():
    """静态红线：判定脚本源码不许出现分摊逻辑。"""
    from pathlib import Path

    src = Path(C.__file__).read_text(encoding="utf-8")
    for bad in ("均分", "按比例分摊", "proportion"):
        assert bad not in src


# ══════════════════════════════════════════════════════════
# 六、真实金标端到端（有本地测试数据才跑）
# ══════════════════════════════════════════════════════════
GOLD_EXPORTS = GOLD_DIR / "01_智云导出"
GOLD_LEDGER = GOLD_DIR / "02_我的表副本" / "2026年盈亏核算表1-12月（副本）.xlsx"


@pytest.mark.skipif(not GOLD_LEDGER.is_file(), reason="无本地金标数据（真实数据不进仓库）")
def test_gold_20260722_end_to_end():
    """
    2026-07-22 真实 13 笔：判定结果必须与明妹当天手工填的**逐格一致**。
    这份数据抓出过三个真 bug（静默丢单 / SOD 拆行误判 / 冲预收规则错），是主回归闸。
    """
    import datetime as _dt

    from openpyxl import load_workbook

    payments = C.load_exports(GOLD_DIR)
    assert len(payments) == 13
    records = C.expand_payments(payments, {})
    assert len(records) == 157
    ledger = C.LedgerIndex(GOLD_LEDGER)
    res = C.classify_records(records, ledger, {})

    assert res["counts"]["auto"] == 143, res["e_code_dist"]
    assert res["counts"]["exception"] == 0
    assert res["e_code_dist"].get("E8") is None  # 同SO多行歧义已被 SOD 化解
    assert {p["ar"] for p in payments} == {r["ar"] for r in records}  # 一笔都没丢

    wb = load_workbook(GOLD_LEDGER, read_only=True, data_only=True)
    rows = list(wb["明细"].iter_rows(values_only=True))
    wb.close()
    hdr = list(rows[0])
    col = {
        "计提": hdr.index("计提金额"), "回款明细": hdr.index("回款明细"),
        "是否结账": hdr.index("是否结账（是/否）"), "收款时间": hdr.index("收款时间"),
        "收款方式": hdr.index("收款方式(支/汇/现)"), "实收SOD": hdr.index("实收金额"),
        "SO": hdr.index("新智云单号"),
    }

    def norm(v):
        if v is None:
            return ""
        if isinstance(v, (_dt.date, _dt.datetime)):
            return v.strftime("%Y-%m-%d")
        s = str(v).strip()
        try:
            return f"{float(s):.2f}"
        except (TypeError, ValueError):
            return s

    mismatched = []
    for it in res["auto"]:
        real = rows[it["ledger_row_ref"] - 1]
        assert str(real[col["SO"]]).strip() == it["so"]
        for k, v in it["five_cols"].items():
            if v is not None and norm(real[col[k]]) != norm(v):
                mismatched.append((it["case_id"], k, norm(real[col[k]]), norm(v)))
    assert not mismatched, mismatched[:10]
