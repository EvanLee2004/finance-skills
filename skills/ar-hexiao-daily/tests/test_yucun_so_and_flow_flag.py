# -*- coding: utf-8 -*-
"""取数落地：四张表读得对；流转空/部分=待办。"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import classify_hexiao as C  # noqa: E402
import rescan_holds as R  # noqa: E402


def test_flow_needs_update_empty_and_partial():
    assert R.flow_needs_update(None) is True
    assert R.flow_needs_update("") is True
    assert R.flow_needs_update("  ") is True
    assert R.flow_needs_update("部分") is True
    assert R.flow_needs_update("部份") is True
    assert R.flow_needs_update("是") is False
    assert R.flow_needs_update("yes") is False


def _write(path, headers, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)


def _mk_exports(root: Path, *, with_writeoff=True, with_sod=True):
    exp = root / "01_智云导出"
    exp.mkdir(parents=True, exist_ok=True)
    _write(exp / "回款记录_test.xlsx",
           ["回款记录ID", "核销日期", "到账日期", "到账金额/原币", "到账金额/本币",
            "手续费/原币", "原币币种", "回款类型", "核销状态", "开票客户"],
           [["AR26079999", "2026-07-22", "2026-07-17", 100, 100, 0,
             "人民币CNY", "预存回款", "核销成功", "测试客户甲"]])
    _write(exp / "订单交付_test.xlsx",
           ["回款记录ID", "SO", "交付额/原币", "汇率", "订单名称"],
           [["AR26079999", "SO26070111", 250, 1, "某单"]])
    if with_writeoff:
        _write(exp / "核销明细_test.xlsx",
               ["回款记录NUM", "核销日期", "本次核销金额", "币种", "汇率", "SO", "订单名称"],
               [["AR26079999", "2026-07-22", 100, "人民币CNY", 1, "SO26070111", "某单"],
                # 上个月的旧核销：累积子表里会有，必须被日期过滤掉，否则重复回填
                ["AR26079999", "2026-06-15", 500, "人民币CNY", 1, "SO26070111", "某单"]])
    if with_sod:
        _write(exp / "订单明细_test.xlsx",
               ["SO", "SOD", "交付额/原币"],
               [["SO26070111", "SOD26070222", 60],
                ["SO26070111", "SOD26070221", 40],
                ["SO26070111", "SOD26070220", 150]])
    return exp


def test_load_exports_reads_four_tables(tmp_path):
    _mk_exports(tmp_path)
    pays = C.load_exports(tmp_path)
    assert len(pays) == 1
    p = pays[0]
    assert p["ar"] == "AR26079999"
    assert p["orders"] == [{"so": "SO26070111", "deliver": 250.0, "rate": 1.0,
                            "currency": "", "name": "某单"}]
    # 累积子表按核销日期过滤：只留本次的 100，不含 6-15 的 500
    assert p["writeoffs"] == {"SO26070111": 100.0}
    assert len(p["sod_lines"]["SO26070111"]) == 3


def test_expand_uses_sod_subset(tmp_path):
    _mk_exports(tmp_path)
    pays = C.load_exports(tmp_path)
    recs = C.expand_payments(pays, {})
    # 本次核销 100 = SOD…222(60) + SOD…221(40)，SOD…220(150) 不动
    assert {r["sod"] for r in recs} == {"SOD26070222", "SOD26070221"}
    assert sum(r["amount_orig"] for r in recs) == 100.0


def test_no_writeoff_table_means_full_settle(tmp_path):
    """核销明细表整份没有 → 全额核销（对整笔/自动核销类就是这样）。"""
    _mk_exports(tmp_path, with_writeoff=False)
    # 到账额改成 250 才等于交付额
    _write(tmp_path / "01_智云导出" / "回款记录_test.xlsx",
           ["回款记录ID", "核销日期", "到账日期", "到账金额/原币", "到账金额/本币",
            "手续费/原币", "原币币种", "回款类型", "核销状态", "开票客户"],
           [["AR26079999", "2026-07-22", "2026-07-21", 250, 250, 0,
             "人民币CNY", "整笔回款", "手动核销", "测试客户甲"]])
    recs = C.expand_payments(C.load_exports(tmp_path), {})
    assert len(recs) == 3
    assert sum(r["amount_orig"] for r in recs) == 250.0


def test_missing_order_table_is_hard_error(tmp_path):
    """缺「订单交付」→ 报错退出并指路重新取数，不拿旧对账表凑合。"""
    import pytest

    exp = tmp_path / "01_智云导出"
    exp.mkdir(parents=True)
    _write(exp / "回款记录_test.xlsx", ["回款记录ID", "核销日期"], [["AR1", "2026-07-22"]])
    with pytest.raises(C.InputError) as ei:
        C.load_exports(tmp_path)
    assert "订单交付" in str(ei.value)
