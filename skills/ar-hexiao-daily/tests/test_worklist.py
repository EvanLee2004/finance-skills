# -*- coding: utf-8 -*-
import json
import tempfile
from pathlib import Path

import openpyxl

import build_worklist as W


def _sample_result():
    return {
        "auto": [
            {
                "ar": "AR1",
                "so": "SO2601",
                "sod": "SOD2601",
                "customer_masked": "测*户",
                "locate_hint": "在盈亏『明细』按「新智云单号」筛选：SO2601（禁止用行号）",
                "five_cols": {
                    "计提": 100.0,
                    "回款明细": 100.0,
                    "是否结账": "是",
                    "收款时间": "2026-07-08",
                    "收款方式": "汇",
                    "实收SOD": "SOD2601",
                },
                "current_values": {"计提": None, "回款明细": None},
                "reason": "ok",
                "channel": "duizhang",
            }
        ],
        "hold": [
            {
                "ar": "AR2",
                "so": "SO2602",
                "sod": "",
                "customer_masked": "挂*户",
                "code": "E1",
                "reason": "分笔",
                "candidates": [],
            }
        ],
        "exception": [
            {
                "ar": "AR3",
                "so": "",
                "sod": "",
                "customer_masked": "",
                "code": "E7",
                "reason": "无单号",
            }
        ],
        "counts": {"auto": 1, "hold": 1, "exception": 1, "total": 3},
        "e_code_dist": {"OK": 1, "E1": 1, "E7": 1},
    }


def test_worklist_three_sheets():
    tmp = Path(tempfile.mkdtemp())
    out = tmp / "清单.xlsx"
    W.build_workbook(_sample_result(), out)
    wb = openpyxl.load_workbook(out)
    # 三个用户页签必在；ar_summary 存在时另有「按到账汇总」
    assert {"今天能填", "挂账待办", "异常"} <= set(wb.sheetnames)
    ws = wb["今天能填"]
    headers = [c.value for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}
    assert "怎么找到这行(按SO筛选)" in col
    assert "应填_计提" in col
    assert "当前_回款明细" in col
    assert "这笔到账在流转表哪一行" in col  # 迭代v2：流转表定位
    # 无「行号」作为定位列名（她的表会插行，行号隔天失效）
    assert not any(h == "行号" for h in headers)
    row2 = [c.value for c in ws[2]]
    assert row2[col["SO"]] == "SO2601"
    assert "禁止用行号" in str(row2[col["怎么找到这行(按SO筛选)"]])


def test_worklist_hold_sheet_has_action():
    tmp = Path(tempfile.mkdtemp())
    out = tmp / "清单.xlsx"
    W.build_workbook(_sample_result(), out)
    wb = openpyxl.load_workbook(out)
    wh = wb["挂账待办"]
    assert wh.max_row >= 2
    hh = [c.value for c in wh[1]]
    assert wh.cell(2, hh.index("码") + 1).value == "E1"
    assert "这笔到账在流转表哪一行" in hh


def test_main_from_result_json(tmp_path):
    res = tmp_path / "判定结果_20260722.json"
    res.write_text(json.dumps(_sample_result(), ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out.xlsx"
    rc = W.main(["--result", str(res), "--out", str(out), "--workspace", str(tmp_path)])
    assert rc == 0
    assert out.is_file()
