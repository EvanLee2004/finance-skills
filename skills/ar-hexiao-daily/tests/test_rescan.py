# -*- coding: utf-8 -*-
"""挂账重扫幂等。"""
import json
from pathlib import Path

import openpyxl

import rescan_holds as R


def test_save_load_roundtrip(tmp_path):
    path = tmp_path / "挂账台账.xlsx"
    rows = [
        {
            "AR": "AR1",
            "SO列表": "SO1",
            "挂起日": "2026-07-01",
            "E码": "E1",
            "原因": "分笔",
            "重扫次数": 0,
            "最近重扫日": "",
            "状态": "挂起",
        }
    ]
    R.save_ledger(path, rows)
    loaded = R.load_ledger(path)
    assert len(loaded) == 1
    assert loaded[0]["AR"] == "AR1"
    assert loaded[0]["状态"] == "挂起"


def test_rescan_idempotent_same_day(tmp_path):
    path = tmp_path / "挂账台账.xlsx"
    rows = [
        {
            "AR": "AR1",
            "SO列表": "SO1",
            "挂起日": "2026-07-01",
            "E码": "E2",
            "原因": "未交付",
            "重扫次数": 0,
            "最近重扫日": "",
            "状态": "挂起",
        }
    ]
    R.save_ledger(path, rows)
    today = "2026-07-22"
    r1 = R.rescan_idempotent(R.load_ledger(path), None, today)
    R.save_ledger(path, r1)
    snap1 = R.snapshot(r1)
    r2 = R.rescan_idempotent(R.load_ledger(path), None, today)
    R.save_ledger(path, r2)
    snap2 = R.snapshot(r2)
    assert snap1 == snap2
    assert r2[0]["重扫次数"] == 1  # 只 +1 一次


def test_merge_auto_marks_ready():
    rows = [
        {
            "AR": "AR9",
            "SO列表": "SO9",
            "挂起日": "2026-07-01",
            "E码": "E2",
            "原因": "未交付",
            "重扫次数": 1,
            "最近重扫日": "2026-07-10",
            "状态": "挂起",
        }
    ]
    result = {
        "auto": [{"ar": "AR9", "so": "SO9"}],
        "hold": [],
        "exception": [],
    }
    out = R.merge_from_classify(rows, result, "2026-07-22")
    assert out[0]["状态"] == "可补做"


def test_merge_new_hold():
    rows = []
    result = {
        "auto": [],
        "hold": [{"ar": "AR_NEW", "so": "SO_NEW", "code": "E1", "reason": "分笔"}],
        "exception": [],
    }
    out = R.merge_from_classify(rows, result, "2026-07-22")
    assert len(out) == 1
    assert out[0]["E码"] == "E1"


def test_main_empty_ok(tmp_path):
    ws = tmp_path / "工作区"
    (ws / "03_台账").mkdir(parents=True)
    (ws / "04_产出").mkdir(parents=True)
    rc = R.main(["--workspace", str(ws)])
    assert rc == 0


def test_main_twice_idempotent(tmp_path):
    ws = tmp_path / "工作区"
    (ws / "03_台账").mkdir(parents=True)
    (ws / "04_产出").mkdir(parents=True)
    result = {
        "auto": [],
        "hold": [{"ar": "ARX", "so": "SOX", "code": "E1", "reason": "分笔"}],
        "exception": [],
        "counts": {"auto": 0, "hold": 1, "exception": 0},
    }
    rp = ws / "04_产出" / "判定结果_20260722.json"
    rp.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    assert R.main(["--workspace", str(ws), "--result", str(rp)]) == 0
    p = ws / "03_台账" / "挂账台账.xlsx"
    s1 = R.snapshot(R.load_ledger(p))
    assert R.main(["--workspace", str(ws), "--result", str(rp)]) == 0
    s2 = R.snapshot(R.load_ledger(p))
    assert s1 == s2
