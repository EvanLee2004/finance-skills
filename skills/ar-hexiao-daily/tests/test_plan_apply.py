# -*- coding: utf-8 -*-
"""plan → validate → execute：写入前校验 + 写副本 + 写后回读。"""
import datetime as dt
import json
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from conftest import LEDGER_FULL  # noqa: E402

import validate_plan as V  # noqa: E402
import apply_to_copy as A  # noqa: E402

HDR = ["部门", "销售人员", "客户名称", "单号", "新智云单号", "应收金额",
       "计提金额", "回款明细", "是否结账（是/否）", "收款时间", "收款方式(支/汇/现)", "实收金额"]


def _ledger(tmp_path, rows):
    """造一张最小盈亏表：表头 + 若干行。rows=[(SO, SOD, 五列值 or None)]"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(HDR)
    for so, sod, five in rows:
        r = ["部", "人", "客", "AB", so, 100, None, None, None, None, None, sod]
        if five:
            r[6], r[7], r[8], r[9], r[10] = (
                five.get("计提"), five.get("回款明细"), five.get("是否结账"),
                five.get("收款时间"), five.get("收款方式"),
            )
        ws.append(r)
    p = tmp_path / "盈亏.xlsx"
    wb.save(str(p))
    return p


def _item(row, so="SO26010001", sod="SOD26010001", **five):
    f = {"计提": 100.0, "回款明细": 100.0, "是否结账": "是",
         "收款时间": "2026-07-08", "收款方式": "汇", "实收SOD": sod}
    f.update(five)
    return {"case_id": f"AR1|{so}", "ar": "AR1", "so": so, "sod": sod,
            "ledger_row_ref": row, "five_cols": f}


# ---------- 校验 ----------

def test_empty_row_is_writable(tmp_path):
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    rows = V.read_ledger_rows(led)
    assert V.check_one(_item(2), rows)["verdict"] == "write"


def test_already_filled_same_is_skip(tmp_path):
    five = {"计提": 100.0, "回款明细": 100.0, "是否结账": "是",
            "收款时间": dt.date(2026, 7, 8), "收款方式": "汇"}
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", five)])
    rows = V.read_ledger_rows(led)
    assert V.check_one(_item(2), rows)["verdict"] == "skip"


def test_already_filled_different_is_conflict(tmp_path):
    """她填的和我们算的不一样 → 绝不覆盖，交给人（跨月冲预收那条就是这么抓出来的）。"""
    five = {"计提": 100.0, "回款明细": 100.0, "是否结账": "是",
            "收款时间": dt.date(2026, 7, 8), "收款方式": "冲预收"}
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", five)])
    rows = V.read_ledger_rows(led)
    res = V.check_one(_item(2), rows)
    assert res["verdict"] == "conflict" and "收款方式" in res["reason"]


def test_row_shifted_is_conflict(tmp_path):
    """她插过行 → 行号指到别人家去了，必须拦住。"""
    led = _ledger(tmp_path, [("SO_OTHER", "SOD_OTHER", None)])
    rows = V.read_ledger_rows(led)
    res = V.check_one(_item(2), rows)
    assert res["verdict"] == "conflict" and "不是计划里的" in res["reason"]


def test_bad_value_is_conflict(tmp_path):
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    rows = V.read_ledger_rows(led)
    assert V.check_one(_item(2, 是否结账="也许"), rows)["verdict"] == "conflict"
    assert V.check_one(_item(2, 收款方式="随便"), rows)["verdict"] == "conflict"
    assert V.check_one(_item(2, 计提="一百块"), rows)["verdict"] == "conflict"


def test_missing_row_ref_is_conflict(tmp_path):
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    rows = V.read_ledger_rows(led)
    it = _item(2)
    it["ledger_row_ref"] = None
    assert V.check_one(it, rows)["verdict"] == "conflict"


def test_two_plans_same_row_conflict(tmp_path):
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    rows = V.read_ledger_rows(led)
    plan = {"auto": [_item(2), _item(2, so="SO26010001")]}
    res = V.validate(plan, rows)
    assert res["counts"]["write"] == 1 and res["counts"]["conflict"] == 1


# ---------- 写入 ----------

def test_apply_writes_and_verifies(tmp_path):
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    out = tmp_path / "已回填.xlsx"
    changes = A.write_plan(led, out, [_item(2)])
    assert len(changes) == 1
    assert A.verify_written(out, [_item(2)]) == []
    ws = openpyxl.load_workbook(str(out))["明细"]
    assert ws.cell(2, 7).value == 100.0          # 计提
    assert ws.cell(2, 9).value == "是"            # 是否结账
    assert ws.cell(2, 12).value == "SOD26010001"  # 实收金额列存 SOD


def test_apply_never_touches_source(tmp_path):
    """写的是新文件，她给的副本必须一个字节都不动。"""
    import hashlib
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    before = hashlib.sha256(led.read_bytes()).hexdigest()
    A.write_plan(led, tmp_path / "out.xlsx", [_item(2)])
    assert hashlib.sha256(led.read_bytes()).hexdigest() == before


def test_in_place_writes_into_copy_and_backs_up(tmp_path):
    """就地模式（明妹要的）：直接写进她那份副本，并留一份写前备份、不留临时文件。"""
    import hashlib

    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    before = hashlib.sha256(led.read_bytes()).hexdigest()
    checked = tmp_path / "checked.json"
    checked.write_text(
        json.dumps({"write": [_item(2)], "skip": [], "conflict": []},
                   ensure_ascii=False, default=str),
        encoding="utf-8",
    )
    rc = A.main(["--checked", str(checked), "--ledger", str(led),
                 "--report", str(tmp_path / "r.xlsx"), "--in-place"])
    assert rc == 0
    # 副本被就地改了
    ws = openpyxl.load_workbook(str(led))["明细"]
    assert ws.cell(2, 7).value == 100.0            # 计提
    assert ws.cell(2, 9).value == "是"             # 是否结账
    # 写前备份存在，且等于写前内容（真出事能还原）
    backups = list((led.parent / "备份").glob("盈亏_备份_*.xlsx"))
    assert len(backups) == 1
    assert hashlib.sha256(backups[0].read_bytes()).hexdigest() == before
    # 成功后不留临时文件
    assert not list(led.parent.glob(".盈亏_写入中_*"))


def test_partial_leaves_jiti_empty(tmp_path):
    """部分核销：计提留空就是留空，不许写成 0（写 0 会被当成已计提）。"""
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    out = tmp_path / "out.xlsx"
    A.write_plan(led, out, [_item(2, 计提=None, 是否结账="否")])
    ws = openpyxl.load_workbook(str(out))["明细"]
    assert ws.cell(2, 7).value is None
    assert ws.cell(2, 9).value == "否"


def test_verify_catches_wrong_write(tmp_path):
    """回读比对必须真能发现写错——否则这道保险等于没有。"""
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    out = tmp_path / "out.xlsx"
    A.write_plan(led, out, [_item(2)])
    wb = openpyxl.load_workbook(str(out))
    wb["明细"].cell(2, 7).value = 999          # 人为篡改
    wb.save(str(out))
    assert A.verify_written(out, [_item(2)]) != []


def test_main_refuses_when_conflicts(tmp_path):
    """有冲突没处理就想写 → 必须拒绝（除非显式 --force）。"""
    led = _ledger(tmp_path, [("SO26010001", "SOD26010001", None)])
    checked = tmp_path / "checked.json"
    checked.write_text(json.dumps({"write": [_item(2)], "skip": [], "conflict": [_item(3)]},
                                  ensure_ascii=False, default=str), encoding="utf-8")
    rc = A.main(["--checked", str(checked), "--ledger", str(led),
                 "--out", str(tmp_path / "o.xlsx"), "--report", str(tmp_path / "r.xlsx")])
    assert rc == 2


@pytest.mark.skipif(not LEDGER_FULL.is_file(), reason="无真实全年盈亏表")
def test_real_ledger_structure_survives_write(tmp_path):
    """写她的真表副本：透视表/外链/公式必须全须全尾。"""
    import zipfile
    out = tmp_path / "已回填.xlsx"
    rows = V.read_ledger_rows(LEDGER_FULL)
    target = next(r for r, v in rows.items() if v["SO"].startswith("SO"))
    it = _item(target, so=rows[target]["SO"], sod=rows[target]["SOD"] or "SOD_X")
    A.write_plan(LEDGER_FULL, out, [it])

    def n(p, key):
        with zipfile.ZipFile(p) as z:
            return sum(1 for x in z.namelist() if key.lower() in x.lower())
    for key in ("pivotTable", "pivotCache", "externalLink"):
        assert n(out, key) == n(LEDGER_FULL, key), f"{key} 写完少了"
    assert A.verify_written(out, [it]) == []
