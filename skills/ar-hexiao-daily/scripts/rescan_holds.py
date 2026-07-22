#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第8步：挂账重扫。读挂账台账 + 今日判定/数据，标记可补做；幂等（连跑两遍一致）。
只维护 工作区/03_台账/挂账台账.xlsx，不写用户表。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402

HEADERS = [
    "AR",
    "SO列表",
    "挂起日",
    "E码",
    "原因",
    "重扫次数",
    "最近重扫日",
    "状态",
]


def ledger_path(workspace: Path) -> Path:
    d = workspace / "03_台账"
    d.mkdir(parents=True, exist_ok=True)
    return d / "挂账台账.xlsx"


def load_ledger(path: Path) -> List[dict]:
    if not path.is_file():
        return []
    import openpyxl

    wb = openpyxl.load_workbook(str(path), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    headers = [str(c).strip() if c is not None else "" for c in rows[0]]
    # 允许缺列时报错
    required = ["AR", "状态"]
    for r in required:
        if r not in headers:
            raise ValueError(f"挂账台账缺列 {r}；实际表头：{headers}")
    idx = {h: i for i, h in enumerate(headers)}
    out = []
    for row in rows[1:]:
        vals = list(row)
        if not any(vals):
            continue

        def g(name, default=""):
            i = idx.get(name)
            if i is None or i >= len(vals) or vals[i] is None:
                return default
            return vals[i]

        out.append(
            {
                "AR": str(g("AR") or "").strip(),
                "SO列表": str(g("SO列表") or "").strip(),
                "挂起日": str(g("挂起日") or "")[:10],
                "E码": str(g("E码") or "").strip(),
                "原因": str(g("原因") or "").strip(),
                "重扫次数": int(g("重扫次数") or 0),
                "最近重扫日": str(g("最近重扫日") or "")[:10],
                "状态": str(g("状态") or "挂起").strip(),
            }
        )
    return out


def save_ledger(path: Path, rows: List[dict]) -> None:
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "挂账台账"
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(
            [
                r.get("AR") or "",
                r.get("SO列表") or "",
                r.get("挂起日") or "",
                r.get("E码") or "",
                r.get("原因") or "",
                r.get("重扫次数") or 0,
                r.get("最近重扫日") or "",
                r.get("状态") or "挂起",
            ]
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))


def merge_from_classify(rows: List[dict], result: dict, today: str) -> List[dict]:
    """把今日 hold/exception 并入台账；已存在则更新原因，不重复加行。"""
    by_ar = {r["AR"]: r for r in rows if r.get("AR")}
    for item in (result.get("hold") or []) + (result.get("exception") or []):
        ar = item.get("ar") or ""
        if not ar:
            # 无 AR 用 SO 作键
            ar = f"SO:{(item.get('so') or '').strip()}"
        if not ar or ar == "SO:":
            continue
        if ar in by_ar:
            rec = by_ar[ar]
            if rec.get("状态") == "已完成":
                continue
            rec["E码"] = item.get("code") or rec.get("E码") or ""
            rec["原因"] = item.get("reason") or rec.get("原因") or ""
            rec["SO列表"] = item.get("so") or rec.get("SO列表") or ""
        else:
            by_ar[ar] = {
                "AR": ar,
                "SO列表": item.get("so") or "",
                "挂起日": today,
                "E码": item.get("code") or "",
                "原因": item.get("reason") or "",
                "重扫次数": 0,
                "最近重扫日": "",
                "状态": "挂起",
            }
    # auto 命中同一 AR → 可补做
    auto_ars = set()
    for item in result.get("auto") or []:
        ar = item.get("ar") or ""
        if ar:
            auto_ars.add(ar)
        so = item.get("so") or ""
        if so:
            auto_ars.add(f"SO:{so}")
    for ar, rec in by_ar.items():
        if rec.get("状态") == "已完成":
            continue
        if ar in auto_ars:
            rec["状态"] = "可补做"
    return list(by_ar.values())


def rescan(rows: List[dict], result: Optional[dict], today: str) -> List[dict]:
    """重扫：递增次数、更新日期；状态机保持幂等。"""
    if result:
        rows = merge_from_classify(rows, result, today)
    # 稳定排序保证幂等写出
    rows = sorted(rows, key=lambda r: (r.get("AR") or "", r.get("SO列表") or ""))
    for r in rows:
        if r.get("状态") == "已完成":
            continue
        r["重扫次数"] = int(r.get("重扫次数") or 0) + 1
        r["最近重扫日"] = today
    return rows


def snapshot(rows: List[dict]) -> str:
    """用于 diff 的稳定序列化（不含重扫次数与最近重扫日的可比核心？）

    幂等定义：连跑两遍后，除「重扫次数/最近重扫日」外业务字段一致；
    且第二遍相对第一遍次数 +1。验收 A7：台账 diff 为空——
    故第二次跑时若 today 相同，我们采用：同一天重复跑不重复 +1（幂等）。
    """
    core = []
    for r in sorted(rows, key=lambda x: x.get("AR") or ""):
        core.append(
            {
                "AR": r.get("AR"),
                "SO列表": r.get("SO列表"),
                "挂起日": r.get("挂起日"),
                "E码": r.get("E码"),
                "原因": r.get("原因"),
                "状态": r.get("状态"),
                "重扫次数": r.get("重扫次数"),
                "最近重扫日": r.get("最近重扫日"),
            }
        )
    return json.dumps(core, ensure_ascii=False, sort_keys=True)


def rescan_idempotent(rows: List[dict], result: Optional[dict], today: str) -> List[dict]:
    """同一天已重扫过则不再增加次数（保证连跑两遍 diff 空）。"""
    if result:
        rows = merge_from_classify(rows, result, today)
    rows = sorted(rows, key=lambda r: (r.get("AR") or "", r.get("SO列表") or ""))
    for r in rows:
        if r.get("状态") == "已完成":
            continue
        if str(r.get("最近重扫日") or "")[:10] == today:
            # 已扫过今日：不改次数
            continue
        r["重扫次数"] = int(r.get("重扫次数") or 0) + 1
        r["最近重扫日"] = today
    return rows


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="挂账重扫（幂等）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--result", default="", help="今日判定结果 json（可选）")
    ap.add_argument("--init-from-result", action="store_true", help="用判定 hold 初始化台账")
    args = ap.parse_args(argv)

    common.ensure_out_dirs()
    ws = Path(args.workspace)
    path = ledger_path(ws)
    today = dt.date.today().isoformat()

    result = None
    if args.result:
        result = json.loads(Path(args.result).read_text(encoding="utf-8"))
    else:
        out_dir = ws / "04_产出"
        cands = sorted(out_dir.glob("判定结果_*.json")) if out_dir.is_dir() else []
        if cands:
            result = json.loads(cands[-1].read_text(encoding="utf-8"))

    try:
        rows = load_ledger(path)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    if not rows and result is None:
        # 空台账也算成功（可初始化）
        save_ledger(path, [])
        print("挂账台账为空（0 行）；状态=就绪")
        print(f"台账: {path}")
        return 0

    rows = rescan_idempotent(rows, result, today)
    save_ledger(path, rows)

    n_hold = sum(1 for r in rows if r.get("状态") == "挂起")
    n_ready = sum(1 for r in rows if r.get("状态") == "可补做")
    n_done = sum(1 for r in rows if r.get("状态") == "已完成")
    print(f"重扫完成 挂起={n_hold} 可补做={n_ready} 已完成={n_done} 合计={len(rows)}")
    print(f"台账: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
