#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从判定结果生成《流转写入计划_校验后.json》。

verdict=write 当且仅当：强三键唯一命中 + 可定位 + 有可写内容。
弱命中 / 0 / 多命中 → hand（须手填）。不写任何用户 Excel。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402
from flow_ledger import FlowLedger  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

STRONG = frozenset({"三键", "三键(含手续费)"})
_LOC_RE = re.compile(
    r"^(?P<file>.+)#(?P<sheet>.+) 第(?P<row>\d+)行（(?P<by>[^）]*)）\s*$"
)


def _all_items(result: dict) -> List[dict]:
    out: List[dict] = []
    for k in ("auto", "hold", "exception"):
        out.extend(result.get(k) or [])
    return out


def _parse_locate(locate: str) -> Dict[str, Any]:
    m = _LOC_RE.match((locate or "").strip())
    if not m:
        return {}
    return {
        "file": m.group("file"),
        "sheet": m.group("sheet"),
        "row_no": int(m.group("row")),
        "matched_by": m.group("by") or "",
    }


def _group_by_ar(items: List[dict]) -> Dict[str, List[dict]]:
    g: Dict[str, List[dict]] = {}
    for it in items:
        ar = it.get("ar") or "-"
        g.setdefault(ar, []).append(it)
    return g


def plan_item_for_ar(ar: str, items: List[dict], summary_row: Optional[dict]) -> dict:
    """为一笔到账(AR)生成一条流转写入计划。"""
    summary_row = summary_row or {}
    # 取该 AR 上最完整的 flow 信号
    best = items[0]
    for it in items:
        if it.get("flow_hits") == 1 or it.get("flow_file") or it.get("flow_locate"):
            best = it
            if it.get("flow_hits") == 1:
                break

    hits = best.get("flow_hits")
    matched_by = (best.get("flow_matched_by") or "").strip()
    file_ = best.get("flow_file") or ""
    sheet = best.get("flow_sheet") or ""
    row_no = best.get("flow_row_no")
    if (not file_ or not sheet or row_no is None) and best.get("flow_locate"):
        parsed = _parse_locate(best["flow_locate"])
        file_ = file_ or parsed.get("file") or ""
        sheet = sheet or parsed.get("sheet") or ""
        row_no = row_no if row_no is not None else parsed.get("row_no")
        if not matched_by:
            matched_by = parsed.get("matched_by") or ""

    so_list = []
    for it in items:
        so = (it.get("so") or "").strip()
        if so and so not in so_list:
            so_list.append(so)
    existing = best.get("flow_order_existing") or ""
    # 若 item 上只有单 SO 的 suggest，用全量 SO 重算
    order_suggest = FlowLedger.suggest_order_cell(so_list, existing) if so_list else (
        best.get("flow_order_suggest") or ""
    )
    updated = summary_row.get("流转表_是否更新应收款_建议")
    if updated is None:
        updated = ""
    updated = str(updated).strip()
    if updated in ("（空白）", "空白", "空"):
        updated = ""

    base = {
        "ar": ar,
        "file": file_,
        "sheet": sheet,
        "row_no": row_no,
        "hits": hits,
        "matched_by": matched_by,
        "order_suggest": order_suggest,
        "updated_suggest": updated,
        "flow_locate": best.get("flow_locate") or summary_row.get("flow_locate") or "",
        "so_list": so_list,
    }

    # skip：没有任何可写/可展示动作
    if hits is None and not matched_by and not file_:
        return {**base, "verdict": "hand", "reason": "未做三键或无流转表"}

    if hits == 0 or hits is None:
        return {**base, "verdict": "hand", "reason": matched_by or "流转表未命中"}

    if hits and int(hits) > 1:
        return {**base, "verdict": "hand", "reason": f"多命中 hits={hits}，须人工指定行"}

    # hits == 1
    if matched_by not in STRONG and not matched_by.startswith("三键"):
        # 弱命中
        if "名字不符" in matched_by or matched_by == "日期+金额(名字不符)":
            return {**base, "verdict": "hand", "reason": "弱命中（名字不符），须人工确认"}
        if matched_by not in STRONG:
            return {**base, "verdict": "hand", "reason": f"非强三键（{matched_by or '未知'}）"}

    if not file_ or not sheet or row_no is None:
        return {**base, "verdict": "hand", "reason": "强命中但无法解析 file/sheet/row"}

    # 有可写内容：单号建议变化，或是否更新有明确建议（含空白=刻意留空也算可写空）
    has_order = bool(str(order_suggest).strip())
    # 空白 updated 也是合法写入（清成空）
    has_updated_field = True  # always may write updated column to suggested value
    if not has_order and not has_updated_field:
        return {**base, "verdict": "skip", "reason": "无可写字段"}

    return {**base, "verdict": "write", "reason": ""}


def build_plan(result: dict) -> dict:
    items = _all_items(result)
    by_ar = _group_by_ar(items)
    summary_by = {s.get("ar") or "-": s for s in (result.get("ar_summary") or [])}
    plan_items = []
    for ar, group in by_ar.items():
        plan_items.append(plan_item_for_ar(ar, group, summary_by.get(ar)))
    counts = {"write": 0, "hand": 0, "skip": 0}
    for it in plan_items:
        counts[it["verdict"]] = counts.get(it["verdict"], 0) + 1
    return {
        "items": plan_items,
        "counts": counts,
        "payment_count": result.get("payment_count") or len(by_ar),
    }


def _latest(out_dir: Path, pattern: str) -> Optional[Path]:
    c = [p for p in sorted(out_dir.glob(pattern)) if not p.name.startswith("~$")]
    return c[-1] if c else None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="生成流转写入计划（不写用户表）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--result", default="", help="判定结果 json")
    ap.add_argument("--out", default="", help="输出 json 路径")
    args = ap.parse_args(argv)

    ws = Path(args.workspace)
    common.ensure_out_dirs(ws)
    out_dir = ws / "04_产出"
    result_path = Path(args.result) if args.result else _latest(out_dir, "判定结果_*.json")
    if not result_path or not result_path.is_file():
        print("ERROR: 找不到判定结果 json，请先跑 classify_hexiao.py", file=sys.stderr)
        return 2
    result = json.loads(result_path.read_text(encoding="utf-8"))
    plan = build_plan(result)
    out = Path(args.out) if args.out else out_dir / "流转写入计划_校验后.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    c = plan["counts"]
    print(
        f"流转写入计划 → {out}  "
        f"自动写 {c.get('write', 0)} · 手填 {c.get('hand', 0)} · 跳过 {c.get('skip', 0)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
