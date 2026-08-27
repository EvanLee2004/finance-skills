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

STRONG = frozenset({
    "三键",
    "三键(含手续费)",
    "三键(原币公式)",
    "三键(原币公式含手续费)",
    "三键(中英文对照)",
    "三键(含手续费,中英文对照)",
    "三键(原币公式,中英文对照)",
    "三键(原币公式含手续费,中英文对照)",
})
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


def _delivery_amount(item: dict) -> Optional[float]:
    """取智云本次分类保留下来的 SO 最新本币交付金额。"""
    source = item.get("split_payment_source") or {}
    for value in (
        source.get("so_delivery_local"),
        item.get("so_delivery_local"),
        source.get("delivery_local"),
        item.get("delivery_local"),
    ):
        if value is None or value == "":
            continue
        try:
            return float(value)
        except (TypeError, ValueError):
            continue
    return None


def _rich_runs(text: str, red_sos: List[str]) -> List[dict]:
    """按整行生成富文本段；原规则是“部分”时仅未核销 SO 行标红。"""
    red = {str(x).strip().upper() for x in red_sos if str(x).strip()}
    lines = str(text or "").replace("\r", "").split("\n")
    runs: List[dict] = []
    for index, line in enumerate(lines):
        suffix = "\n" if index < len(lines) - 1 else ""
        sos = re.findall(r"(?<![A-Z0-9])(SO[A-Z0-9]+)(?![A-Z0-9])", line, re.I)
        color = "FFFF0000" if any(so.upper() in red for so in sos) else ""
        runs.append({"text": line + suffix, "color": color})
    return runs


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
    so_amounts: Dict[str, Optional[float]] = {}
    so_outcomes: Dict[str, dict] = {}
    for it in items:
        so = (it.get("so") or "").strip()
        if so and so not in so_list:
            so_list.append(so)
        if so and (so not in so_amounts or so_amounts[so] is None):
            so_amounts[so] = _delivery_amount(it)
        if so:
            outcome = so_outcomes.setdefault(
                so, {"so": so, "buckets": [], "case_ids": [], "codes": []}
            )
            outcome["buckets"].append(it.get("bucket") or "")
            if it.get("case_id"):
                outcome["case_ids"].append(it["case_id"])
            if it.get("code"):
                outcome["codes"].append(it["code"])
    existing = (best.get("flow_order_existing") or "").strip()
    # 取数后先写「SO + 最新交付金额」，一个 SO 一行；表内其它已有订单保持不动。
    order_suggest = FlowLedger.suggest_order_amount_cell(
        [(so, so_amounts.get(so)) for so in so_list], existing
    )
    if not str(order_suggest).strip() and best.get("flow_order_suggest"):
        order_suggest = str(best.get("flow_order_suggest") or "")
    # 仍空且表上本来有单号 → 保住表上的（apply 也不会写空单号）
    if not str(order_suggest).strip() and existing:
        order_suggest = existing

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
        # 命中行的身份，交给 apply_flow 在写入前再核一次（防插行导致行号错位）
        "identity": best.get("flow_identity") or {},
        "so_list": so_list,
        "so_entries": [
            {"so": so, "delivery_amount": so_amounts.get(so)} for so in so_list
        ],
        "so_outcomes": list(so_outcomes.values()),
        "red_sos": [],
        "order_rich_runs": _rich_runs(order_suggest, []),
        "write_order": bool(str(order_suggest).strip()),  # 空单号不写列，防抹掉已有
        "write_updated": True,  # 是否更新列：空白=刻意留空，仍可写空
    }

    # skip：没有任何可写/可展示动作
    if hits is None and not matched_by and not file_:
        return {**base, "verdict": "hand", "reason": "未做三键或无流转表"}

    if hits == 0 or hits is None:
        return {**base, "verdict": "hand", "reason": matched_by or "流转表未命中"}

    if hits and int(hits) > 1:
        return {**base, "verdict": "hand", "reason": f"多命中 hits={hits}，须人工指定行"}

    # hits == 1：准入必须是精确强三键集合（禁止 startswith 放宽）
    if matched_by not in STRONG:
        if "名字不符" in matched_by or matched_by == "日期+金额(名字不符)":
            return {**base, "verdict": "hand", "reason": "弱命中（名字不符），须人工确认"}
        return {**base, "verdict": "hand", "reason": f"非强三键（{matched_by or '未知'}）"}

    if not file_ or not sheet or row_no is None:
        return {**base, "verdict": "hand", "reason": "强命中但无法解析 file/sheet/row"}

    missing_delivery = [
        entry["so"] for entry in base["so_entries"]
        if entry.get("delivery_amount") is None
    ]
    if missing_delivery:
        return {
            **base,
            "verdict": "hand",
            "reason": "SO交付金额缺失，不能自动前置写入：" + "、".join(missing_delivery),
        }

    # 至少要写单号或是否更新之一；单号空则只写是否更新
    if not base["write_order"] and not base["write_updated"]:
        return {**base, "verdict": "skip", "reason": "无可写字段"}

    return {**base, "verdict": "write", "reason": ""}


def finalize_plan_after_ledger(flow_plan: dict, checked_plan: dict) -> dict:
    """根据盈亏表实际可写/已写结果回填“是/部分/空白”和未核销红字。"""
    good = {
        ((x.get("ar") or "").strip(), (x.get("so") or "").strip())
        for key in ("write", "skip")
        for x in (checked_plan.get(key) or [])
    }
    bad = {
        ((x.get("ar") or "").strip(), (x.get("so") or "").strip())
        for x in (checked_plan.get("conflict") or [])
    }
    good_cases = {
        str(x.get("case_id") or "").strip()
        for key in ("write", "skip")
        for x in (checked_plan.get(key) or [])
        if str(x.get("case_id") or "").strip()
    }
    bad_cases = {
        str(x.get("case_id") or "").strip()
        for x in (checked_plan.get("conflict") or [])
        if str(x.get("case_id") or "").strip()
    }
    finalized = json.loads(json.dumps(flow_plan, ensure_ascii=False))
    for item in finalized.get("items") or []:
        if item.get("verdict") != "write":
            continue
        ar = (item.get("ar") or "").strip()
        completed: List[str] = []
        incomplete: List[str] = []
        outcomes = item.get("so_outcomes") or [
            {"so": so, "buckets": ["auto"]} for so in (item.get("so_list") or [])
        ]
        for outcome in outcomes:
            so = (outcome.get("so") or "").strip()
            buckets = [str(x or "") for x in (outcome.get("buckets") or [])]
            cases = [str(x).strip() for x in (outcome.get("case_ids") or []) if str(x).strip()]
            if cases:
                is_complete = all(case in good_cases for case in cases) and not any(
                    case in bad_cases for case in cases
                )
            else:
                is_complete = bool(so and (ar, so) in good and (ar, so) not in bad)
            is_complete = is_complete and all(x in ("auto", "ready") for x in buckets)
            (completed if is_complete else incomplete).append(so)
            outcome["completed"] = bool(is_complete)
        if completed and not incomplete:
            status = "是"
        elif completed:
            status = "部分"
        else:
            status = ""
        item["updated_suggest"] = status
        item["red_sos"] = incomplete if status == "部分" else []
        item["order_rich_runs"] = _rich_runs(item.get("order_suggest") or "", item["red_sos"])
        item["phase"] = "post_ledger"
    finalized["phase"] = "post_ledger"
    return finalized


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
        "selection": {
            "mode": "all_results",
        },
    }


def _latest(out_dir: Path, pattern: str) -> Optional[Path]:
    c = [p for p in sorted(out_dir.glob(pattern)) if not p.name.startswith("~$")]
    return c[-1] if c else None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="生成流转写入计划（不写用户表）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--result", default="", help="判定结果 json")
    ap.add_argument("--out", default="", help="输出 json 路径")
    # 防呆：AI 常把 --hexiao-date 顺手传给链上每个脚本。本脚本用不到，
    # 但收下总比 argparse 报错中断整条链好（2026-07-25 opencode 实测踩到）。
    ap.add_argument("--hexiao-date", default="", help="（本脚本用不到，收下防止链路中断）")
    args = ap.parse_args(argv)

    ws = common.ensure_out_dirs(args.workspace)  # 解析真工作区，防产出分家
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
