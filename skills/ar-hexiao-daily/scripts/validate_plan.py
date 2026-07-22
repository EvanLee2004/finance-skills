#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第 7 步 a：**写入前校验计划**（官方 plan → validate → execute 模式的中间环）。

为什么要有这一步：判定结果是"打算怎么填"，但从判定到写入之间，她的表可能已经变了
（月初贴交付会插行、部分核销会在上方插行 → **行号立刻失效**）。
直接按行号写 = 把值写到别人家的行上。所以写之前逐条复核，过不了的不写、并说清为什么。

用法：
    python3 scripts/validate_plan.py --plan 判定结果.json --ledger 盈亏副本.xlsx \
        --out 04_产出/写入计划_校验后.json
退出码：0=有可写的（或全跳过）；2=输入不可用；**1=存在冲突**（有笔需要人看）
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

FIVE = ["计提", "回款明细", "是否结账", "收款时间", "收款方式"]
VALID_JIEZHANG = {"是", "否"}
VALID_WAY = {"汇", "冲预收", "支", "现"}


def _norm(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    # Excel 里 300 与 300.0 是同一个数
    try:
        f = float(s)
        return f"{f:.2f}"
    except (TypeError, ValueError):
        return s


def read_ledger_rows(path: Path) -> Dict[int, dict]:
    """把盈亏『明细』整表读成 {行号: {列名: 值}}，用于逐条复核。"""
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    if "明细" not in wb.sheetnames:
        wb.close()
        raise ValueError(f"盈亏表无『明细』sheet：{wb.sheetnames}")
    ws = wb["明细"]
    aliases = common.load_aliases()
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hrow, headers = common.find_header_row(
        all_rows, "盈亏明细", ["SO", "SOD", "计提", "回款明细", "是否结账"], aliases
    )
    cols = common.resolve_columns(
        headers,
        "盈亏明细",
        ["SO", "SOD", "计提", "回款明细", "是否结账", "收款时间", "收款方式"],
        aliases,
    )
    out: Dict[int, dict] = {}
    for i, row in enumerate(all_rows, start=1):
        if i <= hrow + 1:
            continue
        vals = list(row)

        def cell(key):
            idx = cols.get(key)
            return vals[idx] if idx is not None and idx < len(vals) else None

        out[i] = {
            "SO": str(cell("SO") or "").strip(),
            "SOD": str(cell("SOD") or "").strip(),
            "计提": cell("计提"),
            "回款明细": cell("回款明细"),
            "是否结账": cell("是否结账"),
            "收款时间": cell("收款时间"),
            "收款方式": cell("收款方式"),
        }
    return out


def check_one(item: dict, rows: Dict[int, dict]) -> dict:
    """
    单条复核 → {verdict: write|skip|conflict, reason}
    - write    ：行号对得上、目标格是空的、值合法 → 可以写
    - skip     ：已经填过且与计划一致 → 幂等跳过（重复跑不重复写）
    - conflict ：行号对不上 / 已填但不一致 / 值不合法 → 不写，交给人看
    """
    ref = item.get("ledger_row_ref")
    five = item.get("five_cols") or {}
    so, sod = (item.get("so") or "").strip(), (item.get("sod") or "").strip()

    if not ref:
        return {"verdict": "conflict", "reason": "判定结果里没有行号，无法定位"}
    row = rows.get(int(ref))
    if row is None:
        return {"verdict": "conflict", "reason": f"第 {ref} 行在表里不存在了（表被删过行？）"}

    # ① 行号还指着同一单吗——她插过行的话这里必然对不上
    if sod and row["SOD"] and row["SOD"] != sod:
        return {
            "verdict": "conflict",
            "reason": f"第 {ref} 行现在是 {row['SOD']}，不是计划里的 {sod}（表在判定之后被插过行）",
        }
    if so and row["SO"] and row["SO"] != so:
        return {
            "verdict": "conflict",
            "reason": f"第 {ref} 行现在是 {row['SO']}，不是计划里的 {so}（表被改过）",
        }

    # ② 值本身合法吗
    if five.get("是否结账") not in VALID_JIEZHANG:
        return {"verdict": "conflict", "reason": f"是否结账取值异常：{five.get('是否结账')!r}"}
    if five.get("收款方式") not in VALID_WAY:
        return {"verdict": "conflict", "reason": f"收款方式取值异常：{five.get('收款方式')!r}"}
    for k in ("计提", "回款明细"):
        v = five.get(k)
        if v is None:
            continue  # 部分核销时计提本就留空
        try:
            float(v)
        except (TypeError, ValueError):
            return {"verdict": "conflict", "reason": f"{k} 不是数字：{v!r}"}

    # ③ 目标格现在是什么
    filled = [k for k in FIVE if _norm(row.get(k)) not in ("", "None")]
    if not filled:
        return {"verdict": "write", "reason": "五列为空，可写"}

    same = all(
        _norm(row.get(k)) == _norm(five.get(k))
        for k in FIVE
        if five.get(k) is not None
    )
    if same:
        return {"verdict": "skip", "reason": "已经填过且与本次一致（幂等跳过）"}
    diff = [
        f"{k}: 表里={_norm(row.get(k))!r} 计划={_norm(five.get(k))!r}"
        for k in FIVE
        if five.get(k) is not None and _norm(row.get(k)) != _norm(five.get(k))
    ]
    return {
        "verdict": "conflict",
        "reason": "这行已经填过，且和本次算的不一样 → " + "；".join(diff),
    }


def validate(plan: dict, rows: Dict[int, dict]) -> dict:
    items = list(plan.get("auto") or [])
    checked: List[dict] = []
    seen_rows: Dict[int, str] = {}
    for it in items:
        res = check_one(it, rows)
        ref = it.get("ledger_row_ref")
        # 同一行被两条计划命中 → 都不写（谁对谁错要人定）
        if res["verdict"] == "write" and ref in seen_rows:
            res = {
                "verdict": "conflict",
                "reason": f"第 {ref} 行被两笔计划同时命中（另一笔 {seen_rows[ref]}），需人工指定",
            }
        elif res["verdict"] == "write":
            seen_rows[ref] = it.get("case_id") or ""
        checked.append({**it, "_check": res})
    buckets = {"write": [], "skip": [], "conflict": []}
    for c in checked:
        buckets[c["_check"]["verdict"]].append(c)
    return {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "counts": {k: len(v) for k, v in buckets.items()},
        **buckets,
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="写入前校验计划（plan→validate→execute）")
    ap.add_argument("--plan", required=True, help="判定结果 json")
    ap.add_argument("--ledger", required=True, help="盈亏核算表副本（只读）")
    ap.add_argument("--out", default="", help="校验后计划 json")
    args = ap.parse_args(argv)

    plan_p, ledger_p = Path(args.plan), Path(args.ledger)
    if not plan_p.is_file():
        print(f"ERROR: 找不到判定结果 {plan_p}", file=sys.stderr)
        return 2
    if not ledger_p.is_file():
        print(f"ERROR: 找不到盈亏表 {ledger_p}", file=sys.stderr)
        return 2

    plan = json.loads(plan_p.read_text(encoding="utf-8"))
    try:
        rows = read_ledger_rows(ledger_p)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    result = validate(plan, rows)
    out_p = Path(args.out) if args.out else plan_p.with_name("写入计划_校验后.json")
    out_p.parent.mkdir(parents=True, exist_ok=True)
    out_p.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    c = result["counts"]
    print(f"校验完成 可写={c['write']} 幂等跳过={c['skip']} 冲突={c['conflict']}")
    for x in result["conflict"][:10]:
        print(f"  ⚠ {x.get('case_id')}: {x['_check']['reason']}")
    if c["conflict"] > 10:
        print(f"  …另有 {c['conflict']-10} 条冲突，详见 {out_p.name}")
    print(f"计划: {out_p}")
    return 1 if c["conflict"] else 0


if __name__ == "__main__":
    sys.exit(main())
