#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
确认后统一写入：先盈亏明细，成功后再流转安全子集。

无 --confirmed → 拒绝任何写入。
盈亏失败 → 不写流转。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402
import apply_to_copy  # noqa: E402
import apply_flow  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="确认后：盈亏 → 流转")
    ap.add_argument("--checked", required=True, help="盈亏 写入计划_校验后.json")
    ap.add_argument("--flow-plan", default="", help="流转写入计划_校验后.json；空则跳过流转")
    ap.add_argument("--ledger", required=True, help="盈亏副本")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--confirmed", action="store_true")
    ap.add_argument("--in-place", action="store_true", help="盈亏就地写")
    ap.add_argument("--flow-in-place", action="store_true", help="流转就地写")
    ap.add_argument("--force", action="store_true", help="盈亏跳过冲突只写可写")
    args = ap.parse_args(argv)

    if not args.confirmed:
        print(
            "ERROR: 缺人工确认，拒绝统一写入。\n"
            "  先出《核销日清》→ 她确认 → 再加 --confirmed。",
            file=sys.stderr,
        )
        return 2

    # 1) 盈亏
    ledger_args = [
        "--checked", str(args.checked),
        "--ledger", str(args.ledger),
        "--confirmed",
    ]
    if args.in_place:
        ledger_args.append("--in-place")
    if args.force:
        ledger_args.append("--force")
    rc1 = apply_to_copy.main(ledger_args)
    if rc1 != 0:
        print(
            f"ERROR: 盈亏写入失败 EXIT:{rc1}，**不写流转**。",
            file=sys.stderr,
        )
        return rc1

    # 2) 流转
    flow_plan = args.flow_plan
    if not flow_plan:
        ws = Path(args.workspace)
        cand = ws / "04_产出" / "流转写入计划_校验后.json"
        if cand.is_file():
            flow_plan = str(cand)
    if not flow_plan or not Path(flow_plan).is_file():
        print("WARN: 无流转写入计划，跳过流转写入（盈亏已成功）。")
        return 0

    flow_args = [
        "--plan", flow_plan,
        "--workspace", str(args.workspace),
        "--confirmed",
    ]
    if args.flow_in_place:
        flow_args.append("--in-place")
    rc2 = apply_flow.main(flow_args)
    if rc2 != 0:
        print(
            f"WARN: 盈亏已写成功，但流转写入 EXIT:{rc2}（请看流转变更/手填项）。",
            file=sys.stderr,
        )
        return rc2
    print("统一写入完成：盈亏 + 流转")
    return 0


if __name__ == "__main__":
    sys.exit(main())
