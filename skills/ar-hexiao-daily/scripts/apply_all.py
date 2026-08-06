#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
校验与《核销日清》生成后统一写入：先盈亏明细，成功后再流转安全子集。

不再要求人工确认；--confirmed 仅为旧命令兼容参数。
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


def _record_done(args, *, ledger_written: bool, flow_written: bool) -> None:
    """
    写完表 → 在跑批台账上把这个核销日标成「已写表·收工」。
    只有落到这一步，`batch_ledger gaps` 才不会再把这天算成没跑过。
    """
    try:
        import json

        import batch_ledger
        import fallback_allocation_ledger

        plan = json.loads(Path(args.checked).read_text(encoding="utf-8"))
        d = common.norm_date(plan.get("hexiao_date"))
        if d is None:
            print(
                "WARN: 这份计划里没有核销日期（旧版计划），跑批台账没法登记这一天。",
                file=sys.stderr,
            )
            return
        allocation_path, allocation_added = fallback_allocation_ledger.commit(
            Path(args.workspace), plan
        )
        batch_ledger.record(
            Path(args.workspace), d, "applied",
            written={"盈亏": bool(ledger_written), "流转": bool(flow_written)},
        )
        if plan.get("parent_fallback_allocations"):
            print(
                f"父回款顺序分配台账：新增 {allocation_added} 笔，已复核保存至 {allocation_path.name}"
            )
        print(f"跑批台账：{common.date_cn(d)} 已标记「已写表·收工」")
    except Exception as e:
        print(
            f"WARN: 写后台账登记失败（不影响已写入的数据）：{type(e).__name__}",
            file=sys.stderr,
        )


def _resnapshot_sources(workspace) -> None:
    """统一写入全部成功并登记台账后，把合法新状态登记为下一轮源文件基线。"""
    try:
        import verify_sources

        ws = common.resolve_workspace(workspace, quiet=True)
        verify_sources.do_snapshot(ws)
    except Exception as e:
        print(
            f"WARN: 最终源文件指纹刷新失败（不影响已验证的写入）：{type(e).__name__}",
            file=sys.stderr,
        )


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="日清后直接写入：盈亏 → 流转")
    ap.add_argument("--checked", required=True, help="盈亏 写入计划_校验后.json")
    ap.add_argument("--flow-plan", default="", help="流转写入计划_校验后.json；空则跳过流转")
    ap.add_argument("--ledger", required=True, help="盈亏副本")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument(
        "--confirmed",
        action="store_true",
        help="已废弃的兼容参数；现在日清与写前校验通过后可直接写入",
    )
    ap.add_argument("--in-place", action="store_true", help="盈亏就地写")
    ap.add_argument("--flow-in-place", action="store_true", help="流转就地写")
    ap.add_argument("--force", action="store_true", help="盈亏跳过冲突只写可写")
    args = ap.parse_args(argv)

    # 1) 盈亏
    ledger_args = [
        "--checked", str(args.checked),
        "--ledger", str(args.ledger),
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
        ws = common.resolve_workspace(args.workspace)
        cand = ws / "04_产出" / "流转写入计划_校验后.json"
        if cand.is_file():
            flow_plan = str(cand)
    if not flow_plan or not Path(flow_plan).is_file():
        print("WARN: 无流转写入计划，跳过流转写入（盈亏已成功）。")
        _record_done(args, ledger_written=True, flow_written=False)
        _resnapshot_sources(args.workspace)
        return 0

    flow_args = [
        "--plan", flow_plan,
        "--workspace", str(args.workspace),
    ]
    if args.flow_in_place:
        flow_args.append("--in-place")
    rc2 = apply_flow.main(flow_args)
    if rc2 != 0:
        print(
            f"WARN: 盈亏已写成功，但流转写入 EXIT:{rc2}（请看流转变更/手填项）。",
            file=sys.stderr,
        )
        _record_done(args, ledger_written=True, flow_written=False)
        return rc2
    print("统一写入完成：盈亏 + 流转")
    _record_done(args, ledger_written=True, flow_written=True)
    _resnapshot_sources(args.workspace)
    return 0


if __name__ == "__main__":
    sys.exit(main())
