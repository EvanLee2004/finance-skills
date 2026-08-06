#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""只读补跑审计：检查因父记录移到最新核销日而需从子明细还原的历史订单。"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Optional, Sequence

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import common  # noqa: E402
import classify_hexiao as classify  # noqa: E402


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="审计历史核销日迁移明细（只读，不写业务表）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--date-from", default="", help="只审计此日及以后；补跑时应明确给出")
    ap.add_argument("--date-to", default="", help="只审计此日及以前；补跑时应明确给出")
    ap.add_argument("--out", default="", help="可选：写 JSON 报告")
    args = ap.parse_args(argv)

    ws = common.resolve_workspace(args.workspace)
    date_from = common.resolve_batch_date(args.date_from) if args.date_from else None
    date_to = common.resolve_batch_date(args.date_to) if args.date_to else None
    if args.date_from and date_from is None:
        print(f"ERROR: 认不出 --date-from {args.date_from!r}", file=sys.stderr)
        return 2
    if args.date_to and date_to is None:
        print(f"ERROR: 认不出 --date-to {args.date_to!r}", file=sys.stderr)
        return 2
    if date_from and date_to and date_from > date_to:
        print("ERROR: --date-from 不能晚于 --date-to", file=sys.stderr)
        return 2

    found = classify.assess_shifted_detail_dates(ws, date_from=date_from, date_to=date_to)
    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(found, ensure_ascii=False, indent=2), encoding="utf-8")

    pending = {day: info for day, info in found.items() if info.get("needs_rerun")}
    if not found:
        print("指定日期范围内没有需要从后续核销日文件还原的历史子明细。")
        return 0
    if not pending:
        print("历史子明细都已进入对应日期的判定结果，无需重跑。")
        return 0
    print("补跑审计发现尚未进入原核销日判定的订单：")
    for day, info in pending.items():
        print(
            f"- {day}: 尚缺 {len(info['missing_order_keys'])}/{info['rows']} 个 AR/SO，"
            f"当前可从 {', '.join(info['sources'])} 还原"
        )
    print("这些日期逐日重跑并生成增补《核销日清》；日清与写前校验通过后直接写工作副本。")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
