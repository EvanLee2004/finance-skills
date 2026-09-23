#!/usr/bin/env python3
"""从销售反馈抽出智云核对用的客户和单号。不登录，不下抽查结论。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import prepare  # noqa: E402


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="抽出客户、单号、交付月")
    parser.add_argument("--sales", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args(argv)
    found = [sheet for role, sheet, _score in prepare.sniff_file(args.sales) if role == "sales"]
    if len(found) != 1:
        print("status=ask")
        print("ask=销售反馈里没有唯一一张带账龄、结算阶段和单号的表。")
        print(f"sales_hits={len(found)}")
        print("rows=0")
        print("customers=0")
        return 2
    rows = prepare.load_sales(args.sales, found[0])
    lines = []
    customers = set()
    for row in rows:
        name = str(row["name"] or "").strip()
        if not name:
            continue
        customers.add(name)
        month = "" if row["month"] is None else str(row["month"])
        lines.append(f"{name}\t{row['so']}\t{month}")
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
    print("status=ok")
    print(f"rows={len(lines)}")
    print(f"customers={len(customers)}")
    print(f"out={args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
