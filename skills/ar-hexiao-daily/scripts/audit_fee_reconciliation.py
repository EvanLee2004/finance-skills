#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""只读审计：用历史人工表分析手续费父回款的净额落单规律。

该脚本只用于规则研发/回归取证，不参与日常 classify，也不写任何业务表。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402
from classify_hexiao import (  # noqa: E402
    LedgerIndex,
    _localize_amount,
    _payment_local,
    load_exports,
)


def audit(payments: list[dict], reference: LedgerIndex, target_date=None) -> dict:
    parents: list[dict] = []
    target_iso = target_date.isoformat() if target_date is not None else None
    for payment in payments:
        fee = common.to_number(payment.get("fee")) or 0.0
        if fee <= 0.011:
            continue
        arrival, _ = _payment_local(payment, {})
        fee_local, _ = _localize_amount(float(fee), payment, {})
        amount_orig = common.to_number(payment.get("amount_orig"))
        amount_local = common.to_number(payment.get("amount_local"))
        parent_implied_rate = (
            round(float(amount_local) / float(amount_orig), 8)
            if amount_orig is not None
            and amount_local is not None
            and abs(float(amount_orig)) > 0.011
            else None
        )
        rows: list[dict] = []
        sod_lines = payment.get("sod_lines") or {}
        for order in payment.get("orders") or []:
            so = str(order.get("so") or "").strip()
            if not so:
                continue
            ref_row = reference.comparison_row(so, "")
            snap = reference.row_snapshot.get(ref_row) if ref_row is not None else None
            delivery_local = None
            delivery_orig = common.to_number(order.get("deliver"))
            row_rate = common.to_number(order.get("rate"))
            if order.get("deliver") is not None:
                delivery_local, _ = _localize_amount(
                    float(order["deliver"]),
                    payment,
                    {},
                    row_rate=order.get("rate"),
                )
            reference_received = common.to_number((snap or {}).get("huikuan"))
            reference_accrual = common.to_number((snap or {}).get("jiti"))
            so_sod_lines = sod_lines.get(so) or []
            sod_delivery_orig = (
                round(
                    sum(float(line["deliver"]) for line in so_sod_lines),
                    2,
                )
                if so_sod_lines
                and all(line.get("deliver") is not None for line in so_sod_lines)
                else None
            )
            sod_delivery_local = None
            if sod_delivery_orig is not None:
                sod_delivery_local, _ = _localize_amount(
                    sod_delivery_orig,
                    payment,
                    {},
                    row_rate=order.get("rate"),
                )
            rows.append(
                {
                    "so": so,
                    "order_delivery_orig": delivery_orig,
                    "order_rate": row_rate,
                    "order_currency": str(order.get("currency") or "").strip(),
                    "order_delivery": delivery_local,
                    "sod_count": len(so_sod_lines),
                    "sod_delivery_orig": sod_delivery_orig,
                    "sod_delivery_local": sod_delivery_local,
                    "reference_row": ref_row,
                    "reference_accrual": reference_accrual,
                    "reference_received": reference_received,
                    "reference_closed": str((snap or {}).get("jiezhang") or "").strip(),
                    "reference_received_date": (
                        common.norm_date((snap or {}).get("shoukuan_time")).isoformat()
                        if common.norm_date((snap or {}).get("shoukuan_time"))
                        else None
                    ),
                    "reference_received_way": str(
                        (snap or {}).get("shoukuan_way") or ""
                    ).strip(),
                    "reference_in_target_batch": (
                        bool(target_iso)
                        and (
                            common.norm_date((snap or {}).get("shoukuan_time")).isoformat()
                            if common.norm_date((snap or {}).get("shoukuan_time"))
                            else None
                        )
                        == target_iso
                    ),
                    "reference_over_delivery_orig": (
                        round(float(reference_received) / float(delivery_orig), 8)
                        if reference_received is not None
                        and delivery_orig is not None
                        and abs(float(delivery_orig)) > 0.011
                        else None
                    ),
                    "reference_over_delivery_local": (
                        round(float(reference_received) / float(delivery_local), 8)
                        if reference_received is not None
                        and delivery_local is not None
                        and abs(float(delivery_local)) > 0.011
                        else None
                    ),
                    "reference_over_sod_orig": (
                        round(float(reference_received) / float(sod_delivery_orig), 8)
                        if reference_received is not None
                        and sod_delivery_orig is not None
                        and abs(float(sod_delivery_orig)) > 0.011
                        else None
                    ),
                    "reference_over_sod_local": (
                        round(float(reference_received) / float(sod_delivery_local), 8)
                        if reference_received is not None
                        and sod_delivery_local is not None
                        and abs(float(sod_delivery_local)) > 0.011
                        else None
                    ),
                    "writeoff_orig": common.to_number(
                        (payment.get("writeoffs") or {}).get(so)
                    ),
                    "writeoff_local": common.to_number(
                        (payment.get("writeoffs_local") or {}).get(so)
                    ),
                    "cumulative_writeoff_orig": common.to_number(
                        (payment.get("cumulative_writeoffs") or {}).get(so)
                    ),
                    "cumulative_writeoff_local": common.to_number(
                        (payment.get("cumulative_writeoffs_local") or {}).get(so)
                    ),
                }
            )
        delivery_values = [
            float(row["order_delivery"])
            for row in rows
            if row["order_delivery"] is not None
        ]
        reference_values = [
            float(row["reference_received"])
            for row in rows
            if row["reference_received"] is not None
        ]
        per_row_deductions = [
            round(float(row["order_delivery"]) - float(row["reference_received"]), 2)
            for row in rows
            if row["order_delivery"] is not None
            and row["reference_received"] is not None
        ]
        nonzero = [x for x in per_row_deductions if abs(x) > 0.011]
        delivery_sum = round(sum(delivery_values), 2)
        reference_sum = round(sum(reference_values), 2)
        parents.append(
            {
                "ar": str(payment.get("ar") or ""),
                "currency": str(payment.get("currency") or "").strip(),
                "arrival_orig": amount_orig,
                "arrival_net": arrival,
                "arrival_local": amount_local,
                "parent_implied_rate": parent_implied_rate,
                "fee_orig": fee,
                "fee_local": fee_local,
                "order_count": len(rows),
                "writeoff_count": len(payment.get("writeoffs") or {}),
                "delivery_sum": delivery_sum,
                "reference_sum": reference_sum,
                "all_rows_referenced": len(reference_values) == len(rows),
                "reference_rows_in_target_batch": sum(
                    1 for row in rows if row["reference_in_target_batch"]
                ),
                "all_references_in_target_batch": (
                    bool(rows)
                    and all(row["reference_in_target_batch"] for row in rows)
                ),
                "delivery_equals_net_plus_fee": (
                    arrival is not None
                    and fee_local is not None
                    and abs(delivery_sum - (float(arrival) + float(fee_local))) <= 0.011
                ),
                "delivery_orig_equals_arrival_orig_plus_fee": (
                    amount_orig is not None
                    and abs(
                        sum(
                            float(row["order_delivery_orig"])
                            for row in rows
                            if row["order_delivery_orig"] is not None
                        )
                        - (float(amount_orig) + float(fee))
                    )
                    <= 0.011
                ),
                "reference_equals_net": (
                    arrival is not None
                    and len(reference_values) == len(rows)
                    and abs(reference_sum - float(arrival)) <= 0.011
                ),
                "reference_over_arrival_orig": (
                    round(reference_sum / float(amount_orig), 8)
                    if amount_orig is not None and abs(float(amount_orig)) > 0.011
                    else None
                ),
                "reference_over_arrival_local": (
                    round(reference_sum / float(arrival), 8)
                    if arrival is not None and abs(float(arrival)) > 0.011
                    else None
                ),
                "net_over_gross_orig": (
                    round(float(amount_orig) / (float(amount_orig) + float(fee)), 8)
                    if amount_orig is not None
                    and abs(float(amount_orig) + float(fee)) > 0.011
                    else None
                ),
                "one_row_absorbs_whole_fee": (
                    len(nonzero) == 1
                    and fee_local is not None
                    and abs(nonzero[0] - float(fee_local)) <= 0.011
                ),
                "rows": rows,
            }
        )
    return {"fee_parent_count": len(parents), "parents": parents}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="只读分析手续费净额落单规律")
    ap.add_argument("--workspace", required=True)
    ap.add_argument("--hexiao-date", required=True)
    ap.add_argument("--reference-ledger", required=True, help="历史人工表，仅用于规则研发取证")
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    target_date = common.resolve_batch_date(args.hexiao_date)
    if target_date is None:
        print(f"ERROR: 认不出日期 {args.hexiao_date!r}", file=sys.stderr)
        return 2
    payments = load_exports(Path(args.workspace), target_date=target_date)
    result = audit(
        payments,
        LedgerIndex(Path(args.reference_ledger)),
        target_date=target_date,
    )
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    matched = sum(
        1
        for parent in result["parents"]
        if parent["all_rows_referenced"]
    )
    same_batch = sum(
        1
        for parent in result["parents"]
        if parent["all_rows_referenced"]
        and parent["all_references_in_target_batch"]
    )
    print(
        f"手续费父回款 {result['fee_parent_count']} 笔；"
        f"对比表完整匹配 {matched} 笔；同批完整匹配 {same_batch} 笔；结果 {out}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
