"""Plan-only reservations for receipts without per-order writeoff amounts.

Reservations are not payment history. Connected plans must all pass before any
member can be published; only the existing readback path commits allocations.
"""
from collections import defaultdict
import common
import baseline_receipts as BR

RULE = "arrival_date_ar_ascending"
ZERO = "OK_FALLBACK_ZERO_ALLOCATION"


def ordered(payments):
    fallback = [p for p in payments if not p.get("writeoffs")]
    by_so = defaultdict(list)
    for p in fallback:
        for so in {o.get("so") for o in p.get("orders") or [] if o.get("so")}:
            by_so[so].append(p)
    shared = {id(p) for group in by_so.values() if len(group) > 1 for p in group}
    for p in fallback:
        if id(p) in shared and (not common.norm_date(p.get("arrival_date")) or not p.get("ar")):
            raise ValueError("多笔无明细回款缺少到账日期或AR单号，无法确定固定分配顺序")
    ars = [p.get("ar") for p in fallback]
    if len(ars) != len(set(ars)):
        raise ValueError("父回款AR单号重复，禁止重复分配")
    return sorted(payments, key=lambda p: (
        common.norm_date(p.get("arrival_date")) or common.norm_date("9999-12-31"),
        str(p.get("ar") or "")))


def reserve(p, reservations):
    audit = p.get("_parent_fallback_allocation") or {}
    for row in audit.get("allocations") or []:
        so = row["so"]
        original, local = row.get("allocated_orig"), row.get("allocated_local")
        if audit.get("reused_successful_allocation"):
            if "applied_cases" not in audit:
                continue
            paid = sum(c["amount_local"] for c in audit["applied_cases"].values() if c["so"] == so)
            pending = max(float(local or 0) - paid, 0)
            original = float(original or 0) * pending / local if local else 0
            local = pending
        if float(local or original or 0) > 0.011:
            reservations[so].append((p["ar"], round(float(original or 0), 2), round(float(local or 0), 2)))


def bind_groups(records, dependencies):
    groups = []
    for ar, previous in dependencies.items():
        if not previous:
            continue
        group = {ar, *previous}
        overlap = [g for g in groups if g & group]
        for g in overlap:
            group.update(g)
            groups.remove(g)
        groups.append(group)
    for group in groups:
        members = [r for r in records if r.get("ar") in group]
        keys = sorted({"|".join([r.get("ar") or "-", r.get("so") or "-", *([r["sod"]] if r.get("sod") else [])]) for r in members})
        for r in members:
            r["fallback_batch_cases"] = keys
            r["fallback_batch_ars"] = sorted(group)


def guard(items, *, checked=False, universe=None):
    """Reject an entire reservation component if any member is absent or blocked."""
    universe = items if universe is None else universe
    if not checked:
        for r in items:
            ars = r.get("fallback_batch_ars") or []
            if ars:
                members = [item for item in universe if item.get("ar") in ars]
                keys = {item["case_id"] for item in members}
                keys.update("missing:" + ar for ar in set(ars) - {item.get("ar") for item in members})
                r["fallback_batch_cases"] = sorted(keys)
    by_case = {r.get("case_id"): r for r in universe}
    def good(r):
        return (r.get("_check") or {}).get("verdict") in {"write", "skip"} if checked else r.get("bucket") == "auto"
    failed = set()
    for r in items:
        required = r.get("fallback_batch_cases") or []
        if required and any(key not in by_case or not good(by_case[key]) for key in required):
            failed.update(required)
    for r in items:
        if r.get("case_id") not in failed:
            continue
        reason = "同批顺序分配的关联回款尚未全部通过，禁止使用未完成的前序分配继续核销"
        if checked and good(r):
            r["_check"] = {"verdict": "conflict", "reason": reason}
        elif not checked and r.get("bucket") == "auto":
            r.update(bucket="hold", code="E_FALLBACK_SEQUENCE_DEPENDENCY", reason=reason, five_cols={}, derived_cols={})
            r.pop("row_operation", None)


def ledger_so_rows(ledger, so):
    if ledger is None:
        return {}
    sods = {ledger.row_snapshot[ref].get("sod") or "" for ref in ledger.so_index.get(so, [])}
    return {key: value for sod in sods for key, value in BR.ledger_rows(ledger, so, sod).items()}
