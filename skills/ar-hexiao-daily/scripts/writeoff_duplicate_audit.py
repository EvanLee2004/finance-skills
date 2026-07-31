"""智云核销记录身份、超核销与条件性系统重复纠正。"""

from __future__ import annotations

import hashlib
import json
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Dict, Iterable, List, Optional, Tuple

TOLERANCE_CENTS = 100


def _cents(value: Any) -> Optional[int]:
    if value in (None, ""):
        return None
    try:
        return int(
            (Decimal(str(value)) * Decimal("100")).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP
            )
        )
    except (InvalidOperation, TypeError, ValueError):
        return None


def _money(cents: Optional[int]) -> Optional[float]:
    return None if cents is None else float(Decimal(cents) / Decimal("100"))


def _currency(value: Any) -> str:
    text = str(value or "").strip().upper().replace(" ", "")
    if text in {"人民币", "人民币CNY", "CNY", "RMB"}:
        return "CNY"
    return text


def _is_revoked(row: dict) -> bool:
    value = str(row.get("revoked") or row.get("是否已撤销") or "").strip().lower()
    return value in {"是", "true", "1", "已撤销", "撤销"}


def _sort_token(row: dict, index: int) -> Tuple[str, str, str, int]:
    return (
        str(row.get("snapshot_date") or ""),
        str(row.get("source") or ""),
        str(row.get("rowid") or ""),
        index,
    )


def _public_row(row: dict, disposition: str, reason: str) -> dict:
    return {
        "record_id": str(row.get("record_id") or "").strip(),
        "rowid": str(row.get("rowid") or "").strip(),
        "ar": str(row.get("ar") or "").strip(),
        "date": row.get("date"),
        "so": str(row.get("so") or "").strip(),
        "amount": row.get("amount"),
        "amount_local": row.get("amount_local"),
        "currency": str(row.get("currency") or "").strip(),
        "revoked": _is_revoked(row),
        "source": str(row.get("source") or ""),
        "snapshot_date": row.get("snapshot_date"),
        "disposition": disposition,
        "reason": reason,
    }


def audit_fingerprint(audits: Dict[str, dict]) -> str:
    """计划/校验共享的稳定审计指纹。"""

    def fix(value):
        if isinstance(value, (date, datetime)):
            return value.isoformat()
        if isinstance(value, dict):
            return {k: fix(v) for k, v in sorted(value.items()) if k != "fingerprint"}
        if isinstance(value, list):
            return [fix(v) for v in value]
        return value

    payload = json.dumps(
        fix(audits), ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest().upper()


def audit_parent_writeoffs(
    payment: dict,
    raw_rows: Iterable[dict],
    *,
    tolerance_cents: int = TOLERANCE_CENTS,
) -> Tuple[List[dict], dict]:
    """
    返回 `(logical_rows, audit)`。

    物理跨快照重复先按核销记录NUM处理；只有明显超核销时才按
    同父AR内的 SO+金额+币种口径折叠不同记录号。
    """
    rows = [dict(row) for row in raw_rows]
    marked: List[Optional[dict]] = [None] * len(rows)
    physical_kept: List[dict] = []

    by_id: Dict[str, List[Tuple[int, dict]]] = defaultdict(list)
    no_id: List[Tuple[int, dict]] = []
    for index, row in enumerate(rows):
        record_id = str(row.get("record_id") or "").strip()
        if record_id:
            by_id[record_id].append((index, row))
        else:
            no_id.append((index, row))

    for record_id in sorted(by_id):
        group = sorted(by_id[record_id], key=lambda item: _sort_token(item[1], item[0]))
        kept_index, kept_row = group[-1]
        for index, row in group[:-1]:
            marked[index] = _public_row(
                row,
                "physical_snapshot_duplicate",
                f"核销记录NUM={record_id} 已在更新快照保留一次",
            )
        if _is_revoked(kept_row):
            marked[kept_index] = _public_row(
                kept_row, "revoked", "截至目标核销日该物理核销记录已撤销"
            )
        else:
            marked[kept_index] = _public_row(
                kept_row, "kept", "不同核销记录NUM先作为独立原始记录保留"
            )
            physical_kept.append(kept_row)

    for index, row in no_id:
        if _is_revoked(row):
            marked[index] = _public_row(row, "revoked", "无记录号但明确标记为撤销")
        else:
            marked[index] = _public_row(
                row,
                "kept",
                "缺核销记录NUM，无法做物理跨快照去重；原样保留",
            )
            physical_kept.append(row)

    parent_local = _cents(payment.get("amount_local"))
    parent_orig = _cents(payment.get("amount_orig"))
    parent_currency = _currency(payment.get("currency"))
    local_comparable = parent_local is not None and all(
        _cents(row.get("amount_local")) is not None for row in physical_kept
    )
    original_comparable = (
        parent_orig is not None
        and bool(parent_currency)
        and all(
            _cents(row.get("amount")) is not None
            and _currency(row.get("currency")) == parent_currency
            for row in physical_kept
        )
    )

    basis = "local" if local_comparable else ("original" if original_comparable else "")
    parent_cents = parent_local if basis == "local" else parent_orig

    audit = {
        "status": "normal",
        "comparison_basis": basis or "unavailable",
        "raw_input_count": len(rows),
        "raw_record_count": len(physical_kept),
        "logical_record_count": len(physical_kept),
        "raw_total": None,
        "logical_total": None,
        "delta_raw": None,
        "delta_dedup": None,
        "duplicate_groups": [],
        "records": marked,
        "physical_snapshot_duplicate_count": sum(
            1 for row in marked if row and row["disposition"] == "physical_snapshot_duplicate"
        ),
        "revoked_count": sum(
            1 for row in marked if row and row["disposition"] == "revoked"
        ),
        "ignored_record_count": 0,
        "reason": "",
    }

    if not basis or parent_cents is None:
        if physical_kept:
            audit["status"] = "unresolved"
            audit["reason"] = "父回款与子核销无法形成统一本币或同币种原币比较口径"
        return ([] if audit["status"] == "unresolved" else physical_kept), audit

    def amount_cents(row: dict) -> int:
        value = row.get("amount_local") if basis == "local" else row.get("amount")
        cents = _cents(value)
        assert cents is not None
        return cents

    raw_total_cents = sum(amount_cents(row) for row in physical_kept)
    delta_raw_cents = parent_cents - raw_total_cents
    audit["raw_total"] = _money(raw_total_cents)
    audit["logical_total"] = _money(raw_total_cents)
    audit["delta_raw"] = _money(delta_raw_cents)
    audit["delta_dedup"] = _money(delta_raw_cents)

    if delta_raw_cents >= -tolerance_cents:
        audit["status"] = "tolerance" if delta_raw_cents < 0 else "normal"
        audit["reason"] = (
            "负差在1元容差内，不启动相同SO同金额业务去重"
            if delta_raw_cents < 0
            else "未出现明显超核销，不启动业务去重"
        )
        return physical_kept, audit

    if any(not str(row.get("record_id") or "").strip() for row in physical_kept):
        audit["status"] = "unresolved"
        audit["reason"] = "明显超核销且存在缺少核销记录NUM的明细，无法安全区分物理重复"
        return [], audit

    grouped: Dict[tuple, List[dict]] = defaultdict(list)
    both_amounts = all(
        _cents(row.get("amount")) is not None
        and _cents(row.get("amount_local")) is not None
        for row in physical_kept
    )
    for row in physical_kept:
        key = [
            str(row.get("so") or "").strip(),
            amount_cents(row),
            _currency(row.get("currency")),
        ]
        if both_amounts:
            key.extend([_cents(row.get("amount")), _cents(row.get("amount_local"))])
        grouped[tuple(key)].append(row)

    duplicate_groups: List[dict] = []
    proposed_ignored_ids = set()
    for group in grouped.values():
        ids = sorted(str(row.get("record_id") or "").strip() for row in group)
        if len(ids) <= 1 or len(set(ids)) <= 1:
            continue
        kept_id = ids[0]
        ignored_ids = ids[1:]
        proposed_ignored_ids.update(ignored_ids)
        sample = group[0]
        duplicate_groups.append(
            {
                "so": str(sample.get("so") or "").strip(),
                "amount": _money(amount_cents(sample)),
                "record_ids": ids,
                "kept_record_id": kept_id,
                "ignored_record_ids": ignored_ids,
            }
        )
    duplicate_groups.sort(key=lambda item: (item["so"], item["record_ids"]))
    audit["duplicate_groups"] = duplicate_groups

    if not duplicate_groups:
        audit["status"] = "unresolved"
        audit["reason"] = "明显超核销，但不存在同父AR内相同SO同金额的精确重复组"
        return [], audit

    dedup_rows = [
        row
        for row in physical_kept
        if str(row.get("record_id") or "").strip() not in proposed_ignored_ids
    ]
    logical_total_cents = sum(amount_cents(row) for row in dedup_rows)
    delta_dedup_cents = parent_cents - logical_total_cents
    audit["logical_record_count"] = len(dedup_rows)
    audit["logical_total"] = _money(logical_total_cents)
    audit["delta_dedup"] = _money(delta_dedup_cents)

    if delta_dedup_cents < -tolerance_cents:
        audit["status"] = "unresolved"
        audit["reason"] = "折叠全部精确重复组后仍明显超核销，禁止部分自动处理"
        return [], audit

    audit["status"] = "recovered"
    audit["ignored_record_count"] = len(proposed_ignored_ids)
    audit["reason"] = "智云疑似系统重复核销，本次每组只按一次处理"
    for index, row in enumerate(rows):
        record_id = str(row.get("record_id") or "").strip()
        if record_id in proposed_ignored_ids and marked[index]["disposition"] == "kept":
            marked[index] = _public_row(
                row,
                "system_duplicate_ignored",
                "明显超核销且精确重复折叠后恢复到1元容差内",
            )
    return dedup_rows, audit

