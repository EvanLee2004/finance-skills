"""智云核销记录身份与金额来源优先级审计。"""

from __future__ import annotations

import hashlib
import json
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Dict, Iterable, List, Optional, Tuple

TOLERANCE_CENTS = 100  # 整笔回款的父到账与有效核销明细允许 1 元差额。


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


def _field_present(order: dict, key: str) -> bool:
    flag = f"{key}_present"
    if flag in order:
        return bool(order.get(flag))
    return order.get(key) not in (None, "")


def _latest_orders_by_so(payment: dict) -> Tuple[List[dict], str]:
    """父 AR 下每个 SO 只保留最新关联行；缺 SO 时返回错误。"""
    grouped: Dict[str, dict] = {}
    for order in payment.get("orders") or []:
        so = str(order.get("so") or "").strip()
        if not so:
            return [], "整笔回款存在缺少SO的关联订单"
        grouped[so] = dict(order)
    if not grouped:
        return [], "整笔回款没有关联订单，无法核对订单已核销金额"
    return [grouped[so] for so in sorted(grouped)], ""


def _audit_whole_payment_orders(
    payment: dict,
    audit: dict,
    *,
    tolerance_cents: int,
) -> Tuple[List[dict], dict]:
    """整笔回款：订单已核销金额优先，全部缺失时才按交付额兜底。"""
    orders, order_error = _latest_orders_by_so(payment)
    audit["order_count"] = len(orders)
    audit["order_records"] = []
    audit["fallback_used"] = False
    audit["effective_order_amount_count"] = 0
    if order_error:
        audit["status"] = "unresolved"
        audit["error_code"] = "E_PARENT_WRITEOFF_MISMATCH"
        audit["reason"] = order_error
        return [], audit

    written_presence = [
        _field_present(order, "written_off")
        or _field_present(order, "written_off_local")
        for order in orders
    ]
    if any(written_presence) and not all(written_presence):
        audit["status"] = "unresolved"
        audit["error_code"] = "E_PARENT_WRITEOFF_MISMATCH"
        audit["reason"] = "整笔回款仅部分订单取得订单已核销金额，禁止混用交付额补齐"
        audit["effective_order_amount_count"] = sum(written_presence)
        audit["effective_detail_count"] = sum(written_presence)
        return [], audit

    if all(written_presence):
        source = "order_written_off"
        original_key = "written_off"
        local_key = "written_off_local"
        audit["status"] = "order_written_off_authoritative"
        audit["reason"] = "整笔回款以订单已核销金额为父AR守恒和逐SO金额口径"
    else:
        source = "delivery_fallback"
        original_key = "deliver"
        local_key = "deliver_local"
        audit["status"] = "delivery_fallback"
        audit["fallback_used"] = True
        audit["reason"] = "全部订单均无订单已核销金额，使用完整订单交付额兜底"

    original_cents = [_cents(order.get(original_key)) for order in orders]
    local_cents = [_cents(order.get(local_key)) for order in orders]
    original_complete = all(value is not None for value in original_cents)
    local_complete = all(value is not None for value in local_cents)
    parent_local = _cents(
        payment.get("total_amount_local")
        if payment.get("total_amount_local") is not None
        else payment.get("amount_local")
    )
    parent_orig = _cents(
        payment.get("total_amount_orig")
        if payment.get("total_amount_orig") is not None
        else payment.get("amount_orig")
    )
    parent_currency = _currency(payment.get("currency"))
    original_comparable = (
        original_complete
        and parent_orig is not None
        and bool(parent_currency)
        and all(
            _currency(order.get("currency") or parent_currency) == parent_currency
            for order in orders
        )
    )

    if local_complete and parent_local is not None:
        basis = f"{source}_local"
        chosen = local_cents
        parent_cents = parent_local
    elif original_comparable:
        basis = f"{source}_original"
        chosen = original_cents
        parent_cents = parent_orig
    else:
        audit["status"] = "unresolved"
        audit["comparison_basis"] = "unavailable"
        audit["error_code"] = "E_PARENT_WRITEOFF_MISMATCH"
        audit["reason"] = (
            "整笔回款的订单已核销金额无完整可比口径"
            if source == "order_written_off"
            else "整笔回款缺少完整可比的订单交付额，无法兜底"
        )
        return [], audit

    logical: List[dict] = []
    for order, amount_cents in zip(orders, chosen):
        assert amount_cents is not None
        original_amount = _money(_cents(order.get(original_key)))
        local_amount = _money(_cents(order.get(local_key)))
        amount = original_amount if original_amount is not None else _money(amount_cents)
        so = str(order.get("so") or "").strip()
        logical.append({
            "record_id": f"{source.upper()}|{payment.get('ar') or ''}|{so}",
            "rowid": "",
            "ar": str(payment.get("ar") or "").strip(),
            "date": payment.get("hexiao_date"),
            "so": so,
            "amount": amount,
            "amount_local": local_amount,
            "currency": str(order.get("currency") or payment.get("currency") or "").strip(),
            "revoked": "",
            "source": str(order.get("source") or source),
            "snapshot_date": order.get("snapshot_date"),
        })
        audit["order_records"].append({
            "so": so,
            "amount": amount,
            "amount_local": local_amount,
            "basis_amount": _money(amount_cents),
            "source": source,
        })

    total_cents = sum(value or 0 for value in chosen)
    delta = parent_cents - total_cents
    audit["comparison_basis"] = basis
    audit["logical_record_count"] = len(logical)
    audit["effective_detail_count"] = len(logical)
    audit["effective_order_amount_count"] = len(logical)
    audit["logical_total"] = _money(total_cents)
    audit["order_amount_total"] = _money(total_cents)
    audit["delta_raw"] = _money(delta)
    audit["delta_dedup"] = _money(delta)
    audit["delta"] = _money(delta)
    if abs(delta) > tolerance_cents:
        audit["status"] = "unresolved"
        audit["error_code"] = "E_PARENT_WRITEOFF_MISMATCH"
        audit["reason"] = (
            "整笔回款的父总到账与订单已核销金额合计差额超过1元"
            if source == "order_written_off"
            else "整笔回款的父总到账与订单交付额兜底合计差额超过1元"
        )
    return logical, audit


def audit_parent_writeoffs(
    payment: dict,
    raw_rows: Iterable[dict],
    *,
    tolerance_cents: int = TOLERANCE_CENTS,
) -> Tuple[List[dict], dict]:
    """返回 `(logical_rows, audit)`；整笔回款先过父 AR 金额守恒闸。"""
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

    huikuan_type = str(payment.get("huikuan_type") or "").strip()
    is_whole_payment = huikuan_type == "整笔回款"
    audit = {
        "status": "normal",
        "comparison_basis": "unavailable",
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
        "error_code": "",
        "huikuan_type": huikuan_type,
        "is_whole_payment": is_whole_payment,
        "effective_detail_count": len(physical_kept),
        "threshold": _money(tolerance_cents),
        "delta": None,
        "parent_net_orig": payment.get("amount_orig"),
        "parent_net_local": payment.get("amount_local"),
        "parent_charge_orig": payment.get("charge_amount_orig") or 0.0,
        "parent_charge_local": payment.get("charge_amount_local"),
        "parent_total_orig": (
            payment.get("total_amount_orig")
            if payment.get("total_amount_orig") is not None
            else payment.get("amount_orig")
        ),
        "parent_total_local": (
            payment.get("total_amount_local")
            if payment.get("total_amount_local") is not None
            else payment.get("amount_local")
        ),
    }

    if is_whole_payment:
        return _audit_whole_payment_orders(
            payment,
            audit,
            tolerance_cents=tolerance_cents,
        )

    if not physical_kept:
        parent_local = _cents(
            payment.get("total_amount_local")
            if payment.get("total_amount_local") is not None
            else payment.get("amount_local")
        )
        parent_orig = _cents(
            payment.get("total_amount_orig")
            if payment.get("total_amount_orig") is not None
            else payment.get("amount_orig")
        )
        parent_cents = parent_local if parent_local is not None else parent_orig
        audit["status"] = "parent_fallback"
        audit["comparison_basis"] = (
            "parent_local" if parent_local is not None
            else ("parent_original" if parent_orig is not None else "unavailable")
        )
        audit["delta_raw"] = _money(parent_cents)
        audit["delta_dedup"] = _money(parent_cents)
        audit["delta"] = _money(parent_cents)
        audit["reason"] = "没有有效逐SO本次核销金额，回退父回款总到账金额"
        return [], audit

    parent_local = _cents(
        payment.get("total_amount_local")
        if payment.get("total_amount_local") is not None
        else payment.get("amount_local")
    )
    parent_orig = _cents(
        payment.get("total_amount_orig")
        if payment.get("total_amount_orig") is not None
        else payment.get("amount_orig")
    )
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
    basis = "detail_local" if local_comparable else (
        "detail_original" if original_comparable else ""
    )
    parent_cents = parent_local if local_comparable else parent_orig

    if not basis or parent_cents is None:
        audit["status"] = "unresolved"
        audit["comparison_basis"] = "unavailable"
        audit["error_code"] = "E_SYSTEM_OVER_WRITEOFF_UNRESOLVED"
        audit["reason"] = "父回款与逐单核销无法形成统一本币或同币种原币比较口径"
        return [], audit

    def amount_cents(row: dict) -> int:
        value = row.get("amount_local") if basis == "detail_local" else row.get("amount")
        cents = _cents(value)
        assert cents is not None
        return cents

    raw_total_cents = sum(amount_cents(row) for row in physical_kept)
    delta_raw_cents = parent_cents - raw_total_cents
    audit["comparison_basis"] = basis
    audit["raw_total"] = _money(raw_total_cents)
    audit["logical_total"] = _money(raw_total_cents)
    audit["delta_raw"] = _money(delta_raw_cents)
    audit["delta_dedup"] = _money(delta_raw_cents)
    audit["delta"] = _money(delta_raw_cents)

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
        audit["error_code"] = "E_SYSTEM_OVER_WRITEOFF_UNRESOLVED"
        audit["reason"] = "明显超核销且存在缺少核销记录NUM的明细，无法安全区分系统重复"
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
        duplicate_groups.append({
            "so": str(sample.get("so") or "").strip(),
            "amount": _money(amount_cents(sample)),
            "record_ids": ids,
            "kept_record_id": kept_id,
            "ignored_record_ids": ignored_ids,
        })
    duplicate_groups.sort(key=lambda item: (item["so"], item["record_ids"]))
    audit["duplicate_groups"] = duplicate_groups

    if not duplicate_groups:
        audit["status"] = "unresolved"
        audit["error_code"] = "E_SYSTEM_OVER_WRITEOFF_UNRESOLVED"
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
    audit["effective_detail_count"] = len(dedup_rows)
    audit["logical_total"] = _money(logical_total_cents)
    audit["delta_dedup"] = _money(delta_dedup_cents)
    audit["delta"] = _money(delta_dedup_cents)

    if delta_dedup_cents < -tolerance_cents:
        audit["status"] = "unresolved"
        audit["error_code"] = "E_SYSTEM_OVER_WRITEOFF_UNRESOLVED"
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
