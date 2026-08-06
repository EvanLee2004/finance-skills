"""Persist successful no-detail parent allocations for cross-parent continuation."""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Dict, Iterable, Tuple

import amount_policy


LEDGER_NAME = "父回款顺序分配台账.json"
VERSION = 1


def ledger_path(workspace: Path) -> Path:
    folder = Path(workspace) / "03_台账"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / LEDGER_NAME


def load(workspace: Path) -> dict:
    path = ledger_path(workspace)
    if not path.is_file():
        return {"version": VERSION, "parents": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return {"version": VERSION, "parents": {}}
    if int(data.get("version") or 0) != VERSION:
        raise ValueError(f"不支持的父回款顺序分配台账版本：{data.get('version')!r}")
    data.setdefault("parents", {})
    return data


def _successful_pairs(checked: dict) -> set[Tuple[str, str]]:
    pairs = set()
    for bucket in ("write", "skip"):
        for item in checked.get(bucket) or []:
            ar = str(item.get("ar") or "").strip()
            so = str(item.get("so") or "").strip()
            if ar and so:
                pairs.add((ar, so))
    return pairs


def eligible_entries(checked: dict) -> Dict[str, dict]:
    """Only persist parents whose every positive allocation was written or idempotently present."""
    successful = _successful_pairs(checked)
    out: Dict[str, dict] = {}
    for ar, audit in (checked.get("parent_fallback_allocations") or {}).items():
        ar = str(ar or "").strip()
        if not ar or not isinstance(audit, dict):
            continue
        allocated = {
            str(row.get("so") or "").strip()
            for row in (audit.get("allocations") or [])
            if float(row.get("allocated") or 0.0) > float(amount_policy.TECHNICAL_EPSILON)
        }
        allocated.discard("")
        if allocated and all((ar, so) in successful for so in allocated):
            out[ar] = dict(audit)
    return out


def _stable_payload(entry: dict) -> dict:
    return {
        key: value
        for key, value in entry.items()
        if key not in {"applied_at", "last_verified_at", "reused_successful_allocation"}
    }


def commit(workspace: Path, checked: dict) -> Tuple[Path, int]:
    """Merge successful allocations after the workbook write succeeds; reruns are idempotent."""
    data = load(workspace)
    parents = data.setdefault("parents", {})
    now = dt.datetime.now().isoformat(timespec="seconds")
    changed = 0
    for ar, audit in eligible_entries(checked).items():
        entry = {
            **audit,
            "ar": ar,
            "hexiao_date": checked.get("hexiao_date") or audit.get("hexiao_date") or "",
        }
        old = parents.get(ar)
        if old is not None and _stable_payload(old) != _stable_payload(entry):
            raise ValueError(f"父回款 {ar} 已有成功分配记录，但本次分配不同，禁止覆盖")
        if old is None:
            entry["applied_at"] = now
            parents[ar] = entry
            changed += 1
        else:
            old["last_verified_at"] = now
    path = ledger_path(workspace)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path, changed


def history_totals(
    state: dict,
    *,
    current_ar: str,
    excluded_parent_ars: Iterable[str] = (),
) -> Tuple[Dict[str, float], Dict[str, float]]:
    """Aggregate prior successful fallback allocations, excluding current/detailed parents."""
    excluded = {str(value or "").strip() for value in excluded_parent_ars}
    excluded.add(str(current_ar or "").strip())
    original: Dict[str, float] = {}
    local: Dict[str, float] = {}
    for ar, entry in (state.get("parents") or {}).items():
        if str(ar or "").strip() in excluded:
            continue
        for row in entry.get("allocations") or []:
            so = str(row.get("so") or "").strip()
            if not so:
                continue
            value_orig = row.get("allocated_orig")
            value_local = row.get("allocated_local")
            if value_orig is not None:
                original[so] = round(original.get(so, 0.0) + float(value_orig), 2)
            if value_local is not None:
                local[so] = round(local.get(so, 0.0) + float(value_local), 2)
    return original, local
