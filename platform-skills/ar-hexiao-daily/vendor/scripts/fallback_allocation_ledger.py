"""Persist successful no-detail parent allocations for cross-parent continuation."""

from __future__ import annotations

import datetime as dt
import json
import math
from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple

import fallback_sequence as FS
import amount_policy
import baseline_receipts


LEDGER_NAME = "父回款顺序分配台账.json"
VERSION = 2


def _norm_date(value) -> Optional[dt.date]:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value or "").strip()[:10]
    if not text:
        return None
    try:
        return dt.date.fromisoformat(text)
    except ValueError:
        return None


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
    except (OSError, ValueError, TypeError) as exc:
        raise ValueError("父回款顺序分配台账无法读取，不能按未分配重新核销；请恢复对应工作簿的台账。") from exc
    if not isinstance(data, dict) or not isinstance(data.get("parents"), dict):
        raise ValueError("父回款顺序分配台账结构无效，不能按空历史继续核销。")
    if int(data.get("version") or 0) not in {1, VERSION}:
        raise ValueError(f"不支持的父回款顺序分配台账版本：{data.get('version')!r}")
    for ar, entry in data["parents"].items():
        if not isinstance(entry, dict) or not isinstance(entry.get("allocations"), list):
            raise ValueError(f"父回款 {ar} 的分配记录结构无效，禁止按空历史执行")
        if "applied_cases" in entry:
            cases = entry["applied_cases"]
            if not isinstance(cases, dict) or not isinstance(entry.get("applied_sos"), list):
                raise ValueError(f"父回款 {ar} 的已写记录结构无效，不能核实历史")
            for key, case in cases.items():
                if not isinstance(case, dict):
                    raise ValueError(f"父回款 {ar} 的已写记录结构无效")
                amount = case.get("amount_local")
                if (not isinstance(amount, (int, float)) or not math.isfinite(amount) or amount <= 0
                        or key != f"{ar}|{case.get('so')}|{case.get('sod') or ''}"):
                    raise ValueError(f"父回款 {ar} 的已写金额或订单身份无效，不能核实历史")
    data.setdefault("parents", {})
    baseline_receipts.validate_journal(data.get("baseline_receipts", {}))
    return data


def unexplained_receipts(ar: str, so: str, received: float, known_history: float,
                         *, reused_allocation: bool = False, current_applied: float = 0.0) -> str:
    """Missing allocation provenance must not turn a paid workbook into fresh money."""
    if received <= known_history + float(amount_policy.TECHNICAL_EPSILON):
        return ""
    if reused_allocation:
        return (
            f"父回款 {ar} 的 SO={so} 尚有待处理分配，但原分配账面基准已变化："
            f"表内已收 {received:.2f}，原分配历史 {known_history - current_applied:.2f}，"
            f"本父回款已写 {current_applied:.2f}，比原基准增加 {received - known_history:.2f}。"
            "暂停本父回款待写部分，不能沿用旧累计拆行；请核对分配后新增回款或工作簿变更。"
        )
    return (
        f"父回款 {ar} 的历史分配归属无法完整核实；SO={so} 表内已收 {received:.2f}，"
        f"已确认归属的回款 {known_history:.2f}，其中 {received - known_history:.2f} 的父回款归属未确认。"
        "本父回款全部关联订单暂停分配；请恢复与当前盈亏表对应的分配台账或核销明细，"
        "不得把父总额重新分给未结账订单。"
    )


def _successful_pairs(checked: dict) -> set[Tuple[str, str]]:
    pairs = set()
    for bucket in ("write", "skip"):
        for item in checked.get(bucket) or []:
            if item.get("code") == "OK_SO_ALREADY_SETTLED":
                continue  # 整单业务状态不能证明这个父回款的正金额已写入。
            ar = str(item.get("ar") or "").strip()
            so = str(item.get("so") or "").strip()
            if ar and so:
                pairs.add((ar, so))
    return pairs


def eligible_entries(checked: dict) -> Dict[str, dict]:
    """Keep original allocations after a verified write, including pending orders."""
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
        cases = dict(audit.get("applied_cases") or {})
        for item in [*(checked.get("write") or []), *(checked.get("skip") or [])]:
            if item.get("ar") != ar or item.get("code") == "OK_SO_ALREADY_SETTLED":
                continue
            if item.get("same_so_multi_sod_absorbed") or item.get("tail_tolerance_absorbed"):
                continue
            amount = (item.get("five_cols") or {}).get("回款明细")
            if amount is None or float(amount) <= float(amount_policy.TECHNICAL_EPSILON):
                continue
            source = item.get("split_payment_source") or {}
            so, sod = str(item.get("so") or ""), str(item.get("sod") or "")
            key = f"{ar}|{so}|{sod}"
            evidence = {
                "so": so, "sod": sod, "amount_local": round(float(amount), 2),
                "delivery_local": source.get("delivery_local"),
                "cumulative_local": source.get("cumulative_local"),
            }
            old = cases.get(key)
            if old is not None:
                if (old["so"], old["sod"], old["amount_local"]) != (so, sod, evidence["amount_local"]):
                    raise ValueError(f"父回款 {ar} 的已写 SOD={sod} 金额发生变化，禁止覆盖分配证据")
            else:
                cases[key] = evidence
        if cases:
            if any(case["so"] not in allocated for case in cases.values()):
                raise ValueError(f"父回款 {ar} 的已写记录包含原分配之外的订单，禁止登记")
            completed = []
            for row in audit.get("allocations") or []:
                amount = sum(c["amount_local"] for c in cases.values() if c["so"] == row["so"])
                if amount <= float(amount_policy.TECHNICAL_EPSILON):
                    continue
                planned = row.get("allocated_local")
                if planned is None:
                    raise ValueError(f"父回款 {ar} 的 SO={row['so']} 缺少本币分配额，不能登记已写金额")
                if amount > float(planned) + float(amount_policy.TECHNICAL_EPSILON):
                    raise ValueError(f"父回款 {ar} 的 SO={row['so']} 已写 {amount:.2f} 超过原分配 {float(planned):.2f}")
                if amount > 0 and abs(amount - float(planned)) <= float(amount_policy.TECHNICAL_EPSILON):
                    completed.append(row["so"])
            out[ar] = {**audit, "applied_cases": cases, "applied_sos": sorted(completed)}
        elif not allocated and audit.get("processing_order") and audit.get("allocations") and all(
            any(it.get("ar") == ar and it.get("so") == row["so"] and it.get("code") == FS.ZERO
                for it in checked.get("skip") or []) for row in audit["allocations"]
        ):
            out[ar] = {**audit, "applied_cases": {}, "applied_sos": []}
        elif allocated and all((ar, so) in successful for so in allocated):
            # Compatibility with old checked plans that have no case-level fields.
            out[ar] = dict(audit)
    return out


def readback_payload(entry: dict) -> dict:
    """Compare persisted allocation and execution evidence, excluding timestamps."""
    return {
        key: value
        for key, value in entry.items()
        if key not in {"applied_at", "last_verified_at", "reused_successful_allocation"}
    }


def _stable_payload(entry: dict) -> dict:
    return {key: value for key, value in readback_payload(entry).items()
            if key not in {"applied_sos", "applied_cases"}}


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
            if "applied_cases" in old:
                if not set(old["applied_cases"]).issubset(entry.get("applied_cases") or {}):
                    raise ValueError(f"父回款 {ar} 的本次计划遗漏已有已写记录，禁止丢失历史")
                for key, previous in old["applied_cases"].items():
                    if entry["applied_cases"][key] != previous:
                        raise ValueError(f"父回款 {ar} 的本次计划改变已有已写记录，禁止覆盖历史")
                old["applied_cases"] = entry["applied_cases"]
                old["applied_sos"] = entry["applied_sos"]
            old["last_verified_at"] = now
    receipts = baseline_receipts.merge_journal(data.get("baseline_receipts", {}), checked)
    baseline_receipts.validate_journal(receipts)
    if receipts or "baseline_receipts" in data:
        data["baseline_receipts"] = receipts
    path = ledger_path(workspace)
    # Version 1 readers assume every allocation was fully written. They must
    # reject partial execution records instead of counting pending money.
    if any("applied_cases" in entry for entry in parents.values()):
        data["version"] = VERSION
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path, changed


def history_totals(
    state: dict,
    *,
    current_ar: str,
    excluded_parent_ars: Iterable[str] = (),
    as_of_date=None,
) -> Tuple[Dict[str, float], Dict[str, float]]:
    """Aggregate successful fallback allocations visible at the requested date."""
    excluded = {str(value or "").strip() for value in excluded_parent_ars}
    excluded.add(str(current_ar or "").strip())
    cutoff = _norm_date(as_of_date)
    original: Dict[str, float] = {}
    local: Dict[str, float] = {}
    for ar, entry in (state.get("parents") or {}).items():
        if str(ar or "").strip() in excluded:
            continue
        entry_date = _norm_date(entry.get("hexiao_date"))
        if cutoff is not None and entry_date is not None and entry_date > cutoff:
            continue
        for row in entry.get("allocations") or []:
            so = str(row.get("so") or "").strip()
            if not so:
                continue
            value_orig = row.get("allocated_orig")
            value_local = row.get("allocated_local")
            if "applied_cases" in entry:
                paid = round(sum(case["amount_local"] for case in entry["applied_cases"].values() if case["so"] == so), 2)
                value_orig = round(float(value_orig) * paid / float(value_local), 2) if value_orig is not None and value_local else None
                value_local = paid
            if value_orig is not None:
                original[so] = round(original.get(so, 0.0) + float(value_orig), 2)
            if value_local is not None:
                local[so] = round(local.get(so, 0.0) + float(value_local), 2)
    return original, local
