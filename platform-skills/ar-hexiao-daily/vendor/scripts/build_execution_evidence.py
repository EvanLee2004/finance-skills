#!/usr/bin/env python3
"""Build immutable, task-local evidence without changing a business workbook."""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path

import common
import validate_plan

EVIDENCE_VERSION = "ar-evidence-v1"
ITEM_FIELDS = (
    "ar", "so", "sod", "case_id", "code", "reason", "ledger_year",
    "ledger_row_ref", "delivery_date", "delivery_date_issue", "match_basis",
    "five_cols", "current_values", "warning_codes", "idempotence_audit",
    "write_currency_audit", "split_payment_source", "sod_capacity_audit",
    "ambiguous_sod_waterfall",
    "row_operation", "same_so_multi_sod_absorbed", "tail_tolerance_absorbed",
    "split_chain_group_id", "split_chain_index", "split_chain_count",
    "source_lineage",
)
PARENT_FIELDS = (
    "ar", "status", "comparison_basis", "parent_total_orig", "parent_total_local",
    "parent_net_orig", "parent_net_local", "parent_charge_orig", "parent_charge_local",
    "order_amount_total", "delta", "delta_dedup", "threshold", "reason",
    "error_code", "fallback_used", "is_whole_payment", "unallocated_parent_amount",
)


def digest(path: Path) -> str:
    with path.open("rb") as handle:
        return hashlib.file_digest(handle, "sha256").hexdigest()


def json_value(value):
    if isinstance(value, (dt.date, dt.datetime)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(item) for item in value]
    return value


def owned_file(workspace: Path, path: Path) -> Path:
    resolved = path.resolve()
    if path.is_symlink() or not resolved.is_relative_to(workspace) or not resolved.is_file():
        raise ValueError("证据来源不存在或超出当前工作区")
    return resolved


def record_identity(item: dict) -> str:
    """Row positions and classification outcomes are deliberately not identity."""
    event = (item.get("split_payment_source") or {}).get("writeoff_sequence_key")
    logical = [item.get(key) or "" for key in ("case_id", "ar", "so", "sod")]
    payload = json.dumps([logical, event], ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def build(workspace: Path, date: str) -> dict:
    day = dt.date.fromisoformat(date)
    tag = day.strftime("%Y%m%d")
    result_path = owned_file(workspace, workspace / "04_产出" / f"判定结果_{tag}.json")
    initial_digest = digest(result_path)
    result = json.loads(result_path.read_text(encoding="utf-8"))
    if common.norm_date(result.get("hexiao_date")) != day:
        raise ValueError("判定结果日期与任务日期不一致")
    sources = {result_path.relative_to(workspace).as_posix(): initial_digest}
    for path in sorted((workspace / "01_智云导出").glob(f"*{tag}*")):
        if path.is_file():
            checked = owned_file(workspace, path)
            sources[checked.relative_to(workspace).as_posix()] = digest(checked)

    ledgers = {}
    orders = {}
    parents = {}
    records = []
    identities = set()
    for bucket in ("auto", "hold", "exception"):
        items = result.get(bucket) or []
        if not isinstance(items, list):
            raise ValueError("判定结果记录列表无效")
        for item in items:
            identity = record_identity(item)
            if identity in identities:
                raise ValueError("判定记录身份重复，不能按订单合并后隐藏重复事件")
            identities.add(identity)
            record = {key: json_value(item[key]) for key in ITEM_FIELDS if key in item}
            record.update(record_id=identity, initial_bucket=bucket)
            ar = str(item.get("ar") or "")
            so = str(item.get("so") or "")
            audit = item.get("duplicate_writeoff_audit") or {}
            record["source_events"] = [
                json_value({key: event.get(key) for key in (
                    "record_id", "rowid", "ar", "so", "date", "amount", "amount_local",
                    "currency", "source", "snapshot_date", "disposition", "reason",
                )})
                for event in audit.get("raw_records", audit.get("records", []))
                if isinstance(event, dict) and event.get("so") == so
            ]
            if ar:
                parents[ar] = {key: json_value(audit[key]) for key in PARENT_FIELDS if key in audit}
            raw_path = str(item.get("ledger_path") or "")
            year = item.get("ledger_year")
            order_key = f"{year or '-'}|{so or '-'}"
            record["order_key"] = order_key
            if order_key not in orders:
                orders[order_key] = {"so": so, "ledger_year": year, "rows": [], "available": False}
            if raw_path and so and Path(raw_path).is_file():
                path = owned_file(workspace, Path(raw_path))
                relative = path.relative_to(workspace).as_posix()
                if relative not in ledgers:
                    before = digest(path)
                    rows = validate_plan.read_ledger_rows(path)
                    if digest(path) != before:
                        raise ValueError("读取证据期间盈亏材料发生变化")
                    sources[relative] = before
                    by_so = {}
                    for row_number, values in rows.items():
                        if "应收金额" not in values and "应收" in values:
                            values = {**values, "应收金额": values["应收"]}
                        row_so = str(values.get("SO") or "").strip()
                        if row_so:
                            by_so.setdefault(row_so, []).append({"row": row_number, **json_value(values)})
                    ledgers[relative] = by_so
                orders[order_key] = {
                    "so": so, "ledger_year": year, "available": True,
                    "source": relative, "sha256": sources[relative], "sheet": "明细",
                    "rows": ledgers[relative].get(so, []),
                }
            records.append(record)
    # Preserve sibling context when several SOs share a parent payment.
    for ar, parent in parents.items():
        parent["record_ids"] = [record["record_id"] for record in records if record.get("ar") == ar]
        parent["related_orders"] = [
            {"record_id": record["record_id"], "so": record.get("so"), "sod": record.get("sod"),
             "amounts": record.get("split_payment_source") or {}}
            for record in records if record.get("ar") == ar
        ]
    for relative, expected in sources.items():
        if digest(owned_file(workspace, workspace / relative)) != expected:
            raise ValueError("构建证据期间输入发生变化，请重新检查材料")
    counts = {bucket: sum(record["initial_bucket"] == bucket for record in records)
              for bucket in ("auto", "hold", "exception")}
    if isinstance(result.get("counts"), dict) and result["counts"].get("total") != len(records):
        raise ValueError("判定总数与逐单证据覆盖不一致")
    return {
        "schema_version": EVIDENCE_VERSION, "reconciliation_date": date,
        "sources": sources, "counts": {**counts, "total": len(records)},
        "records": records, "orders": orders, "parents": parents,
        "interpretation": "首次判定证据；不表示已经写入或完成最终复核",
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="构建逐单判定及盈亏证据，不写业务表")
    parser.add_argument("--workspace", required=True)
    parser.add_argument("--hexiao-date", required=True)
    args = parser.parse_args(argv)
    workspace = Path(args.workspace).resolve(strict=True)
    payload = build(workspace, args.hexiao_date)
    target = workspace / "04_产出" / f"逐单证据_{args.hexiao_date.replace('-', '')}.json"
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, indent=2) + "\n"
    if target.exists():
        if target.is_symlink() or target.read_text(encoding="utf-8") != raw:
            raise ValueError("已有逐单证据与当前输入不一致，拒绝覆盖首次证据")
    else:
        with target.open("x", encoding="utf-8", newline="\n") as handle:
            handle.write(raw)
    print(json.dumps({"schema_version": EVIDENCE_VERSION, "counts": payload["counts"],
                      "sha256": digest(target)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
