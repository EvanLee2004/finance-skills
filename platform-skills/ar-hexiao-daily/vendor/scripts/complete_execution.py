#!/usr/bin/env python3
"""Prepare an immutable auxiliary-ledger bundle after confirmed publication.

The platform makes this bundle authoritative with its completion checkpoint.
This script never edits the published workbooks or the source auxiliary ledger.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import shutil
from pathlib import Path

import batch_ledger
import common
import fallback_allocation_ledger
import baseline_receipts
import rescan_holds


def build(workspace: Path, checked: Path, publication: dict, attempt: str) -> Path:
    plan = json.loads(checked.read_text(encoding="utf-8"))
    day = common.norm_date(plan.get("hexiao_date"))
    if (day is None or publication.get("reconciliation_date") != day.isoformat()
            or publication.get("schema_version") != "ar-publication-v1"
            or not publication.get("material_set_id") or not publication.get("files")):
        raise ValueError("正式台账缺少可核实的日期和材料发布记录")
    if not attempt or any(char not in "0123456789abcdef-" for char in attempt):
        raise ValueError("正式台账执行标识无效")
    proof = workspace / "formal-ledger-build" / attempt
    proof.mkdir(parents=True, exist_ok=True)
    names = (fallback_allocation_ledger.LEDGER_NAME, batch_ledger.LEDGER_NAME)
    ledger_dir = proof / "03_台账"
    ledger_dir.mkdir(exist_ok=True)
    for name in names:
        source = workspace / "03_台账" / name
        if source.is_file():
            data = json.loads(source.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("既有辅助台账格式无效，禁止用空台账覆盖历史")
            required = "parents" if name == fallback_allocation_ledger.LEDGER_NAME else "runs"
            if not isinstance(data.get(required), dict):
                raise ValueError("既有辅助台账缺少历史记录集合")
            shutil.copy2(source, ledger_dir / name)
    prior_state = fallback_allocation_ledger.load(proof)
    prior_allocations = prior_state["parents"]
    allocation_path, _ = fallback_allocation_ledger.commit(proof, plan)
    expected = fallback_allocation_ledger.eligible_entries(plan)
    saved = json.loads(allocation_path.read_text(encoding="utf-8"))
    if saved.get("baseline_receipts", {}) != baseline_receipts.merge_journal(prior_state.get("baseline_receipts", {}), plan):
        raise ValueError("正式回款身份及原始应收基线回读不一致")
    for ar, entry in expected.items():
        desired = {**entry, "ar": ar, "hexiao_date": plan.get("hexiao_date") or entry.get("hexiao_date") or ""}
        prior = prior_allocations.get(ar)
        # Legacy successful entries already prove the full allocation; commit
        # preserves them rather than replacing them with a partial replay.
        if prior is not None and "applied_cases" not in prior:
            desired = prior
        if fallback_allocation_ledger.readback_payload(saved["parents"].get(ar, {})) != fallback_allocation_ledger.readback_payload(desired):
            raise ValueError("正式父回款分配台账回读不一致")
    written = publication["written"]
    batch_ledger.record(proof, day, "applied", written=written,
                        note=f"已核实材料版本 {publication['material_version']}；发布任务 {publication['workflow_id']}")
    batch = json.loads((ledger_dir / batch_ledger.LEDGER_NAME).read_text(encoding="utf-8"))
    record = batch.get("runs", {}).get(day.isoformat()) or {}
    if record.get("stage") != "applied" or record.get("written") != written:
        raise ValueError("正式跑批台账回读不一致")
    hold_path = rescan_holds.ledger_path(workspace)
    if not hold_path.is_file():
        raise ValueError("挂账重扫台账不存在，不能完成正式台账登记")
    hold_bytes = hold_path.read_bytes()
    rescan_holds.load_ledger(hold_path)
    payload = {
        "schema_version": "ar-formal-ledgers-v1", "publication": publication,
        "json_ledgers": {name: json.loads((ledger_dir / name).read_text(encoding="utf-8")) for name in names},
        "binary_ledgers": {hold_path.name: {
            "base64": base64.b64encode(hold_bytes).decode("ascii"),
            "sha256": hashlib.sha256(hold_bytes).hexdigest(),
        }},
    }
    target = proof / f"核销辅助台账_{day.strftime('%Y%m%d')}.json"
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return target


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="发布核实后生成正式辅助台账包")
    parser.add_argument("--workspace", required=True)
    parser.add_argument("--checked", required=True)
    parser.add_argument("--publication", required=True)
    parser.add_argument("--attempt", required=True)
    args = parser.parse_args(argv)
    workspace = Path(args.workspace).resolve(strict=True)
    checked, publication_path = Path(args.checked).resolve(strict=True), Path(args.publication).resolve(strict=True)
    if not checked.is_relative_to(workspace) or not publication_path.is_relative_to(workspace):
        raise ValueError("台账提交依据超出当前工作区")
    target = build(workspace, checked, json.loads(publication_path.read_text(encoding="utf-8")), args.attempt)
    print(json.dumps({"formal_ledger_bundle": str(target)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
