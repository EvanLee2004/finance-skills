#!/usr/bin/env python3
"""Install only a platform-selected, independently checked recovery candidate.

This does not call a financial patcher. Copies replace task-owned staging files;
an interrupted installation is retained for investigation, never auto-retried.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from investigate_failed_write import InvestigationError, Snapshot, inside, workbook_names
from prepare_recovery_ledger import HASH, REQUEST_VERSION, RESULT_VERSION, _relative, _root, _write_json
from verify_execution_write import digest


def install(stage: Path, request_path: Path, request_sha256: str, candidate_sha256: str, attempt: str) -> dict:
    import re

    stage = _root(stage)
    if not HASH.fullmatch(request_sha256) or not HASH.fullmatch(candidate_sha256) or not re.fullmatch(r"[0-9a-f]{32}", attempt):
        raise InvestigationError("盈亏采用请求缺少有效指纹或尝试标识")
    snapshot = Snapshot()
    request, actual_request = snapshot.capture(inside(request_path, stage), document=True)
    if (actual_request != request_sha256 or request.get("schema_version") != REQUEST_VERSION
            or request.get("staging_workspace") != str(stage)):
        raise InvestigationError("盈亏采用请求与平台固定请求不一致")
    candidate_root = inside(stage / "execution-recovery" / attempt, stage)
    candidate, candidate_sha = snapshot.capture(candidate_root / "result.json", document=True)
    if (candidate_sha != candidate_sha256 or candidate.get("schema_version") != RESULT_VERSION
            or candidate.get("request_sha256") != request_sha256 or candidate.get("ledger_verified") is not True
            or candidate.get("new_write_count") != 0 or candidate.get("flow_adopted") is not False
            or candidate.get("authorizes_install") is not False
            or any(candidate.get(key) != request.get(key) for key in (
                "source_workflow_id", "target_workflow_id", "target_skill_hash", "reconciliation_date", "material_set_id"))):
        raise InvestigationError("盈亏采用候选与核查结果、任务或日期不一致")
    manifest, manifest_sha = snapshot.capture(inside(stage / "execution-manifest.json", stage), document=True)
    checked = inside(Path(manifest.get("checked_plan") or ""), stage)
    plan, checked_sha = snapshot.capture(checked, document=True)
    if (manifest_sha != request.get("target_manifest_sha256") or manifest.get("workflow_id") != request.get("target_workflow_id")
            or checked_sha != candidate.get("checked_plan_sha256") or checked_sha != manifest.get("staged_plan_fingerprint")):
        raise InvestigationError("新任务清单或校验计划在候选核查后变化")
    if type(candidate.get("adopted_record_count")) is not int or candidate["adopted_record_count"] != len(plan.get("write") or []):
        raise InvestigationError("候选采用记录数与新版本计划不一致")
    files = manifest["files"]
    if workbook_names(stage) != set(files):
        raise InvestigationError("采用前新暂存工作簿集合发生变化")
    for name, expected in files.items():
        if snapshot.capture(inside(stage / _relative(name), stage))[1] != expected:
            raise InvestigationError("采用前新暂存已存在改动，禁止覆盖或重复采用")
    ledgers = request["ledger_files"]
    tag = request["reconciliation_date"].replace("-", "")
    mapping = {}
    for year, name in ledgers.items():
        mapping[f"actual/{_relative(name).as_posix()}"] = _relative(name)
        mapping[f"receipts/盈亏执行_{tag}_{year}.json"] = Path("04_产出") / f"盈亏执行_{tag}_{year}.json"
    if set(candidate.get("artifacts") or {}) != set(mapping):
        raise InvestigationError("盈亏候选工作簿与年度凭据集合不完整")
    # An exclusive marker makes every interrupted attempt non-repeatable, even
    # when only its first workbook was replaced or a response was lost.
    journal = inside(candidate_root / "installation", stage)
    journal.mkdir(exist_ok=False)
    entries = []
    for name, relative in mapping.items():
        source = inside(candidate_root / name, stage)
        target = inside(stage / relative, stage)
        fact = candidate["artifacts"][name]
        payload = journal / "payload" / relative
        if (source.stat().st_size != fact["size"] or snapshot.capture(source, target=payload)[1] != fact["sha256"]):
            raise InvestigationError("候选工作簿或凭据内容已变化，未采用结果")
        if relative.parts[0] == "04_产出" and target.exists():
            raise InvestigationError("新任务已有盈亏凭据，禁止覆盖原执行事实")
        entries.append({"target": relative.as_posix(), "sha256": fact["sha256"], "size": fact["size"],
                        "before_sha256": files.get(relative.as_posix())})
    if not snapshot.verify_unchanged():
        raise InvestigationError("采用前核查输入变化，未替换工作簿")
    _write_json(journal / "prepared.json", {"request_sha256": request_sha256,
                "candidate_sha256": candidate_sha256, "entries": entries})
    for entry in entries:
        relative = Path(entry["target"])
        target, payload = inside(stage / relative, stage), journal / "payload" / relative
        if (digest(payload) != entry["sha256"] or (entry["before_sha256"] is not None
                and digest(target) != entry["before_sha256"]) or (entry["before_sha256"] is None and target.exists())):
            raise InvestigationError("采用过程中材料发生变化，已停止并保留安装记录，不能自动重复执行")
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = inside(target.with_name(f".{target.name}.{attempt}.adopting"), stage)
        with temporary.open("xb") as output, payload.open("rb") as incoming:
            copied = 0
            while chunk := incoming.read(min(1024 * 1024, entry["size"] - copied + 1)):
                copied += len(chunk)
                if copied > entry["size"]:
                    raise InvestigationError("采用副本读取超过核实大小，已停止")
                output.write(chunk)
        if digest(temporary) != entry["sha256"]:
            raise InvestigationError("采用临时副本与核实结果不一致，已停止")
        temporary.replace(target)
        if digest(target) != entry["sha256"]:
            raise InvestigationError("采用后工作簿或凭据回读不一致，禁止继续执行")
    final_files = {name: digest(inside(stage / _relative(name), stage)) for name in files}
    expected_files = {**files, **{name: request["source_files"][name]["sha256"] for name in ledgers.values()}}
    if workbook_names(stage) != set(files) or final_files != expected_files:
        raise InvestigationError("采用后的完整工作簿集合与预期不一致，禁止继续或发布")
    if any(digest(inside(stage / entry["target"], stage)) != entry["sha256"] for entry in entries):
        raise InvestigationError("采用完成登记前工作簿或凭据发生变化，禁止继续或发布")
    result = {"ledger_verified": True, "ledger_written": bool(plan.get("write")), "publication_pending": True,
              "origin": "adopted", "source_workflow_id": request["source_workflow_id"],
              "request_sha256": request_sha256, "candidate_sha256": candidate_sha256,
              "new_write_count": 0, "adopted_record_count": candidate["adopted_record_count"], "files": final_files}
    _write_json(journal / "completed.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="采用已独立核实的盈亏结果，不重新执行财务写入")
    for key in ("workspace", "request", "request-sha256", "candidate-sha256", "attempt"):
        parser.add_argument(f"--{key}", required=True)
    args = parser.parse_args()
    try:
        result = install(Path(args.workspace), Path(args.request), args.request_sha256, args.candidate_sha256, args.attempt)
    except Exception as exc:
        print(json.dumps({"adopted": False, "reason": str(exc) if isinstance(exc, InvestigationError)
                          else "盈亏采用未完成，保留候选、暂存及安装记录；不得自动重写或重复采用。"}, ensure_ascii=False))
        return 2
    print("AR_WRITE_RESULT " + json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
