#!/usr/bin/env python3
"""Prepare adopted ledger evidence using the recovery task's newly checked plan.

Only fresh evidence files are written. The platform must separately recheck the
source investigation, stopped execution, permissions and material head before
installing a candidate. A successful preparation does not complete any phase.
"""
from __future__ import annotations

import argparse
import copy
import json
import re
from functools import partial
from pathlib import Path

import apply_to_copy
import common
import workbook_finalize
from investigate_failed_write import (
    InvestigationError, MAX_FILES, MAX_UNPACKED_INPUT_BYTES, Snapshot,
    compare_parts, inside, package_size, private_call, workbook_names,
)
from verify_execution_write import digest

REQUEST_VERSION = "ar-ledger-adoption-request-v1"
RESULT_VERSION = "ar-ledger-adoption-candidate-v1"
HASH = re.compile(r"[0-9a-f]{64}")


def _relative(value: object) -> Path:
    if not isinstance(value, str):
        raise InvestigationError("恢复文件引用缺失")
    path = Path(value)
    if (path.is_absolute() or path.as_posix() != value
            or path.parent != Path("02_我的表副本") or ":" in value or "\\" in value):
        raise InvestigationError("恢复文件引用超出工作簿副本目录")
    return path


def _write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2, default=str)


def _root(path: Path) -> Path:
    if not path.is_absolute() or ".." in path.parts:
        raise InvestigationError("恢复工作区路径不是固定绝对路径")
    for part in (path, *path.parents):
        if part.is_symlink() or getattr(part.lstat(), "st_file_attributes", 0) & 0x400:
            raise InvestigationError("恢复工作区包含链接或重解析点")
    return path.resolve(strict=True)


def prepare(baseline: Path, stage: Path, source: Path, request_path: Path,
            request_sha256: str, attempt: str) -> Path:
    if not HASH.fullmatch(request_sha256) or not re.fullmatch(r"[0-9a-f]{32}", attempt):
        raise InvestigationError("恢复请求指纹或尝试标识无效")
    baseline, stage, source = (_root(path) for path in (baseline, stage, source))
    if baseline == source or stage == source or stage == baseline:
        raise InvestigationError("恢复必须使用分别固定的原材料、新暂存和原执行结果")
    snapshot = Snapshot()
    request, request_sha = snapshot.capture(inside(request_path, stage), document=True)
    if request_sha != request_sha256 or request.get("schema_version") != REQUEST_VERSION:
        raise InvestigationError("恢复请求与平台固定指纹或契约不一致")
    if any(request.get(key) != str(root) for key, root in (
            ("baseline_workspace", baseline), ("staging_workspace", stage), ("source_workspace", source))):
        raise InvestigationError("恢复命令的材料目录与固定请求不一致")
    for name in ("source_skill_hash", "target_skill_hash", "investigation_sha256", "input_fingerprint"):
        if not isinstance(request.get(name), str) or not HASH.fullmatch(request[name]):
            raise InvestigationError("恢复请求缺少固定版本或调查输入指纹")
    for name in ("source_workflow_id", "target_workflow_id", "investigation_action_id", "material_set_id"):
        if not isinstance(request.get(name), str) or not request[name] or len(request[name]) > 128:
            raise InvestigationError("恢复请求缺少原任务、新任务或材料身份")
    if request["source_workflow_id"] == request["target_workflow_id"]:
        raise InvestigationError("恢复不能覆盖原任务身份")
    manifest, manifest_sha = snapshot.capture(inside(stage / "execution-manifest.json", stage), document=True)
    if (manifest_sha != request.get("target_manifest_sha256")
            or manifest.get("schema_version") != "ar-execution-v2"
            or manifest.get("workflow_id") != request["target_workflow_id"]
            or manifest.get("material_set_id") != request["material_set_id"]):
        raise InvestigationError("恢复请求与新任务暂存清单不一致")
    checked = inside(Path(manifest.get("checked_plan") or ""), stage)
    plan, plan_sha = snapshot.capture(checked, document=True)
    if plan_sha != manifest.get("staged_plan_fingerprint"):
        raise InvestigationError("新版本校验计划与暂存登记不一致")
    day = common.norm_date(plan.get("hexiao_date"))
    if day is None or request.get("reconciliation_date") != day.isoformat():
        raise InvestigationError("新版本校验计划与恢复日期不一致")
    files, observed, ledgers = manifest.get("files"), request.get("source_files"), request.get("ledger_files")
    if (not isinstance(files, dict) or not 1 <= len(files) <= MAX_FILES
            or not isinstance(observed, dict) or set(observed) != set(files)
            or not isinstance(ledgers, dict) or not 1 <= len(ledgers) < len(files)):
        raise InvestigationError("恢复请求未完整绑定工作簿及年度集合")
    names = {name: _relative(name) for name in files}
    for name, item in observed.items():
        if (not isinstance(files[name], str) or not HASH.fullmatch(files[name])
                or not isinstance(item, dict) or type(item.get("size")) is not int or item["size"] < 0
                or not isinstance(item.get("sha256"), str) or not HASH.fullmatch(item["sha256"])):
            raise InvestigationError("恢复请求的工作簿指纹或大小无效")
    if (any(not re.fullmatch(r"20[0-9]{2}", year) or not isinstance(name, str) or name not in files
            for year, name in ledgers.items()) or len(set(ledgers.values())) != len(ledgers)
            or request.get("flow_file") not in files or request["flow_file"] in ledgers.values()):
        raise InvestigationError("恢复年度或流转文件与完整材料集合不一致")
    writes = plan.get("write")
    if (not isinstance(writes, list) or any(not isinstance(item, dict)
            or not isinstance(item.get("case_id"), str) or not item["case_id"]
            or str(item.get("ledger_year")) not in ledgers for item in writes)
            or len({item["case_id"] for item in writes}) != len(writes)):
        raise InvestigationError("新版本计划的案例身份或年度映射无效")
    stamps = {}
    for root in (baseline, stage, source):
        if workbook_names(root) != set(files):
            raise InvestigationError("恢复工作簿集合与固定清单不完整对应")
        stamps[root] = Snapshot.stamp(root / "02_我的表副本")
    destination = inside(stage / "execution-recovery" / attempt, stage)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.mkdir(exist_ok=False)
    before_root, actual_root = destination / "baseline", destination / "actual"
    unpacked = 0
    for name, relative in names.items():
        before, current, previous = (inside(root / relative, root) for root in (baseline, stage, source))
        if snapshot.capture(before, target=before_root / relative)[1] != files[name]:
            raise InvestigationError("恢复原材料与新任务写前清单不一致")
        if snapshot.capture(current)[1] != files[name]:
            raise InvestigationError("新暂存已经变化，不能覆盖或重复采用旧写入")
        if previous.stat().st_size != observed[name]["size"]:
            raise InvestigationError("原执行结果在调查后大小发生变化")
        if snapshot.capture(previous, target=actual_root / relative)[1] != observed[name]["sha256"]:
            raise InvestigationError("原执行结果在调查后内容发生变化")
        for root in (before_root, actual_root):
            unpacked += package_size(root / relative)
        if unpacked > MAX_UNPACKED_INPUT_BYTES:
            raise InvestigationError("恢复材料解压总量超过核查上限")
        if name not in ledgers.values() and name != request["flow_file"] and observed[name]["sha256"] != files[name]:
            raise InvestigationError("原执行改动了非目标工作簿，不能采用盈亏结果")
    # Derive row mappings with the new Skill, never from an old writer receipt.
    results, receipts = {}, {}
    for year, name in sorted(ledgers.items()):
        before, actual = before_root / names[name], actual_root / names[name]
        mapped = copy.deepcopy([item for item in writes if str(item["ledger_year"]) == year])
        if mapped:
            expected = destination / "expected" / names[name]
            expected.parent.mkdir(parents=True, exist_ok=True)
            _, patch = private_call(destination, partial(apply_to_copy.write_plan, return_patch_result=True),
                                    before, expected, mapped)
            private_call(destination, workbook_finalize.finalize_workbook, expected, {"明细": patch})
            try:
                compare_parts(expected, actual)
            except ValueError as exc:
                raise InvestigationError("原盈亏工作簿完整部件与新版本计划不一致，不能采用旧写入结果") from exc
            if private_call(destination, apply_to_copy.verify_written, actual, mapped):
                raise InvestigationError("原盈亏单元格与新版本校验计划不一致，未生成可采用凭据")
        elif digest(before) != digest(actual):
            raise InvestigationError("新版本没有本年度写入计划，但原盈亏文件已变化")
        results[year] = {"verified": True, "planned_count": len(mapped),
                         "baseline_sha256": files[name], "actual_sha256": observed[name]["sha256"]}
        receipts[year] = {"hexiao_date": day.isoformat(), "verified": True, "items": mapped,
                          "origin": "adopted", "source_workflow_id": request["source_workflow_id"],
                          "request_sha256": request_sha, "new_write_count": 0}
    if not snapshot.verify_unchanged() or any(Snapshot.stamp(root / "02_我的表副本") != stamp for root, stamp in stamps.items()):
        raise InvestigationError("核查期间恢复输入发生变化，未登记可采用结果")
    for year, receipt in receipts.items():
        _write_json(destination / "receipts" / f"盈亏执行_{day.strftime('%Y%m%d')}_{year}.json", receipt)
    artifacts = {}
    for year, name in ledgers.items():
        actual = actual_root / names[name]
        if digest(actual) != observed[name]["sha256"]:
            raise InvestigationError("候选盈亏副本在核对期间变化，不能登记结果")
        for path in (actual, destination / "receipts" / f"盈亏执行_{day.strftime('%Y%m%d')}_{year}.json"):
            artifacts[path.relative_to(destination).as_posix()] = {"sha256": digest(path), "size": path.stat().st_size}
    result = {"schema_version": RESULT_VERSION, "request_sha256": request_sha,
              "source_workflow_id": request["source_workflow_id"], "target_workflow_id": request["target_workflow_id"],
              "reconciliation_date": day.isoformat(), "material_set_id": request["material_set_id"],
              "target_skill_hash": request["target_skill_hash"], "checked_plan_sha256": plan_sha,
              "artifacts": artifacts,
              "ledger_verified": True, "ledger_results": results, "adopted_record_count": len(writes),
              "new_write_count": 0, "flow_adopted": False, "authorizes_install": False,
              "message": "原盈亏结果符合新版本计划，已生成独立候选与凭据；尚未安装、完成阶段或发布材料。"}
    _write_json(destination / "result.json", result)
    return destination / "result.json"


def main() -> int:
    parser = argparse.ArgumentParser(description="按新版本计划核对旧盈亏结果，仅生成隔离候选")
    for argument in ("baseline", "workspace", "source", "request", "request-sha256", "attempt"):
        parser.add_argument(f"--{argument}", required=True)
    args = parser.parse_args()
    try:
        path = prepare(Path(args.baseline), Path(args.workspace), Path(args.source),
                       Path(args.request), args.request_sha256, args.attempt)
    except Exception as exc:
        print(json.dumps({"prepared": False, "reason": str(exc) if isinstance(exc, InvestigationError)
                          else "恢复候选核对未完成，保留原任务与材料；没有采用或重写盈亏结果。"}, ensure_ascii=False))
        return 2
    print(json.dumps({"prepared": True, "result_sha256": digest(path), "authorizes_install": False}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
