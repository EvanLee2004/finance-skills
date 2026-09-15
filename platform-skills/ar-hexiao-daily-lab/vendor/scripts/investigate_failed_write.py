#!/usr/bin/env python3
"""Independently investigate a failed write in a fresh, private evidence copy.

The platform supplies the pinned manifest hash and task identity. Reconstruction
uses the existing deterministic patchers only on new expected files. Nothing in
this report completes an original phase or authorizes publication/re-execution.
"""
from __future__ import annotations

import argparse
import contextlib
import copy
import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path

import apply_flow
import apply_to_copy
import build_flow_plan
import common
import workbook_finalize
from verify_execution_write import digest

VERSION = "ar-write-investigation-v1"
INPUT_VERSION = "ar-investigation-inputs-v1"
MAX_JSON = 16 * 1024 * 1024
MAX_BYTES = 256 * 1024 * 1024
MAX_FILES = 32
MAX_PARTS = 10000
MAX_PART_BYTES = 128 * 1024 * 1024
MAX_PACKAGE_BYTES = 512 * 1024 * 1024
MAX_UNPACKED_INPUT_BYTES = 1024 * 1024 * 1024


class InvestigationError(ValueError):
    """Safe, fixed explanation of an investigation input failure."""


def package_size(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        parts = archive.infolist()
        if len(parts) > MAX_PARTS or len({part.filename for part in parts}) != len(parts):
            raise InvestigationError("调查工作簿部件重复或数量超过上限")
        size = sum(part.file_size for part in parts)
        if size > MAX_PACKAGE_BYTES or any(part.file_size > MAX_PART_BYTES for part in parts):
            raise InvestigationError("调查工作簿解压大小超过上限")
        return size


def compare_parts(expected: Path, actual: Path) -> int:
    package_size(expected)
    package_size(actual)
    with zipfile.ZipFile(expected) as left, zipfile.ZipFile(actual) as right:
        names = left.namelist()
        if set(names) != set(right.namelist()):
            raise ValueError("工作簿部件集合不一致")
        for name in names:
            with left.open(name) as authorized, right.open(name) as observed:
                size = 0
                while True:
                    wanted, found = authorized.read(1024 * 1024), observed.read(1024 * 1024)
                    size += max(len(wanted), len(found))
                    if size > MAX_PART_BYTES:
                        raise InvestigationError("调查工作簿单部件读取超过上限")
                    if wanted != found:
                        raise ValueError("工作簿部件与预期不一致")
                    if not wanted:
                        break
        return len(names)


def private_call(destination: Path, function, *args):
    with (destination / "reconstruction.log").open("a", encoding="utf-8") as log:
        with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
            return function(*args)


def workbook_names(root: Path) -> set[str]:
    folder = inside(root / "02_我的表副本", root)
    names = set()
    for path in folder.glob("*.xls*"):
        if path.name.startswith(".") or "便携版" in path.name:
            continue
        inside(path, root)
        if not path.is_file():
            raise InvestigationError("工作簿集合含非普通文件")
        names.add(path.relative_to(root).as_posix())
        if len(names) > MAX_FILES:
            raise InvestigationError("工作簿集合超过调查上限")
    return names


def inside(raw: Path, root: Path) -> Path:
    path = Path(raw)
    if not path.is_absolute() or not path.is_relative_to(root) or ".." in path.parts:
        raise InvestigationError("调查输入或输出超出固定工作区")
    for component in (path, *path.parents):
        if component.is_symlink():
            raise InvestigationError("调查路径含符号链接")
        if component == root:
            break
    if not path.resolve().is_relative_to(root):
        raise InvestigationError("调查路径超出固定工作区")
    return path


class Snapshot:
    def __init__(self):
        self.remaining = MAX_BYTES
        self.sources = {}

    @staticmethod
    def stamp(path: Path) -> tuple:
        stat = path.stat()
        return stat.st_dev, stat.st_ino, stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns

    def capture(self, path: Path, *, target: Path | None = None, document: bool = False):
        if not path.is_file() or path.is_symlink():
            raise InvestigationError("调查输入文件缺失或不是普通文件")
        limit = min(self.remaining, MAX_JSON if document else self.remaining)
        if path.stat().st_size > limit:
            raise InvestigationError("调查输入超过读取上限")
        before_stat = self.stamp(path)
        digest_value, chunks, size = hashlib.sha256(), [], 0
        if target is not None:
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                raise InvestigationError("调查副本目标已存在，不能覆盖")
        from contextlib import nullcontext

        with path.open("rb") as source, (target.open("xb") if target is not None else nullcontext()) as output:
            while chunk := source.read(min(1024 * 1024, limit - size + 1)):
                size += len(chunk)
                if size > limit:
                    raise InvestigationError("调查输入超过读取上限")
                digest_value.update(chunk)
                if output is not None:
                    output.write(chunk)
                if document:
                    chunks.append(chunk)
        before = digest_value.hexdigest()
        if self.stamp(path) != before_stat or (target is not None and digest(target) != before):
            raise InvestigationError("调查副本与读取时输入不一致")
        value = None
        if document:
            value = json.loads(b"".join(chunks))
            if not isinstance(value, dict):
                raise InvestigationError("调查文档不是有效对象")
        if path in self.sources and self.sources[path] != (before, size):
            raise InvestigationError("调查读取期间输入发生变化")
        self.remaining -= size
        self.sources[path] = (before, size)
        return value, before

    def verify_unchanged(self) -> bool:
        try:
            for path, (expected, expected_size) in self.sources.items():
                if not path.is_file() or path.is_symlink() or path.stat().st_size != expected_size:
                    return False
                actual, size = hashlib.sha256(), 0
                with path.open("rb") as handle:
                    while chunk := handle.read(min(1024 * 1024, expected_size - size + 1)):
                        size += len(chunk)
                        if size > expected_size:
                            return False
                        actual.update(chunk)
                if size != expected_size or actual.hexdigest() != expected:
                    return False
            return True
        except OSError:
            return False


def ledger_result(before: Path, actual: Path, writes: list[dict], expected: Path) -> dict:
    result = {"verified": False, "planned_count": len(writes), "rows_verified": False,
              "package_verified": False, "baseline_sha256": digest(before), "actual_sha256": digest(actual)}
    if not writes:
        same = result["baseline_sha256"] == result["actual_sha256"]
        return {**result, "verified": same, "rows_verified": same, "package_verified": same,
                "state": "unchanged_no_write" if same else "unexpected_change",
                "reason": "固定计划没有待写记录，文件与原副本一致。" if same else "固定计划没有待写记录，文件却发生变化。"}
    # Derive actual row addresses from the approved insertion/translation rules,
    # not from possibly missing or partially saved writer receipts.
    mapped = copy.deepcopy(writes)
    expected.parent.mkdir(parents=True, exist_ok=True)
    _, patch = apply_to_copy.write_plan(before, expected, mapped, return_patch_result=True)
    workbook_finalize.finalize_workbook(expected, {"明细": patch})
    result["expected_sha256"] = digest(expected)
    result["execution_rows"] = [{"case_id": item.get("case_id"), "initial_row": item.get("ledger_row_ref"),
                                 "applied_row": item.get("_applied_row_ref"), "inserted_row": item.get("_inserted_row_ref"),
                                 "unpaid_row": item.get("_chain_unpaid_row_ref")} for item in mapped]
    problems = apply_to_copy.verify_written(actual, mapped)
    # Details stay in the private investigation report, never stdout/exception.
    result["row_problems"] = problems
    result["rows_verified"] = not problems
    try:
        result["parts_checked"] = compare_parts(expected, actual)
        result["package_verified"] = True
    except InvestigationError:
        raise
    except ValueError:
        result["package_reason"] = "完整工作簿部件与固定计划推导的预期结果不一致。"
    result["verified"] = result["rows_verified"] and result["package_verified"]
    result["state"] = "matches_expected" if result["verified"] else "differs_from_expected"
    result["reason"] = ("实际单元格及完整工作簿与固定计划的预期结果一致；原执行与发布状态仍须平台核查。"
                        if result["verified"] else "实际单元格或工作簿部件与预期不一致，不能认定本年度写入完整。")
    return result


def flow_result(before: Path, actual: Path, relative: Path, plan: dict, flow_plan: dict,
                proof: Path, phase: str) -> dict:
    same = digest(before) == digest(actual)
    result = {"verified": False, "baseline_sha256": digest(before), "actual_sha256": digest(actual)}
    if phase == "write_ledger":
        return {**result, "verified": same, "state": "not_started_unchanged" if same else "unexpected_change",
                "reason": "尚未进入流转阶段，原流转文件保持不变。" if same else "尚未进入流转阶段，流转文件已出现改动，需调查。"}
    expected = proof / relative
    expected.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(before, expected)
    phases = {}
    for name in ("prefill", "status"):
        approved = copy.deepcopy(flow_plan) if name == "prefill" else build_flow_plan.finalize_plan_after_ledger(copy.deepcopy(flow_plan), plan)
        items = approved.get("items")
        if not isinstance(items, list):
            raise InvestigationError("固定流转计划缺少完整条目")
        for item in items:
            if item.get("verdict") == "write":
                target = apply_flow._resolve_flow_path(proof, item.get("file") or "")
                if target is None or target.resolve() != expected.resolve():
                    raise InvestigationError("流转计划引用固定到账流转表之外的目标")
        changes, problems = apply_flow.write_flow_items(proof, items, in_place=True, phase=name)
        phases[name] = {"expected_changed_count": len(changes), "problems": problems}
        if problems:
            return {**result, "state": "expected_plan_unresolved", "phases": phases,
                    "reason": "固定流转计划在隔离副本中仍有不能执行的条目，无法构造完整预期结果。"}
    result["phases"] = phases
    result["expected_sha256"] = digest(expected)
    try:
        result["parts_checked"] = compare_parts(expected, actual)
        return {**result, "verified": True, "state": "matches_expected",
                "reason": "流转文件全部部件与固定计划的预期结果一致；不据此补记原动作完成或发布。"}
    except InvestigationError:
        raise
    except ValueError:
        return {**result, "state": "baseline_retained" if same else "differs_from_expected",
                "reason": "当前流转保留原副本，未体现完整计划结果；不能推断原执行从未写入。" if same else "流转文件与完整预期结果不一致，可能未完成或存在额外改动。"}


def investigate(baseline: Path, workspace: Path, checked: Path, flow: Path, *,
                manifest_sha256: str, workflow_id: str, phase: str, attempt: str,
                ledger_years: dict[int, Path]) -> Path:
    if not re.fullmatch(r"[0-9a-f]{32}", attempt) or phase not in {"write_ledger", "write_receipt_flow"}:
        raise InvestigationError("调查尝试标识或失败阶段无效")
    if not re.fullmatch(r"[0-9a-f]{64}", manifest_sha256):
        raise InvestigationError("调查缺少平台固定清单指纹")
    baseline, workspace = baseline.resolve(strict=True), workspace.resolve(strict=True)
    if baseline == workspace:
        raise InvestigationError("调查必须分别提供原材料和写入暂存")
    checked, flow = inside(checked, workspace), inside(flow, workspace)
    snapshot = Snapshot()
    manifest, actual_manifest_sha = snapshot.capture(inside(workspace / "execution-manifest.json", workspace), document=True)
    if (actual_manifest_sha != manifest_sha256 or manifest.get("workflow_id") != workflow_id
            or manifest.get("schema_version") != "ar-execution-v2"):
        raise InvestigationError("调查清单与平台固定任务或指纹不一致")
    plan, plan_sha = snapshot.capture(checked, document=True)
    if plan_sha != manifest.get("staged_plan_fingerprint") or Path(manifest.get("checked_plan", "")).resolve() != checked:
        raise InvestigationError("调查计划与原暂存清单不一致")
    day = common.norm_date(plan.get("hexiao_date"))
    if day is None:
        raise InvestigationError("调查计划缺少有效核销日期")
    files = manifest.get("files")
    protected = manifest.get("protected_inputs")
    if (not isinstance(files, dict) or not 1 <= len(files) <= MAX_FILES
            or not isinstance(protected, dict) or not 1 <= len(protected) <= MAX_FILES):
        raise InvestigationError("调查清单缺少完整文件集合或超过数量上限")
    required = {checked.relative_to(workspace).as_posix(), "04_产出/流转写入计划_校验后.json",
                f"04_产出/判定结果_{day.strftime('%Y%m%d')}.json", f"04_产出/首次核销日清_{day.strftime('%Y%m%d')}.xlsx"}
    if not required.issubset(protected):
        raise InvestigationError("调查清单缺少必要的受保护首次结果")
    directory_stamps = {}
    for root in (baseline, workspace):
        if workbook_names(root) != set(files):
            raise InvestigationError("实际工作簿集合与原清单不完整对应")
        directory_stamps[root] = Snapshot.stamp(root / "02_我的表副本")
    # All writes from this point are in a fresh directory, never the input areas.
    destination = inside(workspace / "execution-investigations" / attempt, workspace)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.mkdir(exist_ok=False)
    before_root, actual_root = destination / "baseline", destination / "actual"
    for relative, expected_sha in files.items():
        relative_path = Path(relative)
        if relative_path.is_absolute() or relative_path.parent != Path("02_我的表副本"):
            raise InvestigationError("调查工作簿清单包含非业务副本路径")
        before = inside(baseline / relative, baseline)
        actual = inside(workspace / relative, workspace)
        if snapshot.capture(before, target=before_root / relative)[1] != expected_sha:
            raise InvestigationError("原工作簿与写前清单不一致")
        snapshot.capture(actual, target=actual_root / relative)
    flow_plan = None
    for relative, expected_sha in protected.items():
        source = inside(workspace / relative, workspace)
        is_flow_plan = relative == "04_产出/流转写入计划_校验后.json"
        document, fingerprint = snapshot.capture(source, target=inside(actual_root / relative, actual_root),
                                                 document=is_flow_plan)
        if fingerprint != expected_sha:
            raise InvestigationError("首次结果或计划与固定清单不一致")
        if is_flow_plan:
            flow_plan = document
    flow_relative = flow.relative_to(workspace)
    if flow_relative.as_posix() not in files:
        raise InvestigationError("到账流转表不属于固定工作簿集合")
    unpacked = 0
    invalid_packages = {}
    for root in (before_root, actual_root):
        for relative in files:
            try:
                unpacked += package_size(root / relative)
            except (zipfile.BadZipFile, OSError, InvestigationError) as exc:
                invalid_packages[root / relative] = {
                    "verified": False, "state": "invalid_package", "error_type": type(exc).__name__,
                    "reason": str(exc) if isinstance(exc, InvestigationError) else "工作簿损坏或无法读取，未尝试解析；其他文件单独调查。"}
            if unpacked > MAX_UNPACKED_INPUT_BYTES:
                raise InvestigationError("调查输入的解压总量超过上限")
    if not ledger_years or len(ledger_years) > MAX_FILES:
        raise InvestigationError("调查必须使用原任务固定的完整年度文件对应关系")
    ledgers = {}
    for year, path in ledger_years.items():
        relative = inside(path, baseline).relative_to(baseline)
        if (not re.fullmatch(r"20[0-9]{2}", str(year)) or relative.as_posix() not in files
                or relative == flow_relative or before_root / relative in ledgers.values()):
            raise InvestigationError("固定年度文件对应关系与原清单不一致")
        ledgers[year] = before_root / relative
    writes = plan.get("write")
    if (not isinstance(writes, list) or any(not isinstance(item, dict)
            or not isinstance(item.get("case_id"), str) or not item["case_id"]
            or int(item.get("ledger_year") or 0) not in ledgers for item in writes)
            or len({item["case_id"] for item in writes}) != len(writes)):
        raise InvestigationError("校验计划引用缺失年度")
    results = {}
    for year, before in sorted(ledgers.items()):
        actual = actual_root / before.relative_to(before_root)
        if before in invalid_packages or actual in invalid_packages:
            results[str(year)] = invalid_packages.get(before) or invalid_packages[actual]
            continue
        try:
            results[str(year)] = private_call(destination, ledger_result, before, actual,
                [item for item in writes if int(item["ledger_year"]) == year],
                destination / "expected-ledgers" / f"{year}.xlsx")
        except Exception as exc:
            results[str(year)] = {"verified": False, "state": "unconfirmed", "error_type": type(exc).__name__,
                                  "reason": str(exc) if isinstance(exc, InvestigationError)
                                      else "本年度的预期构造或实际回读未完成，其他年度结果单独保留。"}
    try:
        flow_review = invalid_packages.get(before_root / flow_relative) or invalid_packages.get(actual_root / flow_relative)
        if flow_review is None:
            flow_review = private_call(destination, flow_result, before_root / flow_relative, actual_root / flow_relative, flow_relative,
                                      plan, flow_plan, destination / "expected-flow", phase)
    except Exception as exc:
        flow_review = {"verified": False, "state": "unconfirmed", "error_type": type(exc).__name__,
                       "reason": str(exc) if isinstance(exc, InvestigationError)
                           else "流转预期构造或完整文件核对未完成，盈亏结果单独保留。"}
    reviewed = {path.relative_to(before_root).as_posix() for path in ledgers.values()} | {flow_relative.as_posix()}
    other_results = {relative: invalid_packages.get(before_root / relative) or invalid_packages.get(actual_root / relative)
                    or {"verified": digest(before_root / relative) == digest(actual_root / relative),
                               "reason": "核对非目标工作簿是否与写前文件完全一致。"}
                     for relative in files if relative not in reviewed}
    try:
        unchanged = snapshot.verify_unchanged() and all(
            workbook_names(root) == set(files) and Snapshot.stamp(root / "02_我的表副本") == stamp
            for root, stamp in directory_stamps.items())
    except (OSError, ValueError):
        unchanged = False
    input_files = {"baseline": {}, "staging": {}}
    for path, (fingerprint, size) in snapshot.sources.items():
        scope, root = ("staging", workspace) if path.is_relative_to(workspace) else ("baseline", baseline)
        if not path.is_relative_to(root):
            raise InvestigationError("调查输入指纹包含固定工作区之外的文件")
        input_files[scope][path.relative_to(root).as_posix()] = {"sha256": fingerprint, "size": size}
    payload = {"schema_version": VERSION, "workflow_id": workflow_id, "attempt": attempt,
               "input_contract_version": INPUT_VERSION, "input_files": input_files,
               "ledger_files": {str(year): path.relative_to(before_root).as_posix() for year, path in ledgers.items()},
               "flow_file": flow_relative.as_posix(),
               "failed_phase": phase, "reconciliation_date": day.isoformat(),
               "manifest_sha256": manifest_sha256, "plan_sha256": plan_sha,
               "material_set_id": manifest.get("material_set_id"), "ledger_results": results,
               "flow_result": flow_review, "other_workbooks": other_results, "inputs_unchanged": unchanged,
               "business_files_match_expected": unchanged and bool(results)
                   and all(item["verified"] for item in results.values()) and flow_review["verified"]
                   and all(item["verified"] for item in other_results.values()),
               "authorizes_resume": False, "publication_verified": False}
    if not unchanged:
        payload["reason"] = "调查期间原输入发生变化，隔离副本比较结果不能用于当前任务恢复。"
    target = destination / "result.json"
    with target.open("x", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, default=str)
    return target


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="失败写入的独立业务调查，仅写新建调查副本")
    for name in ("baseline", "workspace", "checked", "flow-file", "manifest-sha256", "workflow-id", "phase", "attempt"):
        parser.add_argument(f"--{name}", required=True)
    parser.add_argument("--ledger-year", action="append", required=True)
    args = parser.parse_args(argv)
    try:
        result = investigate(Path(args.baseline), Path(args.workspace), Path(args.checked), Path(args.flow_file),
                             manifest_sha256=args.manifest_sha256, workflow_id=args.workflow_id,
                             phase=args.phase, attempt=args.attempt,
                             ledger_years=common.parse_year_ledger_specs(args.ledger_year))
        print(json.dumps({"schema_version": VERSION, "report_sha256": digest(result)}))
        return 0
    except Exception as exc:
        print(json.dumps({"schema_version": VERSION, "error_type": type(exc).__name__,
                          "reason": str(exc) if isinstance(exc, InvestigationError) else "独立调查未完成；本脚本只写调查目录，未改写业务工作簿。"}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
