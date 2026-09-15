#!/usr/bin/env python3
"""Verify actual rows and every workbook part against the fixed approved patch."""
from __future__ import annotations

import argparse
import collections
import copy
import hashlib
import json
import shutil
import zipfile
from pathlib import Path

import apply_to_copy
import common
import workbook_finalize


def digest(path: Path) -> str:
    with path.open("rb") as handle:
        return hashlib.file_digest(handle, "sha256").hexdigest()


def compare_parts(expected: Path, actual: Path) -> int:
    with zipfile.ZipFile(expected) as authorized, zipfile.ZipFile(actual) as written:
        expected_names, actual_names = authorized.namelist(), written.namelist()
        if len(set(actual_names)) != len(actual_names) or set(expected_names) != set(actual_names):
            raise ValueError("实际工作簿部件集合与已校验计划的预期结果不一致")
        changed = [name for name in expected_names if authorized.read(name) != written.read(name)]
        if changed:
            raise ValueError("实际工作簿存在计划不能解释的部件差异：" + "、".join(changed[:10]))
        return len(expected_names)


def verify_package(before: Path, actual: Path, writes: list[dict], expected: Path) -> dict:
    """Reconstruct only in the isolated review directory, never on live material.

    Start from the immutable baseline and the approved plan, not the writer's
    receipt or actual workbook. The existing patcher supplies the permitted
    insertion/formula translations. Compare every ZIP entry, including rows of
    the same SO that were not targeted, sheets, external links and pivot parts.
    verify_written separately checks the actual values and physical row mapping.
    """
    before_hash, actual_hash = digest(before), digest(actual)
    if not writes:
        if before_hash != actual_hash:
            raise ValueError("无写入计划的年度工作簿发生变化")
        return {"baseline_sha256": before_hash, "actual_sha256": actual_hash,
                "mode": "unchanged_file"}
    expected.parent.mkdir(parents=True, exist_ok=True)
    _, patch = apply_to_copy.write_plan(before, expected, copy.deepcopy(writes), return_patch_result=True)
    workbook_finalize.finalize_workbook(expected, {"明细": patch})
    parts_checked = compare_parts(expected, actual)
    if digest(before) != before_hash or digest(actual) != actual_hash:
        raise ValueError("完整工作簿核对期间文件发生变化")
    return {"baseline_sha256": before_hash, "actual_sha256": actual_hash,
            "expected_sha256": digest(expected), "parts_checked": parts_checked,
            "mode": "approved_patch_all_parts"}


def verify_flow(workspace: Path, baseline: Path, actual: Path, checked: dict) -> dict:
    import apply_flow
    import build_flow_plan

    relative = actual.relative_to(workspace)
    before = (baseline / relative).resolve(strict=True)
    if not before.is_relative_to(baseline):
        raise ValueError("流转写前材料超出固定工作区")
    before_hash, actual_hash = digest(before), digest(actual)
    result = json.loads((workspace / "04_产出" / "流转阶段执行结果.json").read_text(encoding="utf-8"))
    if not result.get("flow_written"):
        if not result.get("restored_baseline") or before_hash != actual_hash:
            raise ValueError("未完成的流转没有恢复到已验证原副本")
        return {"mode": "restored_baseline", "baseline_sha256": before_hash, "actual_sha256": actual_hash}
    proof = workspace / "execution-review" / "flow-protection"
    expected = proof / relative
    expected.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(before, expected)
    flow = json.loads((workspace / "04_产出" / "流转写入计划_校验后.json").read_text(encoding="utf-8"))
    for phase in ("prefill", "status"):
        phase_plan = flow if phase == "prefill" else build_flow_plan.finalize_plan_after_ledger(flow, checked, workspace=proof)
        items = phase_plan.get("items") or []
        if phase == "status" and (result.get("manual_items") != phase_plan.get("manual_items")
                                  or result.get("manual_count") != len(phase_plan["manual_items"])):
            raise ValueError("流转人工项目与固定基线的逐笔判定不一致")
        for item in items:
            if item.get("verdict") == "write":
                resolved = apply_flow._resolve_flow_path(proof, item.get("file") or "")
                if resolved is None or resolved.resolve() != expected.resolve():
                    raise ValueError("流转计划包含当前固定流转文件之外的目标")
        changes, problems = apply_flow.write_flow_items(proof, items, in_place=True, phase=phase)
        recorded = result.get("phases", {}).get(phase) or {}
        if problems or recorded.get("state") != "verified" or recorded.get("changed_count") != len(changes):
            raise ValueError("流转实际执行记录与固定计划不能相互核实")
    parts_checked = compare_parts(expected, actual)
    if digest(before) != before_hash or digest(actual) != actual_hash:
        raise ValueError("流转完整性核对期间材料发生变化")
    return {"mode": "approved_patch_all_parts", "parts_checked": parts_checked,
            "baseline_sha256": before_hash, "actual_sha256": actual_hash}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="按实际写后行号复核，保护未写订单")
    parser.add_argument("--workspace", required=True)
    parser.add_argument("--baseline", required=True)
    parser.add_argument("--checked", required=True)
    parser.add_argument("--flow-file", required=True)
    parser.add_argument("--ledger-year", action="append", default=[])
    args = parser.parse_args(argv)
    workspace = Path(args.workspace).resolve(strict=True)
    baseline = Path(args.baseline).resolve(strict=True)
    checked = Path(args.checked).resolve(strict=True)
    flow_file = Path(args.flow_file).resolve(strict=True)
    if not checked.is_relative_to(workspace) or not flow_file.is_relative_to(workspace):
        raise ValueError("复核计划不属于暂存工作区")
    plan = json.loads(checked.read_text(encoding="utf-8"))
    day = common.norm_date(plan.get("hexiao_date"))
    if day is None:
        raise ValueError("计划核销日期无效")
    tag = day.strftime("%Y%m%d")
    if args.ledger_year:
        ledgers = common.parse_year_ledger_specs(args.ledger_year)
        if len(set(ledgers.values())) != len(ledgers):
            raise ValueError("多个年度引用同一盈亏文件，不能核对年度写入")
        before_ledgers = {}
        for year, path in ledgers.items():
            if path.parent != workspace / "02_我的表副本":
                raise ValueError("年度盈亏引用超出暂存副本目录")
            before = (baseline / path.relative_to(workspace)).resolve(strict=True)
            if before.parent != baseline / "02_我的表副本" or not before.is_file():
                raise ValueError("固定年度缺少对应的原始盈亏副本")
            before_ledgers[year] = before
        for root, paths in ((workspace, ledgers), (baseline, before_ledgers)):
            actual = {path.resolve() for path in (root / "02_我的表副本").glob("*盈亏*.xls*")
                      if path.is_file() and not path.name.startswith(("~$", ".")) and "便携版" not in path.stem}
            if actual != set(paths.values()):
                raise ValueError("固定年度映射未完整覆盖写前或写后盈亏文件")
    else:
        # Existing standalone callers retain their original discovery behavior.
        ledgers = common.discover_year_ledgers(workspace)
        before_ledgers = common.discover_year_ledgers(baseline)
    if set(ledgers) != set(before_ledgers):
        raise ValueError("写前与写后年度工作簿集合不一致")
    if any(int(item.get("ledger_year") or 0) not in ledgers for item in plan.get("write") or []):
        raise ValueError("写入计划引用了未提供的年度工作簿")
    problems = []
    verified_cases = []
    execution_rows = {}
    protection = {}
    for year, path in ledgers.items():
        writes = [item for item in plan.get("write", []) if int(item.get("ledger_year") or 0) == year]
        if writes:
            receipt_path = workspace / "04_产出" / f"盈亏执行_{tag}_{year}.json"
            receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
            actual = receipt.get("items") or []
            if (not receipt.get("verified") or common.norm_date(receipt.get("hexiao_date")) != day
                    or collections.Counter(item.get("case_id") for item in actual) != collections.Counter(item.get("case_id") for item in writes)):
                raise ValueError("盈亏执行凭据未覆盖当前年度的全部写入记录")
            problems.extend(apply_to_copy.verify_written(path, actual))
            verified_cases.extend(item.get("case_id") for item in actual)
            for item in actual:
                key = str(item.get("case_id") or "")
                if not key or key in execution_rows:
                    raise ValueError("实际写入案例标识重复或缺失，不能映射首次与最终行")
                execution_rows[key] = {
                    "ledger_year": year, "sheet": "明细",
                    "initial_row": item.get("ledger_row_ref"),
                    "applied_row": item.get("_applied_row_ref"),
                    "inserted_row": item.get("_inserted_row_ref"),
                    "unpaid_row": item.get("_chain_unpaid_row_ref"),
                    "row_operation": item.get("row_operation") or {},
                }
        before = before_ledgers.get(year)
        if before is None:
            raise ValueError("缺少写前对应年度材料，不能证明非目标订单保持原样")
        protection[str(year)] = verify_package(
            before, path, writes,
            workspace / "execution-review" / "protection" / f"expected_{year}.xlsx",
        )
    if problems:
        raise ValueError("写后业务核对失败：" + "；".join(problems[:20]))
    flow_protection = verify_flow(workspace, baseline, flow_file, plan)
    result = {"schema_version": "ar-write-verification-v1", "verified": True,
              "reconciliation_date": day.isoformat(), "verified_case_ids": verified_cases,
              "execution_rows": execution_rows,
              "protected_orders_checked": True, "workbook_protection": protection,
              "flow_protection": flow_protection}
    (workspace / "04_产出" / "写后业务核对.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8",
    )
    # Classification may use the allocation ledger; prepare only the isolated
    # review copy with the already verified execution, never the published ledger.
    review = workspace / "execution-review"
    if review.is_dir():
        import fallback_allocation_ledger
        fallback_allocation_ledger.commit(review, plan)
    print(json.dumps({"verified": True, "written_count": len(verified_cases)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
