#!/usr/bin/env python3
"""Strict hold rescan: missing ledger capability is not reported as success."""
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

import common
import rescan_holds
from classify_hexiao import LedgerIndex
from build_execution_evidence import digest, record_identity
from execution_lineage import indexed_decisions, match_final_records


def rescan_evidence(rows: list[dict], before: dict, initial: dict, links: dict, paths: dict, baseline: dict) -> list[dict]:
    evidence = []
    for row in rows:
        key = str(row.get("案例ID") or "")
        old = before.get(key) or {}
        current = initial.get(key)
        if current and any((row.get(label) or "") != (current.get(field) or "")
                           for label, field in (("AR", "ar"), ("SO", "so"))):
            raise ValueError("挂账案例 ID 对应的 AR/SO 与本批首次判定不一致")
        year = common.to_number(row.get("交付年度"))
        if old.get("状态") == "已完成":
            notice = "本条此前已完成，本轮未重新判定；保留历史完成原因。"
        elif not row.get("SO"):
            notice = "缺少关联 SO，无法按订单查询盈亏业务行；保留待处理。"
        elif year is None:
            notice = "缺少交付年度，无法确定应核对的年度盈亏表；保留待处理。"
        elif int(year) not in paths:
            notice = f"本任务缺少 {int(year)} 年盈亏材料；未复核该项，保留待处理。"
        else:
            notice = str(row.get("原因") or "") + "；" + str(row.get("复查条件") or "")
        evidence.append({
            "case_id": key, "scope": "current_task" if current else "historical",
            "initial_record_ids": [record_identity(current)] if current else [],
            "final_case_ids": links[key] if current else [],
            "before_status": old.get("状态") or "未在台账中", "after_status": row.get("状态"),
            "before_code": old.get("E码") or "", "after_code": row.get("E码") or "",
            "before_reason": old.get("原因") or "", "after_reason": row.get("原因") or "",
            "before_source": {"path": baseline["snapshot"], "sha256": baseline["snapshot_sha256"],
                              "case_id": key, "ar": old.get("AR") or "", "so": old.get("SO") or ""} if old else {},
            "reason": notice, "revisit_condition": row.get("复查条件") or "",
            "ledger_year": int(year) if year is not None else None,
            "state_changed": old.get("状态") != row.get("状态"),
        })
    return evidence


def fixed_hold_input(workspace: Path, target: Path, day: str, sources: dict) -> tuple[list[dict], dict]:
    output = workspace / "execution-hold-input"
    if not output.resolve().is_relative_to(workspace):
        raise ValueError("挂账重扫输入目录超出任务工作区")
    output.mkdir(exist_ok=True)
    manifest = output / f"挂账重扫输入_{day}.json"
    snapshot = output / f"挂账重扫前_{day}.xlsx"
    if any(path.is_symlink() or not path.resolve().is_relative_to(workspace) for path in (target, manifest, snapshot)):
        raise ValueError("挂账重扫台账或证据路径超出受控工作区")
    if manifest.exists():
        baseline = json.loads(manifest.read_text(encoding="utf-8"))
        if (not isinstance(baseline, dict) or baseline.get("schema_version") != "ar-hold-rescan-input-v1"
                or not isinstance(baseline.get("existed"), bool)
                or baseline.get("sources") != sources
                or baseline.get("snapshot") != snapshot.relative_to(workspace).as_posix()):
            raise ValueError("挂账重扫恢复输入与固定来源不一致，不能覆盖原证据")
    else:
        if snapshot.exists():
            # An interrupted snapshot copy is not an accepted checkpoint.
            raise ValueError("挂账重扫前快照缺少输入清单，需核对原始台账后恢复")
        existed = target.is_file()
        original_sha = digest(target) if existed else ""
        if existed:
            shutil.copy2(target, snapshot)
            if digest(snapshot) != original_sha or digest(target) != original_sha:
                raise ValueError("保存重扫前快照时原台账发生变化，禁止继续合并")
        baseline = {"schema_version": "ar-hold-rescan-input-v1", "sources": sources,
                    "snapshot": snapshot.relative_to(workspace).as_posix(),
                    "snapshot_sha256": original_sha, "existed": existed,
                    "original_source": target.relative_to(workspace).as_posix()}
        manifest.write_text(json.dumps(baseline, ensure_ascii=False, indent=2), encoding="utf-8")
    if baseline.get("existed"):
        if not snapshot.is_file() or digest(snapshot) != baseline.get("snapshot_sha256"):
            raise ValueError("重扫前挂账台账快照指纹不一致，不能使用无法核验的历史状态")
        return rescan_holds.load_ledger(snapshot), baseline
    if snapshot.exists() or baseline.get("snapshot_sha256"):
        raise ValueError("原挂账台账不存在的记录与实际快照冲突")
    return [], baseline


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="对已写工作区重扫挂账并输出结构化结果")
    parser.add_argument("--workspace", required=True)
    parser.add_argument("--result", required=True)
    parser.add_argument("--ledger-year", action="append", default=[])
    args = parser.parse_args(argv)
    workspace = Path(args.workspace).resolve(strict=True)
    source = Path(args.result).resolve(strict=True)
    if not source.is_relative_to(workspace):
        raise ValueError("挂账重扫来源超出工作区")
    result = json.loads(source.read_text(encoding="utf-8"))
    day = common.norm_date(result.get("hexiao_date"))
    if day is None:
        raise ValueError("挂账重扫缺少固定核销日期")
    final_source = workspace / "execution-review" / "04_产出" / f"判定结果_{day:%Y%m%d}.json"
    final_result = json.loads(final_source.read_text(encoding="utf-8"))
    if common.norm_date(final_result.get("hexiao_date")) != day:
        raise ValueError("挂账重扫的首次与写后判定日期不一致")
    initial, final = indexed_decisions(result), indexed_decisions(final_result)
    links = match_final_records(initial, final)
    projected = {bucket: [] for bucket in ("auto", "hold", "exception")}
    for key, targets in links.items():
        candidates = [final[target] for target in targets]
        bucket = next(value for value in ("exception", "hold", "auto")
                      if any(item["bucket"] == value for item in candidates))
        representative = next(item for item in candidates if item["bucket"] == bucket)
        projected[bucket].append({**representative, "case_id": key})
    paths = common.discover_year_ledgers(workspace, year_specs=args.ledger_year)
    if not paths:
        raise ValueError("缺少盈亏材料，不能将只合并台账声明为挂账重扫完成")
    ledgers = {year: LedgerIndex(path) for year, path in paths.items()}
    target = rescan_holds.ledger_path(workspace)
    sources = {source.relative_to(workspace).as_posix(): digest(source),
               final_source.relative_to(workspace).as_posix(): digest(final_source),
               **{path.relative_to(workspace).as_posix(): digest(path) for path in paths.values()}}
    rows, baseline = fixed_hold_input(workspace, target, day.strftime("%Y%m%d"), sources)
    before = {}
    for row in rows:
        key = str(row.get("案例ID") or "")
        if not key or key in before:
            raise ValueError("挂账台账案例 ID 缺失或重复，不能覆盖历史记录")
        current = initial.get(key)
        if current and any((row.get(label) or "") != (current.get(field) or "")
                           for label, field in (("AR", "ar"), ("SO", "so"))):
            raise ValueError("历史挂账案例 ID 与本批 AR/SO 冲突，禁止先合并再覆盖身份")
        before[key] = dict(row)
    rows = rescan_holds.merge_from_classify(rows, result, day.isoformat())
    rows = rescan_holds.rescan_idempotent(rows, projected, day.isoformat())
    stats = rescan_holds.reclassify_against_ledger(rows, ledgers, day.isoformat())
    evidence = rescan_evidence(rows, before, initial, links, paths, baseline)
    rescan_holds.save_ledger(target, rows)
    payload = {"schema_version": "ar-hold-rescan-v1", "reconciliation_date": day.isoformat(),
               "rows": rows, "stats": stats, "source": source.name, "records": evidence,
               "scope_counts": {scope: sum(item["scope"] == scope for item in evidence)
                                for scope in ("current_task", "historical")},
               "sources": sources, "baseline": baseline}
    (workspace / "04_产出" / "挂账重扫结果.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8",
    )
    print(json.dumps({"rescanned": True, "count": len(rows)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
