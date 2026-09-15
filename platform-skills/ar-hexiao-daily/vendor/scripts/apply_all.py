#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
校验与《核销日清》生成后统一写入：流转表前置登记、盈亏明细和流转状态
按既定顺序尝试执行。流转表写入失败时只转人工处理，不阻断盈亏明细写入。

不再要求人工确认；--confirmed 仅为旧命令兼容参数。
盈亏写入失败 → 不发布本次结果；流转表任何阶段失败均不阻断盈亏写入。
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402
import validate_plan  # noqa: E402
import apply_to_copy  # noqa: E402
import apply_flow  # noqa: E402
import build_flow_plan  # noqa: E402
import build_task_reports  # noqa: E402
import verify_sources  # noqa: E402
import workbook_finalize  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _record_done(args, *, ledger_written: bool, flow_written: bool) -> None:
    """
    写完表 → 在跑批台账上把这个核销日标成「已写表·收工」。
    只有落到这一步，`batch_ledger gaps` 才不会再把这天算成没跑过。
    """
    try:
        import json

        import batch_ledger
        import fallback_allocation_ledger

        plan = json.loads(Path(args.checked).read_text(encoding="utf-8"))
        d = common.norm_date(plan.get("hexiao_date"))
        if d is None:
            print(
                "WARN: 这份计划里没有核销日期（旧版计划），跑批台账没法登记这一天。",
                file=sys.stderr,
            )
            return
        allocation_path, allocation_added = fallback_allocation_ledger.commit(
            Path(args.workspace), plan
        )
        batch_ledger.record(
            Path(args.workspace), d, "applied",
            written={"盈亏": bool(ledger_written), "流转": bool(flow_written)},
        )
        if plan.get("parent_fallback_allocations"):
            print(
                f"父回款顺序分配台账：新增 {allocation_added} 笔，已复核保存至 {allocation_path.name}"
            )
        print(f"跑批台账：{common.date_cn(d)} 已标记「已写表·收工」")
    except Exception as e:
        print(
            f"WARN: 写后台账登记失败（不影响已写入的数据）：{type(e).__name__}",
            file=sys.stderr,
        )
    finally:
        # Backend uses this small, non-sensitive marker to distinguish a
        # successful ledger publication from a flow-table manual fallback.
        print(
            "AR_WRITE_RESULT "
            + json.dumps(
                {
                    "ledger_written": bool(ledger_written),
                    "flow_written": bool(flow_written),
                },
                ensure_ascii=False,
                separators=(",", ":"),
            )
        )


def _resnapshot_sources(workspace) -> None:
    """统一写入全部成功并登记台账后，把合法新状态登记为下一轮源文件基线。"""
    try:
        import verify_sources

        ws = common.resolve_workspace(workspace, quiet=True)
        verify_sources.do_snapshot(ws)
    except Exception as e:
        print(
            f"WARN: 最终源文件指纹刷新失败（不影响已验证的写入）：{type(e).__name__}",
            file=sys.stderr,
        )


def _merge_annual_reports(parts, target: Path, label: str) -> None:
    """把同一核销日的各年度内部报告合成一份正式报告。"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = "年度表"
    summary.append(["交付年度", "目标盈亏表", "内部报告"])
    for cell in summary[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for year, ledger, report in parts:
        summary.append([year, Path(ledger).name, Path(report).name])
        source = openpyxl.load_workbook(report, data_only=False, read_only=False)
        for index, sheet in enumerate(source.worksheets, 1):
            title = f"{year}_{index}_{sheet.title}"[:31]
            build_task_reports._copy_sheet(sheet, wb.create_sheet(title))
        source.close()
    summary.freeze_panes = "A2"
    summary.auto_filter.ref = summary.dimensions
    summary.column_dimensions["A"].width = 12
    summary.column_dimensions["B"].width = 34
    summary.column_dimensions["C"].width = 38
    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target)
    workbook_finalize.finalize_static_report(target)
    print(f"{label}（多年度合并）→ {target}")


def _year_subplan(plan: dict, year: int, path: Path) -> dict:
    check = (plan.get("ledger_checks") or {}).get(str(year)) or {}
    sub = {
        **plan,
        "write": [x for x in (plan.get("write") or []) if x.get("ledger_year") is not None and int(x["ledger_year"]) == year],
        "skip": [x for x in (plan.get("skip") or []) if x.get("ledger_year") is not None and int(x["ledger_year"]) == year],
        "conflict": [x for x in (plan.get("conflict") or []) if x.get("ledger_year") is not None and int(x["ledger_year"]) == year],
        "ledger_path": str(path),
        "ledger_sha256": check.get("sha256") or "",
    }
    sub["counts"] = {key: len(sub[key]) for key in ("write", "skip", "conflict")}
    return sub


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="日清后直接写入：盈亏 → 流转")
    ap.add_argument("--checked", required=True, help="盈亏 写入计划_校验后.json")
    ap.add_argument("--flow-plan", default="", help="流转写入计划_校验后.json；空则跳过流转")
    ap.add_argument("--ledger", default="", help="本年度盈亏副本")
    ap.add_argument(
        "--ledger-year", action="append", default=[], metavar="YEAR=PATH",
        help="其它年度盈亏工作副本，可重复，例如 2025=...xlsx",
    )
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument(
        "--confirmed",
        action="store_true",
        help="已废弃的兼容参数；现在日清与写前校验通过后可直接写入",
    )
    ap.add_argument("--in-place", action="store_true", help="盈亏就地写")
    ap.add_argument("--flow-in-place", action="store_true", help="流转就地写")
    ap.add_argument("--force", action="store_true", help="盈亏跳过冲突只写可写")
    ap.add_argument("--ledger-only", action="store_true", help="平台分阶段执行：仅写盈亏并回读，不写流转或登记正式收工台账")
    args = ap.parse_args(argv)

    checked_path = Path(args.checked)
    if not checked_path.is_file():
        print(f"ERROR: 找不到校验后计划 {checked_path}", file=sys.stderr)
        return 2
    plan = json.loads(checked_path.read_text(encoding="utf-8"))
    conflicts = plan.get("conflict") or []
    if conflicts and not args.force:
        print(
            f"ERROR: 计划里还有 {len(conflicts)} 笔冲突没处理，所有年度盈亏表均未写。",
            file=sys.stderr,
        )
        return 2

    ws = common.resolve_workspace(args.workspace, quiet=True)
    ledger_paths = {
        int(year): Path((entry or {}).get("path") or "").resolve()
        for year, entry in (plan.get("ledger_checks") or {}).items()
        if (entry or {}).get("path")
    }
    if not ledger_paths:
        ledger_paths = {
            int(year): Path(path).resolve()
            for year, path in (plan.get("ledger_targets") or {}).items()
            if path
        }
    try:
        if args.ledger or args.ledger_year:
            ledger_paths.update(common.discover_year_ledgers(
                ws, primary=args.ledger, year_specs=args.ledger_year
            ))
        elif not ledger_paths:
            ledger_paths = common.discover_year_ledgers(ws)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    writable = plan.get("write") or []
    missing_year = [item for item in writable if item.get("ledger_year") is None]
    if missing_year:
        print(
            "ERROR: 写入计划存在未确定交付年度的项目；禁止默认写入本年度盈亏表。",
            file=sys.stderr,
        )
        return 2
    grouped = {}
    for item in writable:
        year = int(item["ledger_year"])
        grouped.setdefault(year, []).append(item)
    missing = sorted(
        year for year in grouped
        if year not in ledger_paths or not ledger_paths[year].is_file()
    )
    if missing:
        print(
            "ERROR: 缺少要写入的年度盈亏表：" + "、".join(map(str, missing)),
            file=sys.stderr,
        )
        return 2

    subplans = {
        year: _year_subplan(plan, year, ledger_paths[year])
        for year in grouped
    }
    # 整单跳过也将影响流转完成状态；即使该年度没有 write，仍须复核当前表。
    so_skips = [it for it in (plan.get("skip") or []) if it.get("code") == "OK_SO_ALREADY_SETTLED"]
    if any(it.get("ledger_year") is None for it in so_skips):
        print("ERROR: 整单跳过项缺少交付年度，请重新判定；未执行盈亏或流转写入。", file=sys.stderr)
        return 2
    for year in sorted({int(it["ledger_year"]) for it in so_skips}):
        path = ledger_paths.get(year)
        if path is None or not path.is_file():
            print(f"ERROR: 缺少 {year} 年盈亏表，无法复核整单跳过；未执行写入。", file=sys.stderr)
            return 2
        try:
            problems = validate_plan.recheck_so_skips(_year_subplan(plan, year, path), path)
        except (OSError, ValueError) as exc:
            problems = [f"无法读取当前年度盈亏表（{type(exc).__name__}）"]
        if problems:
            print("ERROR: 整单跳过的执行前复核失败：" + "；".join(problems), file=sys.stderr)
            return 2

    # 所有年度表先统一做写前复核，任一失败都不开始写。
    for year, subplan in subplans.items():
        try:
            stale = apply_to_copy.precheck_before_write(
                subplan, subplan["write"], ledger_paths[year]
            )
        except ValueError as e:
            stale = [str(e)]
        if stale:
            print(
                f"ERROR: {year} 年盈亏表写入前复核没过，所有年度盈亏表均未写。",
                file=sys.stderr,
            )
            for item in stale[:10]:
                print(f"  - {item}", file=sys.stderr)
            return 2

    date_value = common.norm_date(plan.get("hexiao_date"))
    date_tag = date_value.strftime("%Y%m%d") if date_value else "未定日期"

    # 先读取流转计划，实际流转表写入放到盈亏写入并回读之后；计划异常只转人工处理。
    flow_plan_path = Path(args.flow_plan) if args.flow_plan else ws / "04_产出" / "流转写入计划_校验后.json"
    flow_plan_data = None
    flow_warnings = []
    if flow_plan_path.is_file():
        try:
            flow_plan_data = json.loads(flow_plan_path.read_text(encoding="utf-8"))
            if not isinstance(flow_plan_data, dict):
                raise ValueError("流转写入计划不是对象")
        except Exception as e:
            flow_warnings.append(
                f"流转写入计划读取未完成（{type(e).__name__}），已继续写入盈亏。"
            )
            flow_plan_data = None
    else:
        flow_warnings.append("未找到流转写入计划，已跳过流转写入并继续写入盈亏。")

    if args.in_place:
        for year in grouped:
            verify_sources.register_mutable(ws, ledger_paths[year])

    report_parts = {"变更清单": [], "订单写入差异": []}
    with tempfile.TemporaryDirectory(prefix="ar-yearly-plan-") as temp_dir:
        # 兼容原有空计划流程：虽然没有单元格需要写，仍让盈亏写入器完成
        # 空计划检查，再继续流转和最终来源指纹刷新。
        if not subplans and ledger_paths:
            preferred_year = common.current_year()
            no_write_ledger = ledger_paths.get(preferred_year)
            if no_write_ledger is None:
                no_write_ledger = ledger_paths[sorted(ledger_paths)[0]]
            no_write_year = next(year for year, path in ledger_paths.items() if path == no_write_ledger)
            empty_path = Path(temp_dir) / "checked_no_write.json"
            empty_path.write_text(json.dumps(_year_subplan(plan, no_write_year, no_write_ledger), ensure_ascii=False, indent=2), encoding="utf-8")
            ledger_args = ["--checked", str(empty_path), "--ledger", str(no_write_ledger)]
            if args.in_place:
                ledger_args.append("--in-place")
            if args.force:
                ledger_args.append("--force")
            rc1 = apply_to_copy.main(ledger_args)
            if rc1 != 0:
                print(f"ERROR: 盈亏空计划检查失败 EXIT:{rc1}，不写流转。", file=sys.stderr)
                return rc1
        for year, subplan in subplans.items():
            sub_path = Path(temp_dir) / f"checked_{year}.json"
            sub_path.write_text(json.dumps(subplan, ensure_ascii=False, indent=2), encoding="utf-8")
            ledger_args = ["--checked", str(sub_path), "--ledger", str(ledger_paths[year])]
            if args.ledger_only:
                ledger_args.extend(["--execution-receipt", str(ws / "04_产出" / f"盈亏执行_{date_tag}_{year}.json")])
            if len(subplans) > 1:
                change_report = ws / "04_产出" / f"变更清单_{date_tag}_{year}.xlsx"
                diff_report = ws / "04_产出" / f"订单写入差异_{date_tag}_{year}.xlsx"
                ledger_args.extend([
                    "--report", str(change_report),
                    "--difference-report", str(diff_report),
                ])
            else:
                change_report = ws / "04_产出" / f"变更清单_{date_tag}.xlsx"
                diff_report = ws / "04_产出" / f"订单写入差异_{date_tag}.xlsx"
            if args.in_place:
                ledger_args.append("--in-place")
            if args.force:
                ledger_args.append("--force")
            rc1 = apply_to_copy.main(ledger_args)
            if rc1 != 0:
                print(
                    f"ERROR: {year} 年盈亏写入失败 EXIT:{rc1}，不写流转。",
                    file=sys.stderr,
                )
                return rc1
            report_parts["变更清单"].append((year, ledger_paths[year], change_report))
            report_parts["订单写入差异"].append((year, ledger_paths[year], diff_report))

    if len(subplans) > 1:
        _merge_annual_reports(
            report_parts["变更清单"],
            ws / "04_产出" / f"变更清单_{date_tag}.xlsx",
            "变更清单",
        )
        _merge_annual_reports(
            report_parts["订单写入差异"],
            ws / "04_产出" / f"订单写入差异_{date_tag}.xlsx",
            "订单写入差异",
        )
    apply_to_copy._mark_review_applied(checked_path)

    if args.ledger_only:
        print("AR_WRITE_RESULT " + json.dumps({
            "ledger_written": bool(writable), "flow_written": False,
            "ledger_verified": True, "publication_pending": True,
        }))
        return 0

    # 3) 盈亏全部写入并回读成功后，再按真实结果处理流转表。
    if flow_plan_data is None:
        for warning in flow_warnings:
            print(f"WARN: {warning}")
        _record_done(args, ledger_written=bool(writable), flow_written=False)
        _resnapshot_sources(args.workspace)
        return 0

    prefill_args = [
        "--plan", str(flow_plan_path),
        "--workspace", str(args.workspace),
        "--phase", "prefill",
        "--report", str(ws / "04_产出" / f"流转前置变更清单_{date_tag}.xlsx"),
    ]
    if args.flow_in_place:
        prefill_args.append("--in-place")
    try:
        rc0 = apply_flow.main(prefill_args)
    except Exception as e:
        flow_warnings.append(
            f"流转表前置登记未完成（{type(e).__name__}），盈亏已写入。"
        )
        flow_plan_data = None
    else:
        if rc0 != 0:
            flow_warnings.append(
                f"流转表前置登记未完成（退出码 {rc0}），盈亏已写入。"
            )
            flow_plan_data = None
    if flow_plan_data is None:
        for warning in flow_warnings:
            print(f"WARN: {warning}")
        print("盈亏已写入并回读；流转表转人工处理。")
        _record_done(args, ledger_written=bool(writable), flow_written=False)
        _resnapshot_sources(args.workspace)
        return 0

    try:
        final_flow_plan = build_flow_plan.finalize_plan_after_ledger(flow_plan_data, plan, workspace=ws)
    except Exception as e:
        flow_warnings.append(
            f"流转状态计划生成未完成（{type(e).__name__}），盈亏已写入。"
        )
        for warning in flow_warnings:
            print(f"WARN: {warning}")
        print("盈亏已写入并回读；流转表转人工处理。")
        _record_done(args, ledger_written=bool(writable), flow_written=False)
        _resnapshot_sources(args.workspace)
        return 0
    final_flow_path = ws / "04_产出" / f"流转状态回填计划_{date_tag}.json"
    final_flow_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        final_flow_path.write_text(
            json.dumps(final_flow_plan, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        flow_warnings.append(
            f"流转状态计划保存未完成（{type(e).__name__}），盈亏已写入。"
        )
        for warning in flow_warnings:
            print(f"WARN: {warning}")
        print("盈亏已写入并回读；流转表转人工处理。")
        _record_done(args, ledger_written=bool(writable), flow_written=False)
        _resnapshot_sources(args.workspace)
        return 0
    flow_args = [
        "--plan", str(final_flow_path),
        "--workspace", str(args.workspace),
        "--phase", "status",
        "--report", str(ws / "04_产出" / f"流转状态变更清单_{date_tag}.xlsx"),
    ]
    if args.flow_in_place:
        flow_args.append("--in-place")
    flow_written = True
    try:
        rc2 = apply_flow.main(flow_args)
    except Exception as e:
        flow_written = False
        flow_warnings.append(
            f"流转状态回填未完成（{type(e).__name__}），盈亏已写入。"
        )
    else:
        if rc2 != 0:
            flow_written = False
            flow_warnings.append(
                f"流转状态回填未完成（退出码 {rc2}），盈亏已写入。"
            )
    if not flow_written:
        for warning in flow_warnings:
            print(f"WARN: {warning}")
        print("盈亏已写入并回读；流转表转人工处理。")
        _record_done(args, ledger_written=bool(writable), flow_written=False)
        _resnapshot_sources(args.workspace)
        return 0
    print("统一写入完成：盈亏 + 流转")
    _record_done(args, ledger_written=bool(writable), flow_written=True)
    _resnapshot_sources(args.workspace)
    return 0


if __name__ == "__main__":
    sys.exit(main())
