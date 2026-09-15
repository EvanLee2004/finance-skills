#!/usr/bin/env python3
"""Run the flow stage after verified ledger writes, retaining both phase outcomes."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import apply_flow
import build_flow_plan
import common


def run(workspace: Path, checked: Path) -> dict:
    plan = json.loads(checked.read_text(encoding="utf-8"))
    date = common.norm_date(plan.get("hexiao_date"))
    if date is None:
        raise ValueError("校验后计划缺少核销日期")
    tag = date.strftime("%Y%m%d")
    source = workspace / "04_产出" / "流转写入计划_校验后.json"
    result = {"schema_version": "ar-flow-result-v1", "flow_written": False, "phases": {}}
    try:
        flow = json.loads(source.read_text(encoding="utf-8"))
        if not isinstance(flow, dict) or not isinstance(flow.get("items"), list):
            raise ValueError("流转计划无效")
    except (OSError, ValueError) as exc:
        result["phases"]["plan"] = {"state": "failed", "error_type": type(exc).__name__, "reason": str(exc)}
        return result
    result["manual_count"] = sum(item.get("verdict") == "hand" for item in flow["items"])
    for phase in ("prefill", "status"):
        try:
            phase_plan = flow if phase == "prefill" else build_flow_plan.finalize_plan_after_ledger(flow, plan, workspace=workspace)
            if phase == "status":
                result["manual_items"] = phase_plan["manual_items"]
                result["manual_count"] = len(result["manual_items"])
                (workspace / "04_产出" / f"流转状态回填计划_{tag}.json").write_text(
                    json.dumps(phase_plan, ensure_ascii=False, indent=2), encoding="utf-8",
                )
            items = phase_plan.get("items") or []
            changes, problems = apply_flow.write_flow_items(workspace, items, in_place=True, phase=phase)
            report = workspace / "04_产出" / f"流转{'前置' if phase == 'prefill' else '状态'}变更清单_{tag}.xlsx"
            apply_flow.write_change_report(changes, report)
            eligible = sum(item.get("verdict") == "write" for item in items)
            monthly_prefill = phase == "prefill" and any(item.get("monthly_schema") for item in items)
            if monthly_prefill:
                eligible = 0
            result["phases"][phase] = {
                "state": "failed" if problems else "verified",
                "applicable": not monthly_prefill,
                "reason": "月度核销在状态阶段统一登记，前置阶段不适用" if monthly_prefill else "",
                "eligible_count": eligible,
                "changed_count": None if problems else len(changes),
                "unchanged_count": None if problems else eligible - len(changes),
                "problems": problems,
                "changes": changes,
            }
            if problems:
                break
        except Exception as exc:
            result["phases"][phase] = {"state": "failed", "error_type": type(exc).__name__, "reason": str(exc)}
            break
    result["flow_written"] = result["phases"].get("status", {}).get("state") == "verified"
    return result


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="盈亏回读后的流转登记及状态回填")
    parser.add_argument("--workspace", required=True)
    parser.add_argument("--checked", required=True)
    args = parser.parse_args(argv)
    workspace = Path(args.workspace).resolve(strict=True)
    checked = Path(args.checked).resolve(strict=True)
    if not checked.is_relative_to(workspace):
        raise ValueError("计划超出当前工作区")
    result = run(workspace, checked)
    target = workspace / "04_产出" / "流转阶段执行结果.json"
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print("AR_FLOW_RESULT " + json.dumps({"flow_written": result["flow_written"]}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
