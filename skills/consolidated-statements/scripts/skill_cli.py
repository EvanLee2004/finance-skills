"""Machine-callable facade. stdout is one JSON object; progress goes to stderr."""
import argparse
import contextlib
import hashlib
import json
import os
from pathlib import Path
import re
import sys

VERSION = "2026.09.21.1"

def check(request, evidence):
    from eliminations import validate_evidence
    issues = []
    period = request.get("parameters", {}).get("period", "")
    if not isinstance(period, str) or not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", period):
        issues.append("invalid_period")
    else:
        try:
            validate_evidence(evidence, period)
        except (ValueError, KeyError, TypeError):
            issues.append("invalid_evidence")
    if "credentials" in request:
        issues.append("credentials_in_request")
    output_value = request.get("output_dir")
    output = Path(output_value).resolve() if isinstance(output_value, str) and output_value else None
    if output is None:
        issues.append("missing_output_dir")
    elif output.exists() and (not output.is_dir() or any(output.iterdir())):
        issues.append("output_not_empty_directory")
    files = request.get("files", {}).get("reports", [])
    if not isinstance(files, list):
        issues.append("reports_must_be_array")
        files = []
    seen = set()
    for i, item in enumerate(files):
        if not isinstance(item, dict) or not isinstance(item.get("local_path"), str):
            issues.append("invalid_report:" + str(i))
            continue
        source = Path(item["local_path"]).resolve()
        if source in seen:
            issues.append("duplicate_source:" + str(i))
        seen.add(source)
        if not source.is_file():
            issues.append("source_missing:" + str(i))
            continue
        if output and (output == source.parent or output in source.parents):
            issues.append("source_output_overlap:" + str(i))
        digest = hashlib.sha256(source.read_bytes()).hexdigest()
        if item.get("sha256") and item["sha256"] != digest:
            issues.append("source_hash_mismatch:" + str(i))
    fetch = request.get("parameters", {}).get("fetch_kingdee", False)
    if not isinstance(fetch, bool):
        issues.append("fetch_kingdee_must_be_boolean")
    if fetch and not (os.environ.get("HTTPS_PROXY") and os.environ.get("FINANCIAL_NETWORK_TARGETS")):
        issues.append("network_policy_missing")
    if not files and not fetch:
        issues.append("no_report_sources")
    return dict(ready=not issues, issues=issues,
                note="Preflight validates structure and files only; report coverage, accounting checks and browser login are verified during execution.")

def dispatch(action, request=None, evidence=None):
    if action == "describe":
        return dict(schema_version="1", skill="consolidated-statements", version=VERSION,
                    status="ok", commands=["describe", "check", "run"],
                    cash_flow_method="out_both", amount_authority="current_kingdee",
                    credentials="stdin only for authorized platform collector",
                    outputs=["six_sheet_workpaper", "three_consolidated_sheets",
                             "three_parent_sheets", "audit_json"])
    readiness = check(request, evidence)
    envelope = dict(schema_version="1", skill="consolidated-statements", version=VERSION)
    if not readiness["ready"]:
        return dict(envelope, status="needs_input", preflight=readiness)
    if action == "check":
        return dict(envelope, status="ready", preflight=readiness)
    from entry import run
    with contextlib.redirect_stdout(sys.stderr):
        result = run(request, evidence)
    complete = bool(result.get("summary", {}).get("complete")) and bool(result.get("summary", {}).get("elimination_complete"))
    return dict(envelope, status="complete" if complete else "partial", result=result)

def main():
    parser = argparse.ArgumentParser(description="AI JSON interface for consolidated-statements")
    parser.add_argument("action", choices=["describe", "check", "run"])
    parser.add_argument("--request")
    parser.add_argument("--evidence")
    args = parser.parse_args()
    if args.action != "describe" and not (args.request and args.evidence):
        parser.error("check/run require --request and --evidence")
    try:
        request = json.loads(Path(args.request).read_text("utf-8-sig")) if args.request else None
        evidence = json.loads(Path(args.evidence).read_text("utf-8-sig")) if args.evidence else None
        result = dispatch(args.action, request, evidence)
    except Exception as exc:
        # Do not echo exception text, source contents, credentials or environment.
        result = dict(schema_version="1", skill="consolidated-statements", version=VERSION,
                      status="failed", error_code=type(exc).__name__,
                      message="Execution failed. Inspect controlled audit records; do not retry into the same output directory.")
    print(json.dumps(result, ensure_ascii=False))
    return {"needs_input": 2, "failed": 1}.get(result["status"], 0)

if __name__ == "__main__":
    raise SystemExit(main())
