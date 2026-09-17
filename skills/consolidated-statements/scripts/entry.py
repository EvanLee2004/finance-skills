"""One entrypoint for standard reports, evidence-based elimination and split outputs."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import base_entry
import workbook
from eliminations import apply_eliminations, validate_evidence, verify_formulas

def run(request, evidence):
    period = request.get("parameters", {}).get("period", "")
    if not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", period):
        raise ValueError("期间必须是YYYYMM")
    validate_evidence(evidence, period)
    output = Path(request["output_dir"]).resolve()
    if output.exists() and any(output.iterdir()):
        raise ValueError("交付目录必须为空；请使用新的任务目录，避免重复抵销或覆盖")
    if "credentials" in request:
        raise ValueError("凭据不能放在请求文件中")
    source_hashes = {}
    files = request.get("files", {}).get("reports", [])
    if not isinstance(files, list):
        files = [files]
    for item in files:
        source = Path(item["local_path"]).resolve()
        if output == source.parent or output in source.parents:
            raise ValueError("输入与输出目录必须分离")
        digest = hashlib.sha256(source.read_bytes()).hexdigest()
        if item.get("sha256") and item["sha256"] != digest:
            raise ValueError("来源文件哈希变化")
        item["sha256"] = digest
        source_hashes[source] = digest
    original_build = base_entry.build_workbook
    original_verify = workbook.verify_formulas
    def build(reports, period, path, **kwargs):
        result = original_build(reports, period, path, **kwargs)
        return apply_eliminations(path, reports, period, evidence, result)
    base_entry.build_workbook = build
    workbook.verify_formulas = verify_formulas
    try:
        result = base_entry.execute(request)
    finally:
        base_entry.build_workbook = original_build
        workbook.verify_formulas = original_verify
    for path, digest in source_hashes.items():
        if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
            raise ValueError("原始文件被修改")
    audit = json.loads((output / "报表来源与核验记录.json").read_text("utf-8"))
    result["summary"]["elimination_complete"] = audit["elimination_complete"]
    result["summary"]["rules_version"] = audit["rules_version"]
    result["warnings"].extend(str(x) for x in audit["elimination_issues"])
    return result

def main():
    p = argparse.ArgumentParser(description="甲骨易六表汇总、抵销及分册")
    p.add_argument("--request", help="平台协议JSON，files.reports含local_path/name/file_id/sha256")
    p.add_argument("--period", help="YYYYMM")
    p.add_argument("--input-dir", help="本次已选单体报表目录")
    p.add_argument("--out-dir", help="新的空交付目录")
    p.add_argument("--evidence", required=True, help="本次金蝶抵销证据JSON")
    p.add_argument("--confirmed-bs-hashes", help="已确认重分类口径的来源SHA256数组JSON")
    p.add_argument("--fetch-kingdee", action="store_true", help="仅在平台已提供出站策略时启用；凭据通过stdin")
    p.add_argument("--result", help="结果JSON的目标路径")
    args = p.parse_args()
    if args.request:
        request = json.loads(Path(args.request).read_text("utf-8"))
    else:
        if not args.period or not args.out_dir:
            p.error("需要--period和--out-dir，或--request")
        files = []
        if args.input_dir:
            folder = Path(args.input_dir)
            if not folder.is_dir():
                p.error("输入目录不存在")
            for path in sorted(folder.iterdir()):
                if path.suffix.lower() in {".xls", ".xlsx"} and not path.name.startswith("~$"):
                    files.append(dict(local_path=str(path.resolve()), name=path.name, file_id=path.name, sha256=""))
        confirmed = json.loads(Path(args.confirmed_bs_hashes).read_text("utf-8")) if args.confirmed_bs_hashes else []
        request = dict(parameters=dict(period=args.period, fetch_kingdee=args.fetch_kingdee),
                       files=dict(reports=files), output_dir=str(Path(args.out_dir).resolve()),
                       source_metadata=dict(confirmed_bs_hashes=confirmed))
    evidence = json.loads(Path(args.evidence).read_text("utf-8"))
    result = run(request, evidence)
    if args.result:
        Path(args.result).write_text(json.dumps(result, ensure_ascii=False, indent=2), "utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("consolidation_failed:" + type(exc).__name__, file=sys.stderr)
        raise SystemExit(1)
