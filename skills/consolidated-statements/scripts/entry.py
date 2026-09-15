"""Platform subprocess protocol. Sources are immutable, results are versioned per Run."""
from pathlib import Path
from dataclasses import asdict
import argparse
import json
import re
import sys
from engine import parse_file, select_sources, classify, validate
from workbook import build_workbook, split_workbook

def emit(message,progress):
    print(json.dumps({"type":"progress","message":message,"progress":progress},ensure_ascii=False),flush=True)

def execute(request):
    period=request["parameters"]["period"]
    output=Path(request["output_dir"]);output.mkdir(parents=True,exist_ok=True)
    reports=[];issues=[];files=request.get("files",{}).get("reports",[])
    files=files if isinstance(files,list) else [files] if files else []
    emit("正在核对来源的公司、月份和字段",10)
    for index,item in enumerate(files):
        parsed,problems=parse_file(item["local_path"],period,file_name=item["name"],bs_reclassified=item["sha256"] in request.get("source_metadata",{}).get("confirmed_bs_hashes",[]))
        for report in parsed:report.source_file_id=item["file_id"]
        reports.extend(parsed);issues.extend(problems)
        emit("已检查 "+str(index+1)+" 个来源文件",10+int(30*(index+1)/max(1,len(files))))
    chosen,conflicts=select_sources(reports);issues.extend(conflicts)
    if request["parameters"].get("fetch_kingdee"):
        from collector import collect, load_xingchen_creds
        credentials = request.get("credentials") or load_xingchen_creds()
        existing={(r.company,r.kind) for r in chosen}
        downloads,problems=collect(period,output/"金蝶原始导出",credentials,existing,emit)
        issues.extend(problems)
        for path in downloads:
            parsed,problems=parse_file(path,period,bs_reclassified=True)
            reports.extend(parsed);issues.extend(problems)
        chosen,conflicts=select_sources(reports);issues.extend(conflicts)
    accepted={(r.company,r.kind) for r in chosen}
    def replaced(issue):
        key=(issue.get("company"),issue.get("kind"))
        if key not in accepted:
            return False
        if issue.get("status") in {"missing","needs_review"}:
            return True
        # Only the known ambiguous-column export is superseded by its valid replacement.
        # Unknown fields, conflicting inputs and unrelated rejected files remain visible.
        return (issue.get("message")=="本月或本年累计列不明确" and
                any((r.company,r.kind)==key and r.file_name==issue.get("file") for r in chosen))
    resolved_issues=[i for i in issues if replaced(i)]
    issues=[i for i in issues if not replaced(i)]
    emit("正在执行分类、汇总和校验",65)
    snapshots=[]
    for r in chosen:
        item={"company":r.company,"period":r.period,"kind":r.kind,"file_name":r.file_name,"file_hash":r.file_hash,"sheet":r.sheet,"source_file_id":r.source_file_id,"template":r.template,"metric_basis":r.metric_basis,"bs_reclassified":r.kind=="bs" and not any("重分类" in issue for issue in r.issues),"issues":r.issues,"raw_values":[],"classified_values":[],"checks":validate(classify(r))}
        for key,src in [("raw_values",r),("classified_values",classify(r))]:
            item[key]=[{"item":i,"metric":m,"amount":str(v.amount),"state":v.state,"refs":v.refs} for (i,m),v in sorted(src.values.items())]
        snapshots.append(item)
    path=output/("甲骨易"+period[:4]+"年"+period[4:]+"月合并报表底稿.xlsx")
    result=build_workbook(chosen,period,path,input_issues=issues)
    if not result["coverage_complete"]:
        renamed=path.with_stem(path.stem+"_部分")
        path.rename(renamed);path=renamed
    elif not result["complete"]:
        renamed=path.with_stem(path.stem+"_待核实")
        path.rename(renamed);path=renamed
    split_paths=split_workbook(path)
    audit_path=output/"报表来源与核验记录.json"
    audit_path.write_text(json.dumps({"period":period,"reports":snapshots,"issues":issues,"resolved_input_issues":resolved_issues,**result},ensure_ascii=False,indent=2),encoding="utf-8")
    emit("已生成六张主表及来源核验记录" if result["complete"] else "已生成部分底稿或待核实底稿，请检查范围与提示",95)
    output_files=[{"name":path.name,"path":str(path)},{"name":audit_path.name,"path":str(audit_path)}]
    output_files.extend({"name":p.name,"path":str(p)} for p in split_paths)
    for original in sorted((output/"金蝶原始导出").glob("*.xlsx")) if (output/"金蝶原始导出").exists() else []:
        output_files.append({"name":original.name,"path":str(original)})
    for diagnostic in sorted((output/"取数诊断").glob("*")):
        if diagnostic.suffix in {".png",".json"}:
            output_files.append({"name":diagnostic.name,"path":str(diagnostic)})
    warnings=[x["message"] for x in issues]
    if not result["complete"]:warnings.insert(0,"本次为部分范围底稿或存在待核实项目，详见报表来源与核验记录。")
    return {"status":"success","summary":{"complete":result["complete"],"coverage_complete":result["coverage_complete"],"missing_report_count":24-len(chosen),"report_count":result["report_count"],"formula_count":result["formula_count"],"period":period},"coverage":result["coverage"],"issues":issues,"output_files":output_files,"warnings":warnings}

def request_from_folder(period, input_dir, out_dir, fetch_kingdee):
    files = []
    if input_dir:
        folder = Path(input_dir)
        if not folder.is_dir():
            raise SystemExit(f"ask=材料夹不存在：{folder}")
        for path in sorted(folder.glob("*.xlsx")) + sorted(folder.glob("*.xls")):
            files.append({
                "local_path": str(path.resolve()),
                "name": path.name,
                "sha256": "",
                "file_id": path.name,
            })
    output = Path(out_dir) if out_dir else Path.cwd() / f"合并报表_{period}"
    return {
        "parameters": {"period": period, "fetch_kingdee": bool(fetch_kingdee)},
        "files": {"reports": files},
        "output_dir": str(output.resolve()),
        "source_metadata": {"confirmed_bs_hashes": []},
    }


if __name__=="__main__":
    parser=argparse.ArgumentParser(description="八主体合并报表")
    parser.add_argument("--request")
    parser.add_argument("--result")
    parser.add_argument("--period", help="会计月份 YYYYMM")
    parser.add_argument("--input-dir", help="人放的公司原始报表夹")
    parser.add_argument("--out-dir")
    parser.add_argument("--fetch-kingdee", action="store_true")
    args=parser.parse_args()
    try:
        if args.request:
            payload=json.loads(Path(args.request).read_text(encoding="utf-8"))
            result=execute(payload)
            if args.result:
                Path(args.result).write_text(json.dumps(result,ensure_ascii=False),encoding="utf-8")
        else:
            if not args.period:
                raise SystemExit("ask=请给 --period YYYYMM")
            if not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", args.period):
                raise SystemExit("ask=期间必须是 YYYYMM")
            result=execute(request_from_folder(args.period, args.input_dir, args.out_dir, args.fetch_kingdee))
            print(json.dumps({
                "status": result.get("status"),
                "summary": result.get("summary"),
                "warnings": result.get("warnings"),
                "output_files": result.get("output_files"),
            }, ensure_ascii=False, indent=2))
    except Exception as exc:
        print("consolidation_failed:"+type(exc).__name__,file=sys.stderr)
        sys.exit(1)
