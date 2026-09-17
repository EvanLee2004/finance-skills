#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按内容认文件夹：销项发票 / 付款 / 收款。缺材料列出人话。"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
CONFIG = SKILL / "config"

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def load_aliases() -> dict:
    p = CONFIG / "列名别名.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.is_file() else {}


def code_str(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s


def norm_name(name: str) -> str:
    s = (name or "").strip().replace("(", "（").replace(")", "）")
    return "".join(s.split())


def find_header_row(ws, alias_map: dict | None, required: list[str], max_scan: int = 20) -> tuple[int, list]:
    """前几行里找「像表头」的那一行，不假定第 1 行就是列名。"""
    best_i, best, best_score = 1, [], -1
    alias_map = alias_map or {}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        headers = list(row or [])
        cleaned = _cleaned(headers)
        score = sum(1 for key in required if field_hit(cleaned, alias_map, key))
        if score > best_score:
            best_score, best_i, best = score, i, headers
        if score >= len(required) and required:
            return i, headers
    return best_i, best


def header_names(path: Path, aliases: dict | None = None) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    aliases = aliases or load_aliases()
    try:
        wb = load_workbook(path, read_only=True, data_only=False)
    except Exception:
        return out
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            title = str(name or "")
            if "收款" in title and "付款" not in title:
                _, headers = find_header_row(ws, aliases.get("收款_列别名"), ["客户名称", "借方（增加）"])
            elif "付款" in title:
                _, headers = find_header_row(ws, aliases.get("付款_列别名"), ["供应商", "应付金额本币"])
            else:
                _, headers = find_header_row(ws, aliases.get("销项发票_列别名"), ["单位名称", "价税合计"])
            out[name] = _cleaned(headers)
    finally:
        wb.close()
    return out


def field_hit(headers: list[str], aliases: dict, field: str) -> bool:
    names = aliases.get(field) or [field]
    return any(h in names for h in headers)


def _cleaned(headers) -> list[str]:
    return [str(h).strip() for h in (headers or []) if h is not None and str(h).strip()]


def _field_index(headers, aliases: dict, field: str) -> int | None:
    names = (aliases or {}).get(field) or [field]
    compact = {"".join(str(n).split()) for n in names}
    for i, h in enumerate(headers or []):
        if h is None:
            continue
        s = str(h).strip()
        if s in names or "".join(s.split()) in compact:
            return i
    return None


def is_sales_sheet(name: str, headers: list[str], aliases: dict | None = None) -> bool:
    title = str(name or "")
    if "组织架构" in title or "流水" in title:
        return False
    if "收款" in title and "发票" not in title:
        return False
    sales_a = (aliases or {}).get("销项发票_列别名") or {}
    cleaned = _cleaned(headers)
    return field_hit(cleaned, sales_a, "单位名称") and field_hit(cleaned, sales_a, "价税合计")


def is_receipt_sheet(name: str, headers: list[str], aliases: dict | None = None) -> bool:
    title = str(name or "")
    if "流水" in title:
        return False
    rec_a = (aliases or {}).get("收款_列别名") or {}
    cleaned = _cleaned(headers)
    if any("价税合计" in h for h in cleaned):
        return False
    if any(h == "贷方（减少）" or h.startswith("贷方（减少）") for h in cleaned):
        return False
    has_cust = "客户名称" in cleaned or field_hit(cleaned, rec_a, "客户名称")
    has_debit = "借方（增加）" in cleaned or field_hit(cleaned, rec_a, "借方（增加）")
    return has_cust and has_debit


def is_payment_sheet(name: str, headers: list[str], aliases: dict | None = None) -> bool:
    title = str(name or "")
    if "流水" in title or "组织架构" in title:
        return False
    if "中行付款" in title or ("中行" in title and "付款" in title):
        return False
    if "收款" in title:
        return False
    pay_a = (aliases or {}).get("付款_列别名") or {}
    cleaned = _cleaned(headers)
    if any("价税合计" in h for h in cleaned):
        return False
    return field_hit(cleaned, pay_a, "供应商") and field_hit(cleaned, pay_a, "应付金额本币")


def workbook_has_receipt(path: Path, aliases: dict | None = None) -> bool:
    return any(is_receipt_sheet(name, headers, aliases) for name, headers in header_names(path).items())


def _pick_named(named: list[str], headered: list[str]) -> str | None:
    if named:
        return named[0]
    if headered:
        return headered[0]
    return None


def find_receipt_sheet(wb, aliases: dict | None = None) -> str | None:
    named = []
    headered = []
    for name in wb.sheetnames:
        ws = wb[name]
        _, raw = find_header_row(ws, (aliases or {}).get("收款_列别名"), ["客户名称", "借方（增加）"])
        headers = _cleaned(raw)
        if not is_receipt_sheet(name, headers, aliases):
            continue
        if "收款" in str(name) and "流水" not in str(name) and "付款" not in str(name):
            named.append(name)
        else:
            headered.append(name)
    return _pick_named(named, headered)


def find_payment_sheet(wb, aliases: dict | None = None) -> str | None:
    named = []
    headered = []
    for name in wb.sheetnames:
        ws = wb[name]
        _, raw = find_header_row(ws, (aliases or {}).get("付款_列别名"), ["供应商", "应付金额本币"])
        headers = _cleaned(raw)
        if not is_payment_sheet(name, headers, aliases):
            continue
        if "付款" in str(name) and "中行" not in str(name):
            named.append(name)
        else:
            headered.append(name)
    return _pick_named(named, headered)


def load_org_map(wb, aliases: dict | None = None) -> dict[str, str]:
    org_a = (aliases or {}).get("组织架构_列别名") or {"姓名": ["姓名"], "部门编码": ["部门编码"]}
    chosen = None
    for name in wb.sheetnames:
        if name == "组织架构":
            chosen = name
            break
    if chosen is None:
        for name in wb.sheetnames:
            if "组织架构" in str(name) and "营销" not in str(name):
                chosen = name
                break
    if chosen is None:
        return {}
    ws = wb[chosen]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        return {}
    headers = [str(c).strip() if c is not None else "" for c in rows[0]]
    name_i = dept_i = None
    for i, h in enumerate(headers):
        if h in (org_a.get("姓名") or ["姓名"]) and name_i is None:
            name_i = i
        if h in (org_a.get("部门编码") or ["部门编码"]) and dept_i is None:
            dept_i = i
    if name_i is None or dept_i is None:
        return {}
    out: dict[str, str] = {}
    for row in rows[1:]:
        if not row or name_i >= len(row) or dept_i >= len(row):
            continue
        person = str(row[name_i] or "").strip()
        dept = code_str(row[dept_i])
        if person and dept and not dept.startswith("="):
            out[person] = dept
            out[norm_name(person)] = dept
    return out


def dept_for_sales(
    sales: str,
    org_map: dict[str, str],
    applicant_dept: dict[str, str],
    emp_dept: str = "",
) -> str:
    raw = str(sales or "").strip()
    if not raw:
        return ""
    return (
        org_map.get(raw)
        or org_map.get(norm_name(raw))
        or applicant_dept.get(raw)
        or applicant_dept.get(norm_name(raw))
        or str(emp_dept or "").strip()
    )


def classify_xlsx_kinds(path: Path, aliases: dict) -> list[str]:
    if "结果" in path.stem:
        return []
    if "核算项目余额表" in path.name:
        return []
    sheets = header_names(path)
    if not sheets:
        return []
    kinds: list[str] = []
    for name, headers in sheets.items():
        if is_sales_sheet(name, headers, aliases) and "销项发票" not in kinds:
            kinds.append("销项发票")
        if is_payment_sheet(name, headers, aliases) and "付款" not in kinds:
            kinds.append("付款")
        if is_receipt_sheet(name, headers, aliases) and "收款" not in kinds:
            kinds.append("收款")
    return kinds


def classify_xlsx(path: Path, aliases: dict) -> str | None:
    kinds = classify_xlsx_kinds(path, aliases)
    return kinds[0] if kinds else None


def receipt_sales_complete(root: Path, aliases: dict | None = None) -> bool:
    """每一行有客户的收款都填了销售才 True；缺列或有空值则 False（这时才登智云）。"""
    root = Path(root)
    if not root.is_dir():
        return False
    aliases = aliases or load_aliases()
    rec_a = aliases.get("收款_列别名") or {}
    saw = False
    for p in sorted(root.iterdir()):
        if p.suffix.lower() not in {".xlsx", ".xlsm"} or p.name.startswith("~$"):
            continue
        if "收款" not in classify_xlsx_kinds(p, aliases):
            continue
        try:
            wb = load_workbook(p, data_only=False)
        except Exception:
            return False
        try:
            sheet = find_receipt_sheet(wb, aliases)
            if not sheet:
                return False
            ws = wb[sheet]
            header_r, raw = find_header_row(ws, rec_a, ["客户名称", "借方（增加）"])
            cust_i = _field_index(raw, rec_a, "客户名称")
            sales_i = _field_index(raw, rec_a, "销售")
            if cust_i is None or sales_i is None:
                return False
            n = 0
            for row in ws.iter_rows(min_row=header_r + 1, max_row=ws.max_row or header_r, values_only=True):
                if not row or cust_i >= len(row):
                    continue
                cust = str(row[cust_i] or "").strip()
                if not cust:
                    continue
                n += 1
                sales = row[sales_i] if sales_i < len(row) else None
                if not str(sales or "").strip():
                    return False
            if n == 0:
                return False
            saw = True
        finally:
            wb.close()
    return saw


def payment_folders(root: Path, ledger: Path | None) -> list[Path]:
    skip = {ledger.resolve()} if ledger else set()
    out = []
    for p in sorted(root.iterdir()):
        if p.is_dir() and p.resolve() not in skip:
            pdfs = list(p.glob("*.pdf")) + list(p.glob("*.PDF"))
            if pdfs:
                out.append(p)
    return out


def inspect_dir(input_dir: Path, scene: str | None = None) -> dict:
    root = Path(input_dir)
    aliases = load_aliases()
    found: dict[str, list[str]] = {"销项发票": [], "付款": [], "收款": []}
    if not root.is_dir():
        return {
            "ready": False,
            "scene": scene,
            "mixed": False,
            "missing": ["材料文件夹"],
            "files": {},
            "ask": "请把当批材料放到一个文件夹里，再告诉我路径。",
        }
    for p in sorted(root.iterdir()):
        if p.suffix.lower() in {".xlsx", ".xlsm"} and not p.name.startswith("~$"):
            for kind in classify_xlsx_kinds(p, aliases):
                if str(p) not in found[kind]:
                    found[kind].append(str(p))
    posting_hits = [k for k in ("销项发票", "收款") if found.get(k)]
    mixed = len(posting_hits) > 1
    if scene is None:
        if not posting_hits and found.get("付款"):
            return {
                "ready": False,
                "scene": None,
                "scenes": [],
                "mixed": False,
                "missing": ["销项或收款表"],
                "files": {},
                "ask": "这是付款材料。请说「付款入金蝶」，走供应商付款技能。本技能只做销项和收款。",
            }
        targets = posting_hits
    else:
        targets = [scene]
    missing: list[str] = []
    files: dict = {}
    asks: list[str] = []
    ok: list[str] = []
    if not targets:
        return {
            "ready": False,
            "scene": None,
            "scenes": [],
            "mixed": False,
            "missing": ["能认出来的业务表"],
            "files": {},
            "ask": "这个文件夹里我还没认出销项发票、付款或收款表。请放好对应 Excel，或直接说要跑哪一句。",
        }
    for chosen in targets:
        if chosen == "销项发票":
            if not found["销项发票"]:
                missing.append("发票簿（要有单位名称、价税合计、申请人）")
                asks.append("还缺发票簿。把表放进这个文件夹即可，不用改文件名，也不用先填科目。")
            else:
                files["invoice"] = found["销项发票"][0]
                ok.append("销项发票")
        elif chosen == "付款":
            pay_missing: list[str] = []
            if not found["付款"]:
                pay_missing.append("付款三列表（供应商 / 应付金额本币 / 开户名）")
            ledger = Path(found["付款"][0]) if found["付款"] else None
            folders = payment_folders(root, ledger)
            files["ledger"] = str(ledger) if ledger else None
            files["invoice_dirs"] = [str(p) for p in folders]
            if not folders:
                pay_missing.append("各家发票夹（夹里要有 PDF）")
            if pay_missing:
                missing.extend(pay_missing)
                asks.append("付款还缺：" + "；".join(pay_missing) + "。台账放外面，一家一个夹，夹里放发票 PDF。")
            else:
                ok.append("付款")
        elif chosen == "收款":
            if not found["收款"]:
                missing.append("收款表（日期 / 客户名称 / 借方（增加）；中行收款 sheet 即可）")
                asks.append("还缺收款表。把月底稿放进这个文件夹即可，技能读「中行收款」。部门用组织架构，不必另填部门编码列。")
            else:
                files["receipt"] = found["收款"][0]
                ok.append("收款")
        else:
            missing.append("能认出来的业务表")
            asks.append("这个文件夹里我还没认出销项发票、付款或收款表。请放好对应 Excel，或直接说要跑哪一句。")
    if scene:
        ready = not missing and bool(ok)
        scene_out = scene
        scenes_out = [scene]
    else:
        ready = bool(ok)
        scene_out = ok[0] if len(ok) == 1 else ("全部" if ok else None)
        scenes_out = list(ok)
    return {
        "ready": ready,
        "scene": scene_out,
        "scenes": scenes_out,
        "mixed": mixed,
        "missing": missing,
        "files": files,
        "ask": " ".join(asks) if not ready else "",
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="金蝶入账 · 盘点文件夹")
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--scene", choices=["销项发票", "付款", "收款"])
    args = parser.parse_args(argv)
    report = inspect_dir(Path(args.input_dir), args.scene)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report.get("ready") else 2


if __name__ == "__main__":
    raise SystemExit(main())
