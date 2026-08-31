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


def header_names(path: Path) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    try:
        wb = load_workbook(path, read_only=True, data_only=False)
    except Exception:
        return out
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            out[name] = [str(c).strip() for c in (row or []) if c is not None and str(c).strip()]
    finally:
        wb.close()
    return out


def field_hit(headers: list[str], aliases: dict, field: str) -> bool:
    names = aliases.get(field) or [field]
    return any(h in names for h in headers)


def classify_xlsx(path: Path, aliases: dict) -> str | None:
    if "结果" in path.stem:
        return None
    sheets = header_names(path)
    if not sheets:
        return None
    sales_a = aliases.get("销项发票_列别名") or {}
    pay_a = aliases.get("付款_列别名") or {}
    rec_a = aliases.get("收款_列别名") or {}
    for headers in sheets.values():
        if field_hit(headers, sales_a, "单位名称") and field_hit(headers, sales_a, "价税合计"):
            return "销项发票"
        if field_hit(headers, pay_a, "供应商") and field_hit(headers, pay_a, "应付金额本币"):
            return "付款"
        if field_hit(headers, rec_a, "客户名称") and field_hit(headers, rec_a, "借方（增加）"):
            return "收款"
    return None


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
            kind = classify_xlsx(p, aliases)
            if kind:
                found[kind].append(str(p))
    hits = [k for k, v in found.items() if v]
    mixed = len(hits) > 1
    if scene:
        chosen = scene
    elif len(hits) == 1:
        chosen = hits[0]
    else:
        chosen = None
    missing = []
    files: dict = {}
    ask = ""
    if mixed and not scene:
        ask = "这个文件夹里同时有不止一种表。请说要跑「销项发票入金蝶」「付款入金蝶」还是「收款入金蝶」。"
        return {
            "ready": False,
            "scene": None,
            "mixed": True,
            "missing": ["指定模块"],
            "files": {k: v for k, v in found.items() if v},
            "ask": ask,
        }
    if chosen == "销项发票":
        if not found["销项发票"]:
            missing.append("发票簿（要有单位名称、价税合计、申请人）")
            ask = "还缺发票簿。把表放进这个文件夹即可，不用改文件名，也不用先填科目。"
        else:
            files["invoice"] = found["销项发票"][0]
    elif chosen == "付款":
        if not found["付款"]:
            missing.append("付款三列表（供应商 / 应付金额本币 / 开户名）")
        ledger = Path(found["付款"][0]) if found["付款"] else None
        folders = payment_folders(root, ledger)
        files["ledger"] = str(ledger) if ledger else None
        files["invoice_dirs"] = [str(p) for p in folders]
        if not folders:
            missing.append("各家发票夹（夹里要有 PDF）")
        if missing:
            ask = "付款还缺：" + "；".join(missing) + "。台账放外面，一家一个夹，夹里放发票 PDF。"
    elif chosen == "收款":
        if not found["收款"]:
            missing.append("收款表（日期 / 客户名称 / 借方（增加）/ 部门编码）")
            ask = "还缺收款表。日期、客户、金额、部门编码放进这个文件夹即可，销售和科目由技能补。"
        else:
            files["receipt"] = found["收款"][0]
    else:
        missing.append("能认出来的业务表")
        ask = "这个文件夹里我还没认出销项发票、付款或收款表。请放好对应 Excel，或直接说要跑哪一句。"
    ready = not missing and chosen is not None
    return {
        "ready": ready,
        "scene": chosen,
        "mixed": mixed,
        "missing": missing,
        "files": files,
        "ask": ask if not ready else "",
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
