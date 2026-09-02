#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sniff 科目余额表 / 核算项目余额表 / 利润表 by sheet and headers, not filename."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from openpyxl import load_workbook

from common import detect_period_text, discover_input_dir
from layout import load_books, load_export_aliases


def _cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def sheet_blob(ws, max_row: int = 20, max_col: int = 16) -> str:
    parts = [_cell_text(ws.title)]
    for r in range(1, min(ws.max_row or 1, max_row) + 1):
        for c in range(1, min(ws.max_column or 1, max_col) + 1):
            t = _cell_text(ws.cell(r, c).value)
            if t:
                parts.append(t)
    return "\n".join(parts)


def header_map(ws, aliases: dict, max_row: int = 12) -> dict[str, tuple[int, int]]:
    wanted = {}
    for key in (
        "account_code",
        "account_name",
        "period_debit",
        "period_credit",
        "dept_name",
        "dept_code",
        "profit_item",
        "profit_month",
    ):
        for alias in aliases.get(key) or []:
            wanted.setdefault(alias, key)
    found: dict[str, tuple[int, int]] = {}
    last_row = min(ws.max_row or 1, max_row)
    last_col = min(ws.max_column or 1, 40)
    for r in range(1, last_row + 1):
        for c in range(1, last_col + 1):
            t = _cell_text(ws.cell(r, c).value)
            if t in wanted and wanted[t] not in found:
                found[wanted[t]] = (r, c)
    if "period_debit" not in found or "period_credit" not in found:
        for r in range(1, last_row + 1):
            for c in range(1, last_col + 1):
                t = _cell_text(ws.cell(r, c).value)
                if t not in {"本期发生额", "本期发生"}:
                    continue
                if r >= last_row:
                    continue
                left = _cell_text(ws.cell(r + 1, c).value)
                right = _cell_text(ws.cell(r + 1, c + 1).value) if c < last_col else ""
                if left == "借方" and "period_debit" not in found:
                    found["period_debit"] = (r + 1, c)
                if right == "贷方" and "period_credit" not in found:
                    found["period_credit"] = (r + 1, c + 1)
    return found


def _has_any(blob: str, hints: list[str]) -> bool:
    return any(h and h in blob for h in hints)


def classify_sheet(ws, aliases: dict) -> str | None:
    blob = sheet_blob(ws)
    title = _cell_text(ws.title)
    if _has_any(title, aliases.get("skip_sheet_hints") or []) or _has_any(blob.split("\n")[0], aliases.get("skip_sheet_hints") or []):
        if "确认情况" in title:
            return None
    headers = header_map(ws, aliases)
    if _has_any(blob, aliases.get("skip_sheet_hints") or []) and "客户名称" in blob and "确认情况" in blob:
        return None
    title_s = title or ""
    if title_s == "损益表" and ("山东分公司" in blob or "济南子公司" in blob or "四川分公司" in blob):
        return "monthly_overlay"
    if title_s == "利润表" and "山东" in blob and "湖南分公司" in blob and "本月金额" not in blob:
        return "monthly_profit"
    if "profit_month" in headers and ("profit_item" in headers or _has_any(blob, aliases.get("profit_sheet_hints") or [])):
        return "profit"
    if "account_code" in headers and "dept_name" in headers:
        return "assist"
    if _has_any(blob, aliases.get("assist_sheet_hints") or []) and "account_code" in headers:
        return "assist"
    if "account_code" in headers and ("period_debit" in headers or "period_credit" in headers):
        if _has_any(blob, aliases.get("account_sheet_hints") or []) or "科目" in blob:
            return "account"
        return "account"
    if _has_any(blob, aliases.get("profit_sheet_hints") or []) and "profit_month" in headers:
        return "profit"
    return None


def _entity_candidates(text: str, books: dict) -> list[tuple[int, str]]:
    candidates: list[tuple[int, str]] = []
    for acc in books.get("xingchen_accounts") or []:
        header = acc.get("excel_header") or ""
        legal = acc.get("legal_name") or ""
        if legal and legal in text:
            candidates.append((len(legal), header))
        elif header and header in text:
            candidates.append((len(header), header))
    for item in books.get("no_xingchen_leave_blank") or []:
        header = item.get("excel_header") or ""
        if header and header in text:
            candidates.append((len(header), header))
    return candidates


def match_entity(blob: str, filename: str, books: dict) -> str | None:
    company_line = next((line for line in blob.splitlines() if "公司名称" in line), "")
    candidates = _entity_candidates(company_line, books) if company_line else []
    if not candidates:
        candidates = _entity_candidates(blob, books)
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]
    # filename is a hint only after content miss
    name = Path(filename).name
    for acc in sorted(books.get("xingchen_accounts") or [], key=lambda x: len(x.get("excel_header") or ""), reverse=True):
        header = acc.get("excel_header") or ""
        legal = acc.get("legal_name") or ""
        if header and header in name:
            return header
        if legal and legal in name:
            return header
    return None


def inspect_file(path: Path) -> list[dict]:
    aliases = load_export_aliases()
    books = load_books()
    found = []
    if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm"} or path.name.startswith("~$"):
        return found
    try:
        wb = load_workbook(path, data_only=False, read_only=True)
    except Exception:
        return found
    try:
        for title in wb.sheetnames:
            ws = wb[title]
            kind = classify_sheet(ws, aliases)
            if not kind:
                continue
            blob = sheet_blob(ws)
            entity = match_entity(blob + "\n" + path.name, path.name, books)
            headers = header_map(ws, aliases)
            found.append(
                {
                    "path": str(path.resolve()),
                    "sheet": title,
                    "kind": kind,
                    "entity": entity,
                    "period": detect_period_text(blob + "\n" + path.name),
                    "headers": {k: {"row": v[0], "col": v[1]} for k, v in headers.items()},
                }
            )
    finally:
        wb.close()
    return found


def inspect_dir(input_dir: Path) -> list[dict]:
    found = []
    if not input_dir.is_dir():
        return found
    for path in sorted(input_dir.rglob("*")):
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        found.extend(inspect_file(path))
    return found


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", default="")
    args = parser.parse_args(argv)
    rows = inspect_dir(discover_input_dir(args.input_dir))
    print(json.dumps({"files": len(rows), "kinds": [r["kind"] for r in rows], "entities": [r.get("entity") for r in rows]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
