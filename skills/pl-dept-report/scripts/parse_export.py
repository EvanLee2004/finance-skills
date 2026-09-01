#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from common import money
from inspect_inputs import header_map
from layout import load_export_aliases


def _norm_code(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", "").strip()
    if text.endswith(".0") and text.replace(".", "", 1)[:-1].isdigit():
        text = text[:-2]
    return text


def parse_account_sheet(ws, headers: dict) -> dict[str, dict[str, Decimal | None]]:
    code_at = headers.get("account_code")
    debit_at = headers.get("period_debit")
    credit_at = headers.get("period_credit")
    if not code_at:
        return {}
    start_row = code_at[0] + 1
    out: dict[str, dict[str, Decimal | None]] = {}
    for r in range(start_row, (ws.max_row or start_row) + 1):
        code = _norm_code(ws.cell(r, code_at[1]).value)
        if not code or not code[0].isdigit():
            continue
        debit = money(ws.cell(r, debit_at[1]).value) if debit_at else None
        credit = money(ws.cell(r, credit_at[1]).value) if credit_at else None
        if code not in out:
            out[code] = {"debit": debit, "credit": credit}
        else:
            from common import add_money

            out[code]["debit"] = add_money(out[code]["debit"], debit)
            out[code]["credit"] = add_money(out[code]["credit"], credit)
    return out


def parse_assist_sheet(ws, headers: dict) -> list[dict]:
    code_at = headers.get("account_code")
    dept_at = headers.get("dept_name")
    debit_at = headers.get("period_debit")
    credit_at = headers.get("period_credit")
    if not code_at or not dept_at:
        return []
    start_row = max(code_at[0], dept_at[0]) + 1
    rows = []
    for r in range(start_row, (ws.max_row or start_row) + 1):
        code = _norm_code(ws.cell(r, code_at[1]).value)
        dept = str(ws.cell(r, dept_at[1]).value or "").replace("\xa0", "").strip()
        if not code or not dept:
            continue
        rows.append(
            {
                "code": code,
                "dept": dept,
                "debit": money(ws.cell(r, debit_at[1]).value) if debit_at else None,
                "credit": money(ws.cell(r, credit_at[1]).value) if credit_at else None,
            }
        )
    return rows


def parse_profit_sheet(ws, headers: dict, aliases: dict) -> dict[str, Decimal | None]:
    item_at = headers.get("profit_item")
    amt_at = headers.get("profit_month")
    if not amt_at:
        return {}
    start_row = (item_at[0] if item_at else amt_at[0]) + 1
    item_col = item_at[1] if item_at else 1
    out: dict[str, Decimal | None] = {}
    for r in range(start_row, (ws.max_row or start_row) + 1):
        label = str(ws.cell(r, item_col).value or "").replace("\xa0", "").strip()
        if not label:
            continue
        out[label] = money(ws.cell(r, amt_at[1]).value)
    return out


def parse_inspected(items: list[dict]) -> dict:
    aliases = load_export_aliases()
    accounts: dict[str, dict[str, dict[str, Decimal | None]]] = {}
    depts: list[dict] = []
    profits: dict[str, dict[str, Decimal | None]] = {}
    profits_by_period: dict[str, dict[str, dict[str, Decimal | None]]] = {}
    extra_codes: set[str] = set()
    for item in items:
        path = item["path"]
        wb = load_workbook(path, data_only=False)
        try:
            ws = wb[item["sheet"]]
            headers = {k: (v["row"], v["col"]) for k, v in (item.get("headers") or {}).items()}
            if not headers:
                headers = header_map(ws, aliases)
            entity = item.get("entity")
            kind = item["kind"]
            if kind == "account":
                parsed = parse_account_sheet(ws, headers)
                if entity:
                    bucket = accounts.setdefault(entity, {})
                    for code, pair in parsed.items():
                        if code not in bucket:
                            bucket[code] = pair
                        else:
                            from common import add_money

                            bucket[code] = {
                                "debit": add_money(bucket[code]["debit"], pair["debit"]),
                                "credit": add_money(bucket[code]["credit"], pair["credit"]),
                            }
            elif kind == "assist":
                for row in parse_assist_sheet(ws, headers):
                    row["entity"] = entity
                    depts.append(row)
            elif kind == "profit" and entity:
                parsed = parse_profit_sheet(ws, headers, aliases)
                profits[entity] = parsed
                period = item.get("period")
                if period:
                    profits_by_period.setdefault(period, {})[entity] = parsed
        finally:
            wb.close()
    return {
        "accounts": accounts,
        "depts": depts,
        "profits": profits,
        "profits_by_period": profits_by_period,
        "extra_codes": sorted(extra_codes),
    }
