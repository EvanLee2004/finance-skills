#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

from decimal import Decimal

from common import add_money, money
from kingdee_client import get_app_token, request, rows_from
from layout import load_layout, parent_code


def _item_name(row: dict) -> str:
    for key in ("item_name", "name", "project_name", "item", "caption", "project"):
        val = row.get(key)
        if val:
            return str(val).strip()
    return ""


def _current_amount(row: dict):
    for key in ("current_amount", "currentAmount", "this_month_amount", "month_amount", "current"):
        if row.get(key) is not None and str(row.get(key)).strip() != "":
            return money(row.get(key))
    return None


def fetch_profit(creds: dict, token: str, period: str) -> tuple[list[dict], str | None]:
    resp = request(
        "GET",
        "/jdy/v2/fi/profit_report",
        creds,
        params={"period": period},
        extra={"app-token": token},
    )
    if resp.status_code != 200:
        return [], f"profit_report http {resp.status_code}"
    try:
        payload = resp.json()
    except Exception:
        return [], "profit_report not json"
    rows = rows_from(payload)
    return [row for row in rows if isinstance(row, dict)], None


def map_profit_rows(raw_rows: list[dict], layout: dict) -> dict[str, Decimal | None]:
    aliases = layout.get("profit_item_aliases") or {}
    out: dict[str, Decimal | None] = {}
    by_alias: dict[str, str] = {}
    for label, names in aliases.items():
        by_alias[label] = label
        for name in names:
            by_alias[str(name).strip()] = label
    for row in raw_rows:
        name = _item_name(row)
        if not name:
            continue
        label = by_alias.get(name)
        if not label:
            continue
        amt = _current_amount(row)
        if amt is not None:
            out[label] = amt
    return out


def fetch_account_balance(creds: dict, token: str, period: str) -> tuple[list[dict], str | None]:
    attempts = [
        {"period": period},
        {"start_period": period, "end_period": period},
        {"year": period[:4], "period": period[4:6]},
    ]
    last = "not tried"
    for params in attempts:
        resp = request(
            "GET",
            "/jdy/v2/fi/account_balance_report",
            creds,
            params=params,
            extra={"app-token": token},
        )
        if resp.status_code == 200:
            try:
                payload = resp.json()
            except Exception:
                return [], "account_balance not json"
            return [row for row in rows_from(payload) if isinstance(row, dict)], None
        last = f"account_balance_report http {resp.status_code}"
    return [], last


def _entry_account(entry: dict) -> str:
    for key in ("account_number", "accountNumber", "number", "account"):
        val = entry.get(key)
        if val:
            return str(val).strip()
    return ""


def _dept_from_assist(assist) -> str:
    if not isinstance(assist, list):
        return ""
    for item in assist:
        if not isinstance(item, dict):
            continue
        typ = str(item.get("type") or item.get("bd_type") or "")
        if typ in {"bd_department", "department"} or "department" in typ:
            return str(item.get("name") or item.get("number") or "").strip()
    return ""


def fetch_vouchers_pl(creds: dict, token: str, period: str, layout: dict) -> dict:
    extra = {"app-token": token}
    codes = {row["code"] for row in layout["accounts"] if row.get("code")}
    accounts: dict[str, dict] = {}
    depts: list[dict] = []
    page = 1
    seen: set[str] = set()
    total = None
    notes = []
    while page <= 200:
        resp = request(
            "GET",
            "/jdy/v2/fi/voucher",
            creds,
            params={"start_period": period, "end_period": period, "page": str(page), "page_size": "100"},
            extra=extra,
        )
        if resp.status_code != 200:
            notes.append(f"voucher http {resp.status_code}")
            break
        payload = resp.json()
        if total is None and isinstance(payload, dict) and isinstance(payload.get("data"), dict):
            try:
                total = int(payload["data"].get("count") or 0)
            except (TypeError, ValueError):
                total = 0
        rows = rows_from(payload)
        if not rows:
            break
        new_rows = 0
        for head in rows:
            vid = str((head or {}).get("id") or "").strip()
            if not vid or vid in seen:
                continue
            seen.add(vid)
            new_rows += 1
            try:
                detail = request(
                    "GET",
                    "/jdy/v2/fi/voucher_detail",
                    creds,
                    params={"id": vid},
                    extra=extra,
                    timeout=15,
                    retries=2,
                )
            except Exception:
                continue
            if detail.status_code != 200:
                continue
            try:
                body = detail.json()
            except Exception:
                continue
            data = body.get("data") if isinstance(body, dict) else None
            entries = (data or {}).get("entry_list") if isinstance(data, dict) else None
            if not isinstance(entries, list):
                continue
            for entry in entries:
                if not isinstance(entry, dict):
                    continue
                acc = _entry_account(entry)
                if not acc.startswith("5"):
                    continue
                debit = money(entry.get("debit_amount") or entry.get("debitAmount"))
                credit = money(entry.get("credit_amount") or entry.get("creditAmount"))
                chain = [acc]
                cur = acc
                while True:
                    p = parent_code(cur, codes)
                    if not p:
                        break
                    chain.append(p)
                    cur = p
                for code in chain:
                    if code not in codes:
                        continue
                    bucket = accounts.setdefault(code, {"debit": None, "credit": None})
                    bucket["debit"] = add_money(bucket["debit"], debit)
                    bucket["credit"] = add_money(bucket["credit"], credit)
                dept = _dept_from_assist(entry.get("assist"))
                if dept and acc in codes:
                    depts.append({"code": acc, "dept": dept, "debit": debit, "credit": credit, "entity": "甲骨易"})
        if new_rows == 0 or len(rows) < 100 or (total and len(seen) >= total):
            break
        page += 1
    notes.append(f"voucher_count={len(seen)}")
    return {"accounts": accounts, "depts": depts, "notes": notes}


def fetch_departments(creds: dict, token: str) -> list[dict]:
    resp = request(
        "GET",
        "/jdy/v2/bd/department",
        creds,
        params={"page": "1", "page_size": "2000"},
        extra={"app-token": token},
    )
    if resp.status_code != 200:
        return []
    try:
        payload = resp.json()
    except Exception:
        return []
    out = []
    for item in rows_from(payload):
        if not isinstance(item, dict):
            continue
        name = str(item.get("name") or "").strip()
        number = str(item.get("number") or item.get("code") or "").strip()
        if name:
            out.append({"name": name, "number": number})
    return out


def fetch_hq(period: str, creds: dict, *, ledger: bool = True) -> dict:
    layout = load_layout()
    token, _domain = get_app_token(creds)
    notes = []
    profit_rows, profit_err = fetch_profit(creds, token, period)
    if profit_err:
        notes.append(profit_err)
    profit = map_profit_rows(profit_rows, layout)
    notes.append(f"profit_rows={len(profit_rows)}")
    accounts: dict = {}
    depts: list = []
    depts_master: list = []
    if ledger:
        bal_rows, bal_err = fetch_account_balance(creds, token, period)
        if bal_err:
            notes.append(bal_err)
            scanned = fetch_vouchers_pl(creds, token, period, layout)
            accounts = scanned["accounts"]
            depts = scanned["depts"]
            notes.extend(scanned.get("notes") or [])
        else:
            notes.append(f"account_balance_rows={len(bal_rows)}")
            for row in bal_rows:
                acc = str(row.get("account_number") or row.get("number") or row.get("account") or "").strip()
                if not acc.startswith("5"):
                    continue
                accounts[acc] = {
                    "debit": money(row.get("current_debit") or row.get("debit") or row.get("period_debit")),
                    "credit": money(row.get("current_credit") or row.get("credit") or row.get("period_credit")),
                }
            scanned = fetch_vouchers_pl(creds, token, period, layout)
            depts = scanned["depts"]
            notes.append("assist_openapi_skipped")
            notes.extend(scanned.get("notes") or [])
        depts_master = fetch_departments(creds, token)
        notes.append(f"departments={len(depts_master)}")
    return {
        "entity": "甲骨易",
        "accounts": accounts,
        "depts": depts,
        "profit": profit,
        "notes": notes,
        "department_master": depts_master if ledger else [],
    }
