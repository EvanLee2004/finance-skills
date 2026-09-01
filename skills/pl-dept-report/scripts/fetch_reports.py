#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from decimal import Decimal
from pathlib import Path
import json
import os
import time

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


def _match_profit_label(name: str, aliases: dict) -> str | None:
    raw = str(name or "").replace(" ", "").replace("：", ":")
    if not raw:
        return None
    candidates = []
    for label, names in aliases.items():
        keys = [str(label)] + [str(n) for n in (names or [])]
        for key in keys:
            k = key.replace(" ", "").replace("：", ":")
            if k and k in raw:
                candidates.append((len(k), label))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1]


def map_profit_rows(raw_rows: list[dict], layout: dict) -> dict[str, Decimal | None]:
    aliases = layout.get("profit_item_aliases") or {}
    out: dict[str, Decimal | None] = {}
    for row in raw_rows:
        name = _item_name(row)
        label = _match_profit_label(name, aliases)
        if not label:
            continue
        amt = _current_amount(row)
        if amt is not None and label not in out:
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


def _voucher_ids(creds: dict, token: str, period: str) -> tuple[list[str], list[str]]:
    extra = {"app-token": token}
    ids: list[str] = []
    notes: list[str] = []
    page = 1
    total = None
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
        for head in rows:
            vid = str((head or {}).get("id") or "").strip()
            if vid:
                ids.append(vid)
        if not rows or len(rows) < 100 or (total and len(ids) >= total):
            break
        page += 1
    notes.append(f"voucher_count={len(ids)}")
    return ids, notes


def _detail_entries(creds: dict, token: str, vid: str):
    try:
        detail = request(
            "GET",
            "/jdy/v2/fi/voucher_detail",
            creds,
            params={"id": vid},
            extra={"app-token": token},
            timeout=12,
            retries=1,
        )
    except Exception:
        return None
    if detail.status_code != 200:
        return None
    try:
        body = detail.json()
    except Exception:
        return None
    data = body.get("data") if isinstance(body, dict) else None
    entries = (data or {}).get("entry_list") if isinstance(data, dict) else None
    return entries if isinstance(entries, list) else []


def fetch_vouchers_pl(creds: dict, token: str, period: str, layout: dict) -> dict:
    codes = {row["code"] for row in layout["accounts"] if row.get("code")}
    ids, notes = _voucher_ids(creds, token, period)
    accounts: dict[str, dict] = {}
    depts: list[dict] = []
    ok = 0
    miss = 0
    with ThreadPoolExecutor(max_workers=8) as pool:
        futs = {pool.submit(_detail_entries, creds, token, vid): vid for vid in ids}
        for fut in as_completed(futs):
            entries = fut.result()
            if entries is None:
                miss += 1
                continue
            ok += 1
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
    notes.append(f"voucher_detail_ok={ok}")
    notes.append(f"voucher_detail_miss={miss}")
    notes.append(f"hq_account_codes={len(accounts)}")
    notes.append(f"hq_dept_lines={len(depts)}")
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


def cache_path(period: str) -> Path:
    return Path.home() / ".cache" / "finance" / f"pl-hq-{period}.json"


def load_hq_cache(period: str) -> dict | None:
    path = cache_path(period)
    if not path.is_file():
        return None
    if time.time() - path.stat().st_mtime > 12 * 3600:
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    accounts = {}
    for code, pair in (raw.get("accounts") or {}).items():
        accounts[code] = {"debit": money(pair.get("debit")), "credit": money(pair.get("credit"))}
    depts = []
    for row in raw.get("depts") or []:
        depts.append(
            {
                "code": row.get("code"),
                "dept": row.get("dept"),
                "debit": money(row.get("debit")),
                "credit": money(row.get("credit")),
                "entity": "甲骨易",
            }
        )
    profit = {}
    for k, v in (raw.get("profit") or {}).items():
        profit[k] = money(v)
    return {"accounts": accounts, "depts": depts, "profit": profit, "notes": raw.get("notes") or ["hq_cache"]}


def save_hq_cache(period: str, payload: dict) -> None:
    path = cache_path(period)
    path.parent.mkdir(parents=True, exist_ok=True)
    dump = {
        "period": period,
        "notes": payload.get("notes") or [],
        "accounts": {
            k: {
                "debit": str(v["debit"]) if v.get("debit") is not None else None,
                "credit": str(v["credit"]) if v.get("credit") is not None else None,
            }
            for k, v in (payload.get("accounts") or {}).items()
        },
        "depts": [
            {
                "code": r.get("code"),
                "dept": r.get("dept"),
                "debit": str(r["debit"]) if r.get("debit") is not None else None,
                "credit": str(r["credit"]) if r.get("credit") is not None else None,
            }
            for r in (payload.get("depts") or [])
        ],
        "profit": {k: str(v) for k, v in (payload.get("profit") or {}).items() if v is not None},
    }
    path.write_text(json.dumps(dump, ensure_ascii=False), encoding="utf-8")
    path.chmod(0o600)


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
        cached = None if os.environ.get("PL_DEPT_USE_CACHE") == "0" else load_hq_cache(period)
        if cached and cached.get("accounts"):
            notes.append("hq_ledger_cache")
            notes.extend(cached.get("notes") or [])
            return {
                "entity": "甲骨易",
                "accounts": cached["accounts"],
                "depts": cached["depts"],
                "profit": profit,
                "notes": notes,
                "department_master": [],
            }
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
    result = {
        "entity": "甲骨易",
        "accounts": accounts,
        "depts": depts,
        "profit": profit,
        "notes": notes,
        "department_master": depts_master if ledger else [],
    }
    if ledger and accounts:
        save_hq_cache(period, result)
    return result
