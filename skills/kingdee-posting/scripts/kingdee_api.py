#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金蝶云星辰 OpenAPI 只读查档。密钥只读本机文件，不进 git、不打印。"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
import uuid
from pathlib import Path
from urllib.parse import quote

import requests

DEFAULT_LOCAL = Path.home() / ".config" / "finance" / "kingdee.local.json"
DEFAULT_CACHE = Path.home() / ".cache" / "finance" / "kingdee-master.json"
API_HOST = "https://api.kingdee.com"
AUTH_PATH = "/jdyconnector/app_management/kingdee_auth_token"
AUTHORIZE_PATH = "/jdyconnector/app_management/push_app_authorize"
NEED_KEYS = ("client_id", "client_secret", "app_key", "app_secret")
MASTER_KINDS = ("customer", "employee", "supplier", "department")
MASTER_CACHE_TTL_SECONDS = 15 * 60
# 主动获取授权只把这些写回本机；禁止把 accessToken / appToken 落盘。
AUTHORIZE_FIELDS = (
    ("appKey", "app_key"),
    ("appSecret", "app_secret"),
    ("accountId", "account_id"),
    ("accountName", "account_name"),
    ("agreementCompanyName", "company_name"),
    ("serviceId", "service_id"),
    ("outerInstanceId", "outer_instance_id"),
    ("domain", "domain"),
    ("groupName", "group_name"),
)


def local_path() -> Path:
    override = os.environ.get("KINGDEE_LOCAL_JSON", "").strip()
    return Path(override) if override else DEFAULT_LOCAL


def cache_path() -> Path:
    override = os.environ.get("KINGDEE_MASTER_CACHE", "").strip()
    return Path(override) if override else DEFAULT_CACHE


def cache_ttl_seconds() -> int:
    raw = os.environ.get("KINGDEE_MASTER_CACHE_TTL_SECONDS", "").strip()
    try:
        return max(0, int(raw)) if raw else MASTER_CACHE_TTL_SECONDS
    except ValueError:
        return MASTER_CACHE_TTL_SECONDS


def load_fresh_cache() -> dict | None:
    path = cache_path()
    if not path.is_file():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        cached_at = float(payload["cached_at"])
        data = payload["data"]
    except (OSError, ValueError, KeyError, TypeError):
        return None
    if time.time() - cached_at > cache_ttl_seconds():
        return None
    if not isinstance(data, dict) or any(not isinstance(data.get(kind), list) for kind in MASTER_KINDS):
        return None
    return data


def save_cache(data: dict) -> None:
    path = cache_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"cached_at": time.time(), "data": data}, ensure_ascii=False), encoding="utf-8")
        path.chmod(0o600)
    except OSError:
        # 缓存只是性能优化；写缓存失败不能改变本次已成功的只读查档结果。
        return


def load_local() -> dict | None:
    path = local_path()
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    if any(not str(data.get(k) or "").strip() for k in NEED_KEYS):
        return None
    return data


def hmac_sha256_hex_b64(secret: str, message: str) -> str:
    digest_hex = hmac.new(
        secret.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    return base64.b64encode(digest_hex.encode("utf-8")).decode("ascii")


def app_signature(app_key: str, app_secret: str) -> str:
    return hmac_sha256_hex_b64(app_secret, app_key)


def double_encode(value: str) -> str:
    once = quote(str(value), safe="")
    twice = quote(once, safe="")
    return twice


def encode_path(path: str) -> str:
    return quote(path, safe="")


def sign_plain(method: str, path: str, params: dict | None, nonce: str, timestamp: str) -> str:
    encoded_params = ""
    if params:
        encoded_params = "&".join(
            f"{double_encode(k)}={double_encode(v)}" for k, v in sorted(params.items())
        )
    headers_block = f"x-api-nonce:{nonce}\nx-api-timestamp:{timestamp}"
    return f"{method.upper()}\n{encode_path(path)}\n{encoded_params}\n{headers_block}\n"


def x_api_signature(client_secret: str, method: str, path: str, params: dict | None, nonce: str, timestamp: str) -> str:
    return hmac_sha256_hex_b64(client_secret, sign_plain(method, path, params, nonce, timestamp))


def _auth_headers(creds: dict, method: str, path: str, params: dict | None) -> dict:
    nonce = uuid.uuid4().hex[:16]
    timestamp = str(int(time.time() * 1000))
    return {
        "Content-Type": "application/json",
        "X-Api-ClientID": str(creds["client_id"]),
        "X-Api-Auth-Version": "2.0",
        "X-Api-TimeStamp": timestamp,
        "X-Api-Nonce": nonce,
        "X-Api-SignHeaders": "X-Api-TimeStamp,X-Api-Nonce",
        "X-Api-Signature": x_api_signature(
            str(creds["client_secret"]), method, path, params, nonce, timestamp
        ),
    }


def _request(method: str, url: str, creds: dict, path: str, params: dict | None = None, extra_headers: dict | None = None, timeout: int = 30):
    headers = _auth_headers(creds, method, path, params)
    if extra_headers:
        headers.update(extra_headers)
    last_err = None
    for attempt in range(4):
        try:
            resp = requests.request(method, url, headers=headers, params=params, timeout=timeout)
            return resp
        except (requests.ConnectionError, requests.Timeout) as e:
            last_err = e
            time.sleep(0.4 * (attempt + 1))
    raise last_err or RuntimeError("request failed")


def pick_authorize_row(rows: list, creds: dict) -> dict | None:
    """多账套同时授权时只认本机已绑的总部账套，禁止误用空账。"""
    live = [row for row in rows if isinstance(row, dict) and str(row.get("status")).split(".")[0] == "1"]
    wanted_acc = str(creds.get("account_id") or "").strip()
    wanted_svc = str(creds.get("service_id") or "").strip()
    wanted_key = str(creds.get("app_key") or "").strip()

    def _one(pred) -> dict | None:
        got = [row for row in live if pred(row)]
        return got[0] if len(got) == 1 else None

    if wanted_acc:
        hit = _one(lambda row: str(row.get("accountId") or "").strip() == wanted_acc)
        if hit:
            return hit
    if wanted_svc:
        hit = _one(lambda row: str(row.get("serviceId") or "").strip() == wanted_svc)
        if hit:
            return hit
    if wanted_key:
        hit = _one(lambda row: str(row.get("appKey") or "").strip() == wanted_key)
        if hit:
            return hit
    return live[0] if len(live) == 1 else None


def _save_local(creds: dict) -> None:
    path = local_path()
    existing: dict = {}
    if path.is_file():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            loaded = None
        if isinstance(loaded, dict):
            existing = loaded
    existing.update(creds)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(existing, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    path.chmod(0o600)


def apply_authorize_row(creds: dict, row: dict) -> dict:
    updated = dict(creds)
    for src, dst in AUTHORIZE_FIELDS:
        val = row.get(src)
        if val is None or str(val).strip() == "":
            continue
        updated[dst] = str(val)
    updated["secret_refreshed_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    _save_local(updated)
    return updated


def refresh_app_authorize(creds: dict) -> dict | None:
    oid = str(creds.get("outer_instance_id") or "").strip()
    if not oid:
        return None
    params = {"outerInstanceId": oid}
    resp = _request("POST", API_HOST + AUTHORIZE_PATH, creds, AUTHORIZE_PATH, params=params)
    if resp.status_code != 200:
        return None
    try:
        payload = resp.json()
    except Exception:
        return None
    rows = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(rows, list):
        return None
    row = pick_authorize_row(rows, creds)
    if not row or not str(row.get("appSecret") or "").strip():
        return None
    return apply_authorize_row(creds, row)


def _token_from_payload(payload, creds: dict) -> tuple[str, str] | None:
    data = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(data, dict):
        return None
    token = data.get("app-token") or data.get("app_token") or data.get("appToken")
    if not token:
        return None
    domain = data.get("domain") or creds.get("domain") or ""
    return str(token), str(domain).rstrip("/")


def get_app_token(creds: dict, *, allow_refresh: bool = True) -> tuple[str, str]:
    params = {
        "app_key": str(creds["app_key"]),
        "app_signature": app_signature(str(creds["app_key"]), str(creds["app_secret"])),
    }
    resp = _request("GET", API_HOST + AUTH_PATH, creds, AUTH_PATH, params=params)
    if resp.status_code != 200:
        raise RuntimeError(f"auth http {resp.status_code}")
    payload = resp.json()
    got = _token_from_payload(payload, creds)
    if got:
        return got
    err = payload.get("errcode") if isinstance(payload, dict) else None
    desc = str((payload or {}).get("description") or "")[:80] if isinstance(payload, dict) else ""
    if allow_refresh and str(err) == "1030002006":
        refreshed = refresh_app_authorize(creds)
        if refreshed:
            return get_app_token(refreshed, allow_refresh=False)
    raise RuntimeError(f"auth missing data errcode={err} {desc}".strip())


def _rows_from(payload) -> list:
    if isinstance(payload, list):
        return payload
    if not isinstance(payload, dict):
        return []
    data = payload.get("data")
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("rows", "list", "items"):
            rows = data.get(key)
            if isinstance(rows, list):
                return rows
    rows = payload.get("rows")
    return rows if isinstance(rows, list) else []


def _code_name(item: dict) -> dict | None:
    if not isinstance(item, dict):
        return None
    code = item.get("number") or item.get("emp_number") or item.get("code") or item.get("number_id")
    name = item.get("name") or item.get("emp_name")
    code = str(code).strip() if code is not None else ""
    name = str(name).strip() if name is not None else ""
    if not code or not name:
        return None
    return {"code": code, "name": name}


def fetch_list(creds: dict, token: str, domain: str, path: str, page_size: int = 2000) -> list[dict]:
    # app-token 的 domain 是租户域名；基础档案 OpenAPI 固定走官方网关。
    # 使用租户域名会返回 404，不能静默降级为空档案。
    host = API_HOST
    out: list[dict] = []
    page = 1
    extra = {"app-token": token}
    while page <= 80:
        params = {"page": str(page), "page_size": str(page_size)}
        resp = _request("GET", host + path, creds, path, params=params, extra_headers=extra)
        if resp.status_code != 200:
            raise RuntimeError(f"{path} http {resp.status_code}")
        try:
            payload = resp.json()
        except Exception as e:
            raise RuntimeError(f"{path} not json") from e
        rows = _rows_from(payload)
        for item in rows:
            mapped = _code_name(item)
            if mapped:
                out.append(mapped)
        if len(rows) < page_size:
            break
        page += 1
        time.sleep(0.12)
    return out


def try_load_master() -> dict:
    cached = load_fresh_cache()
    if cached is not None:
        return {"ok": True, "missing_credentials": False, "source": "cache", "data": cached}
    creds = load_local()
    if not creds:
        return {"ok": False, "missing_credentials": True, "data": None}
    try:
        token, domain = get_app_token(creds)
        data = {
            "customer": fetch_list(creds, token, domain, "/jdy/v2/bd/customer"),
            "employee": fetch_list(creds, token, domain, "/jdy/v2/bd/emp"),
            "supplier": fetch_list(creds, token, domain, "/jdy/v2/bd/supplier"),
            "department": fetch_list(creds, token, domain, "/jdy/v2/bd/department"),
        }
        save_cache(data)
        return {"ok": True, "missing_credentials": False, "source": "live", "data": data}
    except Exception as e:
        return {
            "ok": False,
            "missing_credentials": False,
            "error": f"{type(e).__name__}: {str(e)[:200]}",
            "data": None,
        }


def ar_cache_path() -> Path:
    override = os.environ.get("KINGDEE_AR_CACHE", "").strip()
    return Path(override) if override else Path.home() / ".cache" / "finance" / "kingdee-ar.json"


def _yyyymm(period: str) -> str:
    s = str(period or "").strip().replace("-", "").replace("/", "").replace(".", "")
    if len(s) >= 6 and s[:6].isdigit():
        return s[:6]
    return ""


def _customer_code_from_assist(assist) -> str:
    if not isinstance(assist, list):
        return ""
    for item in assist:
        if isinstance(item, dict) and str(item.get("type") or "") == "bd_customer":
            return str(item.get("number") or "").strip()
    return ""


def _money_cell(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        from decimal import Decimal, ROUND_HALF_UP

        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return None


def _debit_from_row(row: dict):
    if not isinstance(row, dict):
        return None
    for key in (
        "debit",
        "current_debit",
        "debit_amt",
        "debitAmount",
        "period_debit",
        "fcyamt",
        "allDebit",
        "ytd_debit",
    ):
        if row.get(key) is not None and str(row.get(key)).strip() != "":
            try:
                from decimal import Decimal, ROUND_HALF_UP

                return Decimal(str(row.get(key))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            except Exception:
                continue
    return None


def _shift_yyyymm(yyyymm: str, months: int) -> str:
    y = int(yyyymm[:4])
    m = int(yyyymm[4:6]) + months
    while m <= 0:
        m += 12
        y -= 1
    while m > 12:
        m -= 12
        y += 1
    return f"{y}{m:02d}"


def ar_windows(period: str) -> dict[str, tuple[str, str]]:
    """余额先看近一个月（含开票月），没有这家再拼到当年1月。本期借方仍是开票月。"""
    end = _yyyymm(period)
    if not end:
        return {"recent": ("", ""), "ytd": ("", "")}
    year_start = end[:4] + "01"
    prev = _shift_yyyymm(end, -1)
    recent_start = prev if prev >= year_start else year_start
    return {"recent": (recent_start, end), "ytd": (year_start, end)}


def _ar_window(period: str) -> tuple[str, str]:
    return ar_windows(period)["recent"]


def _window_key(start: str, end: str) -> str:
    return f"{start}:{end}"


def _save_ar_cache(payload: dict) -> None:
    path = ar_cache_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        existing: dict = {}
        if path.is_file():
            try:
                existing = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError, TypeError):
                existing = {}
        accounts = list(payload.get("accounts") or [])
        windows = existing.get("windows") if list(existing.get("accounts") or []) == accounts else None
        if not isinstance(windows, dict):
            windows = {}
            old_start = str(existing.get("start") or "")
            old_end = str(existing.get("end") or "")
            if old_start and old_end and list(existing.get("accounts") or []) == accounts:
                windows[_window_key(old_start, old_end)] = {
                    "balances": existing.get("balances") or [],
                    "debits": existing.get("debits") or [],
                }
        windows[_window_key(str(payload["start"]), str(payload["end"]))] = {
            "balances": payload.get("balances") or [],
            "debits": payload.get("debits") or [],
        }
        path.write_text(
            json.dumps(
                {"cached_at": time.time(), "accounts": accounts, "windows": windows},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        path.chmod(0o600)
    except OSError:
        return


def _load_ar_cache(start: str, end: str, accounts: list[str]) -> dict | None:
    path = ar_cache_path()
    if not path.is_file():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        cached_at = float(payload["cached_at"])
    except (OSError, ValueError, KeyError, TypeError):
        return None
    if time.time() - cached_at > cache_ttl_seconds():
        return None
    if list(payload.get("accounts") or []) != list(accounts):
        return None
    windows = payload.get("windows")
    if isinstance(windows, dict):
        block = windows.get(_window_key(start, end))
        if not isinstance(block, dict):
            return None
        return {
            "start": start,
            "end": end,
            "accounts": accounts,
            "balances": block.get("balances") or [],
            "debits": block.get("debits") or [],
        }
    if payload.get("start") != start or payload.get("end") != end:
        return None
    return payload


def _scan_voucher_ar(creds: dict, token: str, accounts: list[str], start: str, end: str) -> dict:
    from decimal import Decimal

    extra = {"app-token": token}
    wanted = {str(a).strip() for a in accounts if str(a).strip()}
    zero = Decimal("0.00")
    balances: dict[tuple[str, str], Decimal] = {}
    period_debit: dict[tuple[str, str, str], Decimal] = {}
    page = 1
    seen: set[str] = set()
    total = None
    while page <= 200:
        resp = _request(
            "GET",
            API_HOST + "/jdy/v2/fi/voucher",
            creds,
            "/jdy/v2/fi/voucher",
            params={"start_period": start, "end_period": end, "page": str(page), "page_size": "100"},
            extra_headers=extra,
        )
        if resp.status_code != 200:
            raise RuntimeError(f"voucher list http {resp.status_code}")
        payload = resp.json()
        if total is None and isinstance(payload, dict) and isinstance(payload.get("data"), dict):
            try:
                total = int(payload["data"].get("count") or 0)
            except (TypeError, ValueError):
                total = 0
        rows = _rows_from(payload)
        if not rows:
            break
        new_rows = 0
        for head in rows:
            vid = str((head or {}).get("id") or "").strip()
            if not vid or vid in seen:
                continue
            seen.add(vid)
            new_rows += 1
            detail = _request(
                "GET",
                API_HOST + "/jdy/v2/fi/voucher_detail",
                creds,
                "/jdy/v2/fi/voucher_detail",
                params={"id": vid},
                extra_headers=extra,
            )
            if detail.status_code != 200:
                continue
            try:
                body = detail.json()
            except Exception:
                continue
            data = body.get("data") if isinstance(body, dict) else None
            entries = (data or {}).get("entry_list") if isinstance(data, dict) else None
            raw_period = ""
            if isinstance(data, dict):
                raw_period = str(data.get("period") or "")
            if not raw_period:
                raw_period = str((head or {}).get("period") or "")
            voucher_period = _yyyymm(raw_period)
            month_key = f"{voucher_period[:4]}-{voucher_period[4:6]}" if len(voucher_period) == 6 else ""
            if not isinstance(entries, list):
                continue
            for entry in entries:
                if not isinstance(entry, dict):
                    continue
                acc = str(entry.get("account_number") or "").strip()
                if acc not in wanted:
                    continue
                cus = _customer_code_from_assist(entry.get("assist"))
                if not cus:
                    continue
                debit = _money_cell(entry.get("debit_amount")) or zero
                credit = _money_cell(entry.get("credit_amount")) or zero
                net = debit - credit
                key = (cus, acc)
                balances[key] = (balances.get(key) or zero) + net
                if month_key and debit:
                    dkey = (cus, acc, month_key)
                    period_debit[dkey] = (period_debit.get(dkey) or zero) + debit
        if new_rows == 0 or len(rows) < 100 or (total and len(seen) >= total):
            break
        page += 1
    return {"balances": balances, "period_debit": period_debit, "start": start, "end": end, "accounts": list(wanted)}


def _ledger_for(accounts: list[str], start: str, end: str) -> dict:
    if not start or not end:
        return {"ok": False, "missing_credentials": False, "error": "period", "balances": {}, "period_debit": {}}
    accs = [str(a).strip() for a in accounts if str(a).strip()]
    cached = _load_ar_cache(start, end, accs)
    if cached is not None:
        balances = {}
        for item in cached.get("balances") or []:
            balances[(str(item["c"]), str(item["a"]))] = _money_cell(item["b"]) or 0
        debit = {}
        for item in cached.get("debits") or []:
            debit[(str(item["c"]), str(item["a"]), str(item["p"]))] = _money_cell(item["d"]) or 0
        return {"ok": True, "missing_credentials": False, "balances": balances, "period_debit": debit}
    creds = load_local()
    if not creds:
        return {"ok": False, "missing_credentials": True, "balances": {}, "period_debit": {}}
    try:
        token, _domain = get_app_token(creds)
        scanned = _scan_voucher_ar(creds, token, accs, start, end)
    except Exception as e:
        return {"ok": False, "missing_credentials": False, "error": str(e)[:200], "balances": {}, "period_debit": {}}
    dump_bals = [{"c": c, "a": a, "b": str(v)} for (c, a), v in scanned["balances"].items()]
    dump_deb = [{"c": c, "a": a, "p": p, "d": str(v)} for (c, a, p), v in scanned["period_debit"].items()]
    _save_ar_cache({"start": start, "end": end, "accounts": accs, "balances": dump_bals, "debits": dump_deb})
    return {"ok": True, "missing_credentials": False, "balances": scanned["balances"], "period_debit": scanned["period_debit"]}


def _slice_customer(ledger: dict, cus: str, month: str) -> tuple[dict, dict]:
    balances = {}
    for (code, acc), val in (ledger.get("balances") or {}).items():
        if code == cus and val:
            balances[(code, acc)] = val
    period_debit = {}
    for (code, acc, per), val in (ledger.get("period_debit") or {}).items():
        if code == cus and per == month and val:
            period_debit[(code, acc, per)] = val
    return balances, period_debit


def try_fetch_period_debit(customer_code: str, accounts: list[str], period: str) -> dict:
    """本期借方。科目余额表现网拒 YYYYMM，改从凭证分录 assist 汇总。"""
    got = try_fetch_customer_ar(customer_code, accounts, period)
    return {
        "ok": bool(got.get("ok")),
        "missing_credentials": bool(got.get("missing_credentials")),
        "error": got.get("error"),
        "data": got.get("period_debit") or {},
    }


def try_fetch_customer_ar(customer_code: str, accounts: list[str], period: str) -> dict:
    """余额先近一个月（含开票月）凭证净额；没有这家再拼到当年1月。本期借方=开票月借方。"""
    cus = str(customer_code or "").strip()
    wins = ar_windows(period)
    recent_start, recent_end = wins["recent"]
    ytd_start, ytd_end = wins["ytd"]
    month = str(period or "")[:7]
    if len(month) == 6 and month.isdigit():
        month = f"{month[:4]}-{month[4:6]}"
    elif len(month) >= 7:
        month = month[:7]
    recent = _ledger_for(accounts, recent_start, recent_end)
    if not recent.get("ok"):
        return {
            "ok": False,
            "missing_credentials": bool(recent.get("missing_credentials")),
            "error": recent.get("error"),
            "balances": {},
            "period_debit": {},
            "data": {},
        }
    balances, period_debit = _slice_customer(recent, cus, month)
    if balances or period_debit:
        return {
            "ok": True,
            "missing_credentials": False,
            "balances": balances,
            "period_debit": period_debit,
            "data": period_debit,
        }
    if (ytd_start, ytd_end) == (recent_start, recent_end):
        return {
            "ok": True,
            "missing_credentials": False,
            "balances": {},
            "period_debit": {},
            "data": {},
        }
    ytd = _ledger_for(accounts, ytd_start, ytd_end)
    if not ytd.get("ok"):
        return {
            "ok": False,
            "missing_credentials": bool(ytd.get("missing_credentials")),
            "error": ytd.get("error"),
            "balances": {},
            "period_debit": {},
            "data": {},
        }
    balances, period_debit = _slice_customer(ytd, cus, month)
    return {
        "ok": True,
        "missing_credentials": False,
        "balances": balances,
        "period_debit": period_debit,
        "data": period_debit,
    }


def _voucher_no(row) -> int | None:
    if not isinstance(row, dict):
        return None
    raw = row.get("number")
    if raw is None or str(raw).strip() == "":
        raw = row.get("bill_no") or row.get("voucher_no")
    if raw is None or str(raw).strip() == "":
        return None
    s = str(raw).strip().replace("记", "").replace("－", "-")
    if s.startswith("-"):
        s = s[1:]
    s = s.replace("-", "")
    if s.isdigit():
        return int(s)
    try:
        return int(float(s))
    except (TypeError, ValueError):
        return None


def try_fetch_next_voucher_no(period: str | None = None) -> dict:
    """当前月凭证最大 number + 1。没说起始号时用这个，不得默认从 1 起。"""
    from datetime import date as date_cls

    yyyymm = _yyyymm(period or "") or date_cls.today().strftime("%Y%m")
    if not yyyymm:
        return {"ok": False, "missing_credentials": False, "error": "period"}
    creds = load_local()
    if not creds:
        return {"ok": False, "missing_credentials": True, "error": "no credentials"}
    try:
        token, _domain = get_app_token(creds)
        extra = {"app-token": token}
        max_no = 0
        page = 1
        while page <= 80:
            resp = _request(
                "GET",
                API_HOST + "/jdy/v2/fi/voucher",
                creds,
                "/jdy/v2/fi/voucher",
                params={
                    "start_period": yyyymm,
                    "end_period": yyyymm,
                    "page": str(page),
                    "page_size": "100",
                },
                extra_headers=extra,
            )
            if resp.status_code != 200:
                return {
                    "ok": False,
                    "missing_credentials": False,
                    "error": f"voucher list http {resp.status_code}",
                }
            payload = resp.json()
            rows = _rows_from(payload)
            for row in rows:
                n = _voucher_no(row)
                if n is not None and n > max_no:
                    max_no = n
            if len(rows) < 100:
                break
            page += 1
        return {
            "ok": True,
            "missing_credentials": False,
            "next_number": max_no + 1,
            "max_number": max_no,
            "period": yyyymm,
        }
    except Exception as e:
        return {
            "ok": False,
            "missing_credentials": False,
            "error": f"{type(e).__name__}: {str(e)[:200]}",
        }
