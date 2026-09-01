#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金蝶星辰只读签章。复制收窄自入账技能，不 import 入账夹，不改 pick_authorize_row。"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import time
import uuid
from pathlib import Path
from urllib.parse import quote

import requests

from common import local_json_path, load_optional_json

API_HOST = "https://api.kingdee.com"
AUTH_PATH = "/jdyconnector/app_management/kingdee_auth_token"
AUTHORIZE_PATH = "/jdyconnector/app_management/push_app_authorize"
EMPTY_SERVICE = "7914379139221"
HQ_SERVICE = "795589109148"
HQ_ACCOUNT = "1783803326505631821"
NEED_KEYS = ("client_id", "client_secret", "app_key", "app_secret")
DEFAULT_LOCAL = Path.home() / ".config" / "finance" / "kingdee.local.json"
DEFAULT_PL = Path.home() / ".config" / "finance" / "kingdee-pl.local.json"


def client_path() -> Path:
    return local_json_path("KINGDEE_LOCAL_JSON", DEFAULT_LOCAL)


def pl_path() -> Path:
    return local_json_path("KINGDEE_PL_LOCAL_JSON", DEFAULT_PL)


def load_client() -> dict | None:
    data = load_optional_json(client_path())
    if not data:
        return None
    if any(not str(data.get(k) or "").strip() for k in NEED_KEYS):
        return None
    return data


def hmac_sha256_hex_b64(secret: str, message: str) -> str:
    digest_hex = hmac.new(secret.encode("utf-8"), message.encode("utf-8"), hashlib.sha256).hexdigest()
    return base64.b64encode(digest_hex.encode("utf-8")).decode("ascii")


def app_signature(app_key: str, app_secret: str) -> str:
    return hmac_sha256_hex_b64(app_secret, app_key)


def double_encode(value: str) -> str:
    return quote(quote(str(value), safe=""), safe="")


def sign_plain(method: str, path: str, params: dict | None, nonce: str, timestamp: str) -> str:
    encoded_params = ""
    if params:
        encoded_params = "&".join(f"{double_encode(k)}={double_encode(v)}" for k, v in sorted(params.items()))
    headers_block = f"x-api-nonce:{nonce}\nx-api-timestamp:{timestamp}"
    return f"{method.upper()}\n{quote(path, safe='')}\n{encoded_params}\n{headers_block}\n"


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
        "X-Api-Signature": x_api_signature(str(creds["client_secret"]), method, path, params, nonce, timestamp),
    }


def request(method: str, path: str, creds: dict, params: dict | None = None, extra: dict | None = None, timeout: int = 30, retries: int = 4):
    headers = _auth_headers(creds, method, path, params)
    if extra:
        headers.update(extra)
    last_err = None
    attempts = max(1, retries)
    for attempt in range(attempts):
        try:
            return requests.request(method, API_HOST + path, headers=headers, params=params, timeout=timeout)
        except (requests.ConnectionError, requests.Timeout) as e:
            last_err = e
            time.sleep(0.3 * (attempt + 1))
    raise last_err or RuntimeError("request failed")


def official_account_ids() -> set[str]:
    ids = {HQ_ACCOUNT}
    pl = load_optional_json(pl_path()) or {}
    for row in (pl.get("accounts") or {}).values():
        if isinstance(row, dict) and row.get("account_id"):
            ids.add(str(row["account_id"]).strip())
    return ids


def pick_pl_authorize_rows(rows: list, creds: dict) -> list[dict]:
    """只接受 5 本正式星辰账；空账残留丢掉。不替代入账 picker。"""
    live = [row for row in rows if isinstance(row, dict) and str(row.get("status")).split(".")[0] == "1"]
    official = official_account_ids()
    picked = []
    for row in live:
        svc = str(row.get("serviceId") or "").strip()
        acc = str(row.get("accountId") or "").strip()
        if svc == EMPTY_SERVICE:
            continue
        if acc and acc in official:
            picked.append(row)
            continue
        if svc == HQ_SERVICE and acc == str(creds.get("account_id") or HQ_ACCOUNT):
            picked.append(row)
    return picked


def _save_local(creds: dict) -> None:
    path = client_path()
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


def refresh_hq_secret(creds: dict) -> dict | None:
    oid = str(creds.get("outer_instance_id") or "").strip()
    if not oid:
        return None
    resp = request("POST", AUTHORIZE_PATH, creds, params={"outerInstanceId": oid})
    if resp.status_code != 200:
        return None
    try:
        payload = resp.json()
    except Exception:
        return None
    rows = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(rows, list):
        return None
    picked = pick_pl_authorize_rows(rows, creds)
    hq = None
    wanted = str(creds.get("account_id") or HQ_ACCOUNT)
    for row in picked:
        if str(row.get("accountId") or "").strip() == wanted:
            hq = row
            break
    if not hq or not str(hq.get("appSecret") or "").strip():
        return None
    updated = dict(creds)
    for src, dst in (
        ("appKey", "app_key"),
        ("appSecret", "app_secret"),
        ("accountId", "account_id"),
        ("accountName", "account_name"),
        ("serviceId", "service_id"),
        ("outerInstanceId", "outer_instance_id"),
        ("domain", "domain"),
    ):
        val = hq.get(src)
        if val is not None and str(val).strip():
            updated[dst] = str(val)
    updated["secret_refreshed_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    _save_local(updated)
    return updated


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
    resp = request("GET", AUTH_PATH, creds, params=params)
    if resp.status_code != 200:
        raise RuntimeError(f"auth http {resp.status_code}")
    payload = resp.json()
    got = _token_from_payload(payload, creds)
    if got:
        return got
    err = payload.get("errcode") if isinstance(payload, dict) else None
    if allow_refresh and str(err) == "1030002006":
        refreshed = refresh_hq_secret(creds)
        if refreshed:
            return get_app_token(refreshed, allow_refresh=False)
    desc = str((payload or {}).get("description") or "")[:80] if isinstance(payload, dict) else ""
    raise RuntimeError(f"auth missing data errcode={err} {desc}".strip())


def rows_from(payload) -> list:
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
