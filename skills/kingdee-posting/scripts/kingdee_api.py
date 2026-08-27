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
NEED_KEYS = ("client_id", "client_secret", "app_key", "app_secret")
MASTER_KINDS = ("customer", "employee", "supplier", "department")
MASTER_CACHE_TTL_SECONDS = 15 * 60


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


def get_app_token(creds: dict) -> tuple[str, str]:
    params = {
        "app_key": str(creds["app_key"]),
        "app_signature": app_signature(str(creds["app_key"]), str(creds["app_secret"])),
    }
    resp = _request("GET", API_HOST + AUTH_PATH, creds, AUTH_PATH, params=params)
    if resp.status_code != 200:
        raise RuntimeError(f"auth http {resp.status_code}")
    payload = resp.json()
    data = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(data, dict):
        raise RuntimeError("auth missing data")
    token = data.get("app-token") or data.get("app_token") or data.get("appToken")
    domain = data.get("domain") or creds.get("domain") or ""
    if not token:
        raise RuntimeError("auth missing app-token")
    return str(token), str(domain).rstrip("/")


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
