#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""智云只读：下单业务线、回款/下单销售。密钥只读本机文件，不进 git、不打印。"""
from __future__ import annotations

import importlib.util
import os
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
import lookups as lookup_mod  # noqa: E402

DEFAULT_LOCAL = Path.home() / ".config" / "finance" / "zhiyun.local.json"
DEFAULT_BASE = "http://192.168.10.167:18880"
DEFAULT_APP_ID = "6ff4fb2e-e68c-4ee9-83a0-836de8f72c11"
ORDERS_WS = "6501688ebf25d7b91abdb465"
RECEIPTS_WS = "6555d2b1f9460e517040ba6c"
ORDER_FETCH = (
    Path(__file__).resolve().parents[2] / "order-daily-summary" / "scripts" / "fetch_orders.py"
)

CUSTOMER_FIELDS = ("客户", "客户名称", "开票客户")
LINE_FIELDS = ("业务线",)
SALES_FIELDS = ("销售", "销售名称")
DATE_FIELDS = ("到账日期", "日期", "回款日期")
AMOUNT_FIELDS = ("到账金额/本币", "到账金额/原币", "金额")


def local_path() -> Path:
    override = os.environ.get("ZHIYUN_LOCAL_JSON", "").strip()
    return Path(override) if override else DEFAULT_LOCAL


def load_local() -> dict | None:
    path = local_path()
    if not path.is_file():
        return None
    try:
        import json

        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    has_login = bool(str(data.get("username") or "").strip() and str(data.get("password") or "").strip())
    has_token = bool(str(data.get("md_pss_id") or "").strip())
    if not (has_login or has_token):
        return None
    return data


def _fetch_mod():
    spec = importlib.util.spec_from_file_location("kingdee_posting_order_fetch", ORDER_FETCH)
    if spec is None or spec.loader is None:
        raise RuntimeError("找不到九点下单的智云登录脚本")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _first(rec: dict, names: tuple[str, ...]) -> str:
    for name in names:
        val = str(rec.get(name) or "").strip()
        if val:
            return val
    return ""


def split_vals(raw: str) -> list[str]:
    out = []
    for part in re.split(r"[/,，、|]", str(raw or "")):
        item = part.strip()
        if item and item not in out:
            out.append(item)
    return out


def records_to_lookups(order_records, receipt_records) -> dict:
    customer_lines: dict[str, list[str]] = {}
    order_sales: dict[str, list[str]] = {}
    for rec in order_records or []:
        name = _first(rec, CUSTOMER_FIELDS)
        if not name:
            continue
        lines = customer_lines.setdefault(name, [])
        for line in split_vals(_first(rec, LINE_FIELDS)):
            if line not in lines:
                lines.append(line)
        sales = order_sales.setdefault(name, [])
        for person in split_vals(_first(rec, SALES_FIELDS)):
            if person not in sales:
                sales.append(person)
    grouped: dict[tuple[str, str, str], dict] = {}
    for rec in receipt_records or []:
        name = _first(rec, CUSTOMER_FIELDS)
        day = _first(rec, DATE_FIELDS).replace("/", "-").replace(".", "-")[:10]
        amount = _first(rec, AMOUNT_FIELDS)
        key_amt = lookup_mod.amt_key(amount)
        if not (name and len(day) == 10 and key_amt):
            continue
        item = grouped.setdefault(
            (lookup_mod.norm_name(name), day, key_amt),
            {"customer": name, "date": day, "amount": amount, "sales": []},
        )
        for person in split_vals(_first(rec, SALES_FIELDS)):
            if person not in item["sales"]:
                item["sales"].append(person)
    return {
        "customer_lines": customer_lines,
        "order_sales": order_sales,
        "receipt_sales": list(grouped.values()),
        "period_debit": [],
    }


def _worksheet_cfg(cfg: dict) -> dict:
    return {
        "base_url": str(cfg.get("base_url") or DEFAULT_BASE).rstrip("/"),
        "app_id": str(cfg.get("app_id") or DEFAULT_APP_ID),
        "username": str(cfg.get("username") or ""),
        "password": str(cfg.get("password") or ""),
        "md_pss_id": str(cfg.get("md_pss_id") or ""),
        "account_id": str(cfg.get("account_id") or ""),
        "orders_worksheet_id": str(cfg.get("orders_worksheet_id") or ORDERS_WS),
        "receipts_worksheet_id": str(cfg.get("receipts_worksheet_id") or RECEIPTS_WS),
    }


def _fetch_sheet(fo, post, worksheet_id: str, app_id: str) -> list[dict]:
    controls = fo.get_controls(post, worksheet_id, app_id)
    raw = fo.fetch_all_rows(post, worksheet_id, app_id, [])
    return fo.rows_to_records(raw, controls)


def _login_browser(zy: dict):
    """优先本机 Chrome，不再 download Chromium。"""
    from playwright.sync_api import sync_playwright

    base = zy.get("base_url") or DEFAULT_BASE
    user = zy.get("username")
    pwd = zy.get("password")
    if not (base and user and pwd):
        raise RuntimeError("智云配置缺账号")
    with sync_playwright() as p:
        try:
            br = p.chromium.launch(channel="chrome", headless=True)
        except Exception:
            br = p.chromium.launch(headless=True)
        try:
            ctx = br.new_context(ignore_https_errors=True)
            pg = ctx.new_page()
            pg.goto(base, wait_until="networkidle", timeout=30000)
            pg.fill("#txtMobilePhone", user)
            pg.fill("input[type=password]", pwd)
            clicked = False
            for sel in ("text=登 录", "text=登录", ".loginBtn"):
                try:
                    pg.click(sel, timeout=2500)
                    clicked = True
                    break
                except Exception:
                    continue
            if not clicked:
                pg.keyboard.press("Enter")
            pg.wait_for_timeout(6000)
            token = None
            for c in ctx.cookies():
                if c["name"] == "md_pss_id" and c.get("value"):
                    token = c["value"]
                    break
            if not token:
                raise RuntimeError("登录后未取到 md_pss_id")
            account_id = None
            try:
                account_id = pg.evaluate(
                    "() => { try { return md.global.Account.accountId || null; } catch(e) { return null; } }"
                )
            except Exception:
                account_id = None
            return token, account_id
        finally:
            br.close()


def try_load_lookups() -> dict:
    cfg = load_local()
    if not cfg:
        return {"ok": False, "missing_credentials": True, "data": None}
    try:
        fo = _fetch_mod()
        zy = _worksheet_cfg(cfg)
        if not zy["md_pss_id"]:
            try:
                token, account_id = _login_browser(zy)
            except Exception:
                token, account_id = fo.login(zy)
            zy["md_pss_id"] = token
            if account_id:
                zy["account_id"] = account_id
        post = fo._make_post(zy)
        orders = _fetch_sheet(fo, post, zy["orders_worksheet_id"], zy["app_id"])
        receipts = _fetch_sheet(fo, post, zy["receipts_worksheet_id"], zy["app_id"])
        return {
            "ok": True,
            "missing_credentials": False,
            "source": "live",
            "data": records_to_lookups(orders, receipts),
        }
    except Exception as e:
        return {
            "ok": False,
            "missing_credentials": False,
            "error": f"{type(e).__name__}: {str(e)[:200]}",
            "data": None,
        }
