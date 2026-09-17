#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""总部现网：银行科目、供应商建档、凭证号。Playwright 登亮晶号切总部，科目走 OpenAPI。"""
from __future__ import annotations

from pathlib import Path

HERE = Path(__file__).resolve().parent
POSTING_SCRIPTS = HERE.parent.parent / "kingdee-posting" / "scripts"
PL_SCRIPTS = HERE.parent.parent / "pl-dept-report" / "scripts"


def _load_mod(name: str, path: Path):
    import importlib.util
    import sys

    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


kingdee_api = _load_mod("kingdee_payment_live_api", POSTING_SCRIPTS / "kingdee_api.py")

HQ_NAME = "甲骨易（北京）语言科技股份有限公司"
HQ_SEARCH = "语言科技"
SUPPLIER_SPECIAL_CODES = {999, 9999}


def ensure_hq_session() -> str:
    """Playwright 用 xingchen.local.json 登总部。失败 raise SystemExit 人话。"""
    login_path = PL_SCRIPTS / "xingchen_login.py"
    export_path = PL_SCRIPTS / "xingchen_export.py"
    if not login_path.is_file() or not export_path.is_file():
        raise SystemExit("找不到月度损益表技能的登录脚本。请把 pl-dept-report 和本技能装在一起。")
    login = _load_mod("pay_xingchen_login", login_path)
    exporter = _load_mod("pay_xingchen_export", export_path)
    creds = login.load_creds()
    if not creds:
        raise SystemExit(
            "本机没有金蝶网页账密。请把亮晶号写进 ~/.config/finance/xingchen.local.json"
            "（截图在 ~/.config/finance/xingchen_login_liangjing.png）。未生成引入表。"
        )

    async def _run():
        from playwright.async_api import async_playwright

        state = login.state_path()
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            kwargs = {"locale": "zh-CN", "extra_http_headers": {"Accept-Language": "zh-CN,zh;q=0.9"}}
            if state.is_file():
                kwargs["storage_state"] = str(state)
            ctx = await browser.new_context(**kwargs)
            page = await ctx.new_page()
            try:
                await login.ensure_login(page, creds)
            except SystemExit:
                raise
            except Exception as e:
                raise SystemExit(f"金蝶网页现在连不上（{type(e).__name__}）。请检查网络后重跑。") from e
            if await page.get_by_text("进入使用").count():
                await page.get_by_text("进入使用").first.click()
                await page.wait_for_timeout(2000)
            ok = await exporter._switch_book(page, HQ_NAME, HQ_SEARCH)
            try:
                await login.save_state(ctx)
            except Exception:
                pass
            await ctx.close()
            await browser.close()
            return ok

    import asyncio

    ok = asyncio.run(_run())
    if not ok:
        raise SystemExit("金蝶网页登录了但没切到总部。未生成引入表。")
    return "hq"


def fetch_bank_accounts() -> list[dict]:
    creds = kingdee_api.load_local()
    if not creds:
        raise SystemExit("本机没有金蝶开放平台应用号（kingdee.local.json）。未生成引入表。")
    token, _domain = kingdee_api.get_app_token(creds)
    resp = kingdee_api._request(
        "GET",
        kingdee_api.API_HOST + "/jdy/v2/fi/account",
        creds,
        "/jdy/v2/fi/account",
        params={"page": "1", "page_size": "2000"},
        extra_headers={"app-token": token},
    )
    if resp.status_code != 200:
        raise SystemExit(f"会计科目未核验（http {resp.status_code}）。未生成引入表。")
    payload = resp.json()
    rows = ((payload.get("data") or {}).get("rows") if isinstance(payload, dict) else None) or []
    out = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        num = str(r.get("number") or "").strip()
        name = str(r.get("name") or "").strip()
        if not num.startswith("1002"):
            continue
        if not r.get("is_leaf") and num != "1002":
            continue
        out.append({"code": num, "name": name, "full": str(r.get("full_name") or ""), "enable": r.get("enable")})
    return out


def pick_payment_banks(accounts: list[dict], rules: dict) -> tuple[str, str]:
    cfg = (rules or {}).get("payment") or {}
    boc_want = str(cfg.get("boc_bank_account") or "100201")
    citic_want = str(cfg.get("citic_bank_account") or "100206")
    by_code = {str(a.get("code")): a for a in accounts if a.get("code")}
    if boc_want not in by_code:
        raise SystemExit(f"总部会计科目没有中行 {boc_want}。未生成引入表。")
    if citic_want not in by_code:
        raise SystemExit(f"总部会计科目没有中信 {citic_want}。未生成引入表。")
    boc_name = str(by_code[boc_want].get("name") or "")
    citic_name = str(by_code[citic_want].get("name") or "")
    if "中行" not in boc_name and "中国银行" not in boc_name:
        raise SystemExit(f"{boc_want} 现网名称不是中行（{boc_name}）。未生成引入表。")
    if "中信" not in citic_name:
        raise SystemExit(f"{citic_want} 现网名称不是中信（{citic_name}）。未生成引入表。")
    return boc_want, citic_want


def next_supplier_number(suppliers: list) -> str:
    best = 0
    for item in suppliers or []:
        if isinstance(item, dict):
            raw = str(item.get("code") or item.get("number") or "").strip()
        else:
            raw = str(item or "").strip()
        if not raw.isdigit():
            continue
        n = int(raw)
        if n in SUPPLIER_SPECIAL_CODES or n >= 10000:
            continue
        if n > best:
            best = n
    return str(best + 1)


def try_create_supplier(name: str, number: str) -> dict:
    title = str(name or "").strip()
    code = str(number or "").strip()
    if not title or not code:
        return {"ok": False, "error": "缺名称或编码", "number": code, "name": title}
    creds = kingdee_api.load_local()
    if not creds:
        return {"ok": False, "missing_credentials": True, "error": "no credentials", "number": code, "name": title}
    try:
        token, _domain = kingdee_api.get_app_token(creds)
        resp = kingdee_api._request(
            "POST",
            kingdee_api.API_HOST + "/jdy/v2/bd/supplier",
            creds,
            "/jdy/v2/bd/supplier",
            extra_headers={"app-token": token},
            json_body={"name": title, "number": code},
        )
        if resp.status_code != 200:
            return {"ok": False, "error": f"supplier save http {resp.status_code}", "number": code, "name": title}
        payload = resp.json()
        err = payload.get("errcode") if isinstance(payload, dict) else None
        if err not in (None, 0, "0"):
            msg = str(payload.get("description") or payload.get("msg") or payload.get("message") or err)
            return {"ok": False, "error": msg[:200], "number": code, "name": title, "errcode": err}
        return {"ok": True, "number": code, "name": title}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {str(e)[:200]}", "number": code, "name": title}
