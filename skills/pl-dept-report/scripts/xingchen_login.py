#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金蝶工作台网页登录。凭据只读本机 json / 环境变量，会话写 playwright-state。不打印口令。"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

LOGIN_URL = "https://www.jdy.com/login/"
WORKBENCH_URL = "https://service.jdy.com/workbench/web/index.html"
STATE_PATH = Path.home() / ".config" / "finance" / "xingchen.playwright-state.json"
LOCAL_PATH = Path.home() / ".config" / "finance" / "xingchen.local.json"
ASK_CREDS = (
    "本机没有金蝶网页账密。请把用户名密码写进 ~/.config/finance/xingchen.local.json "
    "（可复制技能 config/xingchen.local.example.json），或设环境变量 XINGCHEN_USER / XINGCHEN_PASSWORD。"
)


def state_path() -> Path:
    raw = os.environ.get("XINGCHEN_STATE_JSON", "").strip()
    return Path(raw).expanduser() if raw else STATE_PATH


def local_path() -> Path:
    raw = os.environ.get("XINGCHEN_LOCAL_JSON", "").strip()
    return Path(raw).expanduser() if raw else LOCAL_PATH


def load_creds() -> dict | None:
    user = os.environ.get("XINGCHEN_USER", "").strip()
    password = os.environ.get("XINGCHEN_PASSWORD", "").strip()
    if user and password:
        return {"username": user, "password": password}
    path = local_path()
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    user = str(data.get("username") or "").strip()
    password = str(data.get("password") or "").strip()
    if not user or not password:
        return None
    return {"username": user, "password": password}


def _looks_logged_in(url: str, body: str) -> bool:
    if "login" in url and "workbench" not in url:
        return False
    if "进入使用" in body or "workbench" in url:
        return True
    if "金蝶云星辰" in body or "专业版" in body:
        return True
    return False


async def session_alive(page) -> bool:
    await page.goto(WORKBENCH_URL, wait_until="domcontentloaded", timeout=60000)
    await page.wait_for_timeout(1500)
    url = page.url
    try:
        body = await page.locator("body").inner_text()
    except Exception:
        body = ""
    return _looks_logged_in(url, body[:4000])


async def login_with_password(page, creds: dict) -> None:
    await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    await page.wait_for_timeout(1500)
    tab = page.get_by_text("账号登录", exact=True)
    if await tab.count():
        await tab.first.click()
        await page.wait_for_timeout(400)
    user_box = page.get_by_placeholder("账号", exact=True)
    pwd_box = page.get_by_placeholder("请输入密码", exact=True)
    if not await user_box.count():
        user_box = page.locator('input[placeholder="Account"]')
        pwd_box = page.locator('input[placeholder="Enter Password"]')
    await user_box.first.fill(creds["username"])
    await pwd_box.first.fill(creds["password"])
    box = page.locator('input[type="checkbox"]:visible').first
    if await box.count():
        if not await box.is_checked():
            await box.check()
    btn = page.get_by_role("button", name="登录")
    if not await btn.count():
        btn = page.get_by_role("button", name="Sign in")
    await btn.first.click()
    await page.wait_for_timeout(2500)
    body = ""
    try:
        body = await page.locator("body").inner_text()
    except Exception:
        pass
    if "账号不存在或密码错误" in body or "密码错误" in body:
        raise SystemExit("bad_password")
    agree = page.get_by_text("同意", exact=True)
    if await agree.count():
        await agree.last.click()
        await page.wait_for_timeout(2500)
    try:
        await page.wait_for_url(lambda url: "workbench" in url or "tf.jdy.com" in url, timeout=15000)
    except Exception:
        pass
    if await page.get_by_text("进入使用").count():
        await page.get_by_text("进入使用").first.click()
        await page.wait_for_timeout(2500)


async def ensure_login(page, creds: dict | None = None) -> str:
    if await session_alive(page):
        if await page.get_by_text("进入使用").count():
            await page.get_by_text("进入使用").first.click()
            await page.wait_for_timeout(2000)
        return "session"
    creds = creds or load_creds()
    if not creds:
        raise SystemExit("missing_creds")
    await login_with_password(page, creds)
    if await session_alive(page):
        return "password"
    slider = page.locator("text=滑块")
    sms_box = page.get_by_placeholder("短信验证码")
    if await slider.count() or (await sms_box.count() and await sms_box.first.is_visible()):
        raise SystemExit("need_captcha")
    raise SystemExit("login_failed")


async def save_state(context) -> None:
    path = state_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    await context.storage_state(path=str(path))
    path.chmod(0o600)


async def run_check() -> int:
    from playwright.async_api import async_playwright

    state = state_path()
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        kwargs = {"locale": "zh-CN", "extra_http_headers": {"Accept-Language": "zh-CN,zh;q=0.9"}}
        if state.is_file():
            kwargs["storage_state"] = str(state)
        ctx = await browser.new_context(**kwargs)
        page = await ctx.new_page()
        ok = await session_alive(page)
        print(f"xingchen_login check={'ok' if ok else 'expired'} url_host={page.url.split('/')[2] if page.url else ''}")
        await ctx.close()
        await browser.close()
    return 0 if ok else 2


async def run_login() -> int:
    from playwright.async_api import async_playwright

    creds = load_creds()
    if not creds:
        print("ask=" + ASK_CREDS)
        return 2
    state = state_path()
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        kwargs = {"locale": "zh-CN", "extra_http_headers": {"Accept-Language": "zh-CN,zh;q=0.9"}}
        if state.is_file():
            kwargs["storage_state"] = str(state)
        ctx = await browser.new_context(**kwargs)
        page = await ctx.new_page()
        how = await ensure_login(page, creds)
        await save_state(ctx)
        print(f"xingchen_login how={how} url_host={page.url.split('/')[2] if page.url else ''}")
        await ctx.close()
        await browser.close()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    import asyncio

    if args.check:
        return asyncio.run(run_check())
    return asyncio.run(run_login())


if __name__ == "__main__":
    raise SystemExit(main())
