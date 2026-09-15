#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from urllib.parse import urlsplit


def _normalized_origin(url: str) -> str:
    parsed = urlsplit(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        return ""
    try:
        port = parsed.port or (443 if parsed.scheme == "https" else 80)
    except ValueError:
        return ""
    return f"{parsed.scheme}://{parsed.hostname.lower()}:{port}"


def _allowed_url(url: str) -> bool:
    if os.environ.get("FINANCIAL_NETWORK_POLICY_REQUIRED") != "1":
        return True
    allowed = {
        item.strip().lower()
        for item in os.environ.get("FINANCIAL_NETWORK_TARGETS", "").split(",")
        if item.strip()
    }
    return (
        os.environ.get("FINANCIAL_NETWORK_ACCESS") == "1"
        and _normalized_origin(url) in allowed
    )


def _assert_allowed_url(url: str) -> None:
    if not _allowed_url(url):
        raise RuntimeError("网络目标不在平台批准的精确目标白名单中。")


def _edge_login(
    base_url: str,
    username: str,
    password: str,
    headless: bool = True,
) -> tuple[str, str | None]:
    import fetch_zhiyun
    from playwright.sync_api import sync_playwright

    try:
        _assert_allowed_url(base_url)
        with sync_playwright() as playwright:
            launch_options: dict[str, object] = {"headless": headless}
            if sys.platform == "win32":
                launch_options["channel"] = "msedge"
            proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
            if proxy_url:
                launch_options["proxy"] = {"server": proxy_url}
            browser = playwright.chromium.launch(**launch_options)
            try:
                context = browser.new_context(ignore_https_errors=True)
                context.route(
                    "**/*",
                    lambda route: route.continue_()
                    if _allowed_url(route.request.url)
                    else route.abort(),
                )
                page = context.new_page()
                page.goto(base_url, wait_until="domcontentloaded", timeout=60_000)
                page.locator("#txtMobilePhone").wait_for(state="visible", timeout=60_000)
                page.fill("#txtMobilePhone", username)
                page.fill("input[type=password]", password)
                clicked = False
                for selector in (".btnForLogin", "text=登 录", "text=登录", ".loginBtn"):
                    try:
                        page.click(selector, timeout=2_500)
                        clicked = True
                        break
                    except Exception:
                        continue
                if not clicked:
                    page.keyboard.press("Enter")
                token = None
                for _ in range(60):
                    token = next(
                        (
                            cookie["value"]
                            for cookie in context.cookies()
                            if cookie.get("name") == "md_pss_id" and cookie.get("value")
                        ),
                        None,
                    )
                    if token:
                        break
                    page.wait_for_timeout(500)
                if not token:
                    raise fetch_zhiyun.LoginError(
                        f"登录后未拿到会话凭据（url={page.url}）。"
                        "请检查账号密码、内网连接或登录页结构。"
                    )
                account_id = None
                try:
                    account_id = page.evaluate(
                        "() => { try { return (md && md.global && md.global.Account && "
                        "md.global.Account.accountId) || null; } catch(e) { return null; } }"
                    )
                except Exception:
                    account_id = None
                return token, account_id
            finally:
                browser.close()
    except fetch_zhiyun.LoginError:
        raise
    except Exception as exc:
        raise fetch_zhiyun.LoginError(
            f"登录异常 {type(exc).__name__}，请检查网络和浏览器运行环境。"
        ) from exc


def main() -> int:
    script_dir = Path(__file__).resolve().parent
    sys.path.insert(0, str(script_dir))
    import fetch_zhiyun

    try:
        payload = json.loads(sys.stdin.read())
        account = str(payload["account"]).strip()
        password = str(payload["password"])
        reconciliation_date = str(payload.get("reconciliation_date") or "").strip()
        date_from = str(payload.get("date_from") or "").strip()
        date_to = str(payload.get("date_to") or "").strip()
        workspace = str(Path(payload["workspace"]).resolve())
        supplement_ar_ids = [
            str(value).strip().upper()
            for value in (payload.get("supplement_ar_ids") or [])
            if str(value).strip()
        ]
        supplement_so_ids = [
            str(value).strip().upper()
            for value in (payload.get("supplement_so_ids") or [])
            if str(value).strip()
        ]
    except (KeyError, TypeError, ValueError):
        print("ERROR: 自动取数凭据输入无效。", file=sys.stderr)
        return 2
    has_single_date = bool(reconciliation_date)
    has_date_range = bool(date_from and date_to)
    if (
        not account
        or not password
        or has_single_date == has_date_range
        or bool(date_from) != bool(date_to)
    ):
        print("ERROR: 自动取数缺少账号、密码或唯一日期范围。", file=sys.stderr)
        return 2
    if has_date_range and (supplement_ar_ids or supplement_so_ids):
        print("ERROR: 按 AR/SO 补取只支持单个核销日。", file=sys.stderr)
        return 2

    fetch_zhiyun.login_with_password = _edge_login
    try:
        try:
            dates = fetch_zhiyun.resolve_fetch_dates(
                single_date=reconciliation_date,
                date_from=date_from,
                date_to=date_to,
            )
        except ValueError as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 2
        for item in dates:
            arguments = [
                "--date",
                item,
                "--workspace",
                workspace,
                "--user",
                account,
                "--password",
                password,
                "--skip-gap-check",
            ]
            for identifier in supplement_ar_ids:
                arguments.extend(["--supplement-ar", identifier])
            for identifier in supplement_so_ids:
                arguments.extend(["--supplement-so", identifier])
            result = fetch_zhiyun.main(arguments)
            if result:
                return result
        return 0
    finally:
        password = ""
        payload.clear()


if __name__ == "__main__":
    raise SystemExit(main())
