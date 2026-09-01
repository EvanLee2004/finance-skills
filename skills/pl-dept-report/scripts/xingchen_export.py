#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""本机星辰引出助手：切账套 → 科目余额表 / 核算项目余额表 / 利润表 → xlsx。

只给本机取数用，不是同事必装依赖。口令只读本机文件，不打印、不进 git。
禁止点引入 / 审核 / 过账 / 购买。优先复用已登录的斯佳 Chrome 配置目录。
"""
from __future__ import annotations

import argparse
import asyncio
import re
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from layout import load_books

HOME = Path.home()
STATE = HOME / ".config" / "finance" / "xingchen.playwright-state.json"
PROFILE = HOME / ".cache" / "chrome-devtools-mcp" / "chrome-profile"
SIJIA_MD = Path(
    "/Users/evanlee/Documents/甲骨易实习/项目/长期项目/自动化记账（金蝶）/原始素材/实操批次/20260826_金蝶斯佳账号_内网勿外传.md"
)
WORKBENCH = "https://service.jdy.com/workbench/web/index.html"
XINGCHEN_HOME = "https://tf.jdy.com/ierp/index.html?formId=home_page"
ASSIST_FORM = "https://tf.jdy.com/ierp/index.html?formId=gl_rpt_assistbalance"
ACCT_FORM = "https://tf.jdy.com/ierp/index.html?formId=gl_rpt_acctbalance"
PROFIT_FORM = "https://tf.jdy.com/ierp/index.html?formId=cs_ifs_reportdata_list"
EXPORT_LOG = "https://tf.jdy.com/ierp/index.html?formId=bos_exportlog_list"
DOWNLOADS = HOME / "Downloads"


def _export_books() -> list[tuple[str, str, str]]:
    data = load_books()
    out = []
    for acc in data.get("xingchen_accounts") or []:
        if acc.get("api_ready_default"):
            continue
        key = str(acc.get("key") or "")
        legal = str(acc.get("legal_name") or "")
        header = str(acc.get("excel_header") or key)
        if key and legal:
            out.append((key, legal, header))
    return out


def _sijia_login() -> tuple[str, str]:
    if not SIJIA_MD.is_file():
        raise SystemExit("missing_sijia_file")
    text = SIJIA_MD.read_text(encoding="utf-8")
    user = ""
    password = ""
    for line in text.splitlines():
        if "登录名" in line and "`" in line:
            m = re.search(r"`([^`]+)`", line)
            if m:
                user = m.group(1).strip()
        if ("密码" in line or "口令" in line) and "`" in line and "服务密码" not in line:
            m = re.search(r"`([^`]+)`", line)
            if m and m.group(1).strip() not in {"u13439472096"}:
                password = m.group(1).strip()
    if not user or not password:
        raise SystemExit("missing_sijia_fields")
    return user, password


async def _ensure_login(page) -> None:
    await page.goto(WORKBENCH, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)
    if "workbench" in page.url and await page.get_by_text("进入使用").count():
        await page.get_by_text("进入使用").first.click()
        await page.wait_for_timeout(3000)
        return
    user, password = _sijia_login()
    user_box = page.locator('input[type="text"], input[type="tel"]').first
    pwd_box = page.locator('input[type="password"]').first
    if await user_box.count():
        await user_box.fill(user)
        await pwd_box.fill(password)
        await page.get_by_text("登录", exact=True).first.click()
        await page.wait_for_timeout(4000)
    if await page.get_by_text("进入使用").count():
        await page.get_by_text("进入使用").first.click()
        await page.wait_for_timeout(4000)


async def _dismiss_overlays(page) -> None:
    await page.evaluate(
        """() => {
          document.querySelectorAll('iframe').forEach(f => {
            const src = f.src || '';
            if (/piaozone|fpdk|etax/.test(src)) (f.closest('div') || f).remove();
          });
        }"""
    )


async def _click_text(page, label: str) -> bool:
    return bool(
        await page.evaluate(
            """(label) => {
              const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
              let n;
              while ((n = walker.nextNode())) {
                if ((n.textContent || '').trim() === label) {
                  const el = n.parentElement;
                  (el?.closest('label,li,button,div') || el)?.click();
                  return true;
                }
              }
              return false;
            }""",
            label,
        )
    )


async def _switch_book(page, legal_name: str) -> bool:
    await _dismiss_overlays(page)
    await page.goto(XINGCHEN_HOME, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)
    await _dismiss_overlays(page)
    current = await page.locator("body").inner_text()
    if legal_name in current and current.find(legal_name) < 8000:
        return True
    handle = page.get_by_text("甲骨易（北京）语言科技股份有限公司").first
    if await handle.count():
        await handle.click()
        await page.wait_for_timeout(800)
    target = page.get_by_text(legal_name, exact=False).first
    if not await target.count():
        return False
    await target.click()
    await page.wait_for_timeout(2500)
    return legal_name in await page.locator("body").inner_text()


async def _set_acct_filters(page) -> None:
    await _click_text(page, "展开过滤")
    await page.wait_for_timeout(400)
    await _click_text(page, "展开所有级次")
    await _click_text(page, "显示所有科目")
    await _click_text(page, "查询")
    await page.wait_for_timeout(2500)


async def _set_assist_dept(page) -> None:
    box = page.locator(".kd-table-cell-basedata-container").first
    if await box.count():
        await box.click()
        await page.keyboard.type("部门")
        await page.wait_for_timeout(800)
        if await page.get_by_text("部门", exact=True).count():
            await page.get_by_text("部门", exact=True).last.click()
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(400)
    await _click_text(page, "查询")
    await page.wait_for_timeout(2500)


def _newest_xlsx(folder: Path, after: float) -> Path | None:
    newest = None
    newest_mtime = after
    if not folder.is_dir():
        return None
    for path in folder.glob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        try:
            mtime = path.stat().st_mtime
        except OSError:
            continue
        if mtime > newest_mtime:
            newest = path
            newest_mtime = mtime
    return newest


async def _export_current(page, dest: Path) -> bool:
    dest.parent.mkdir(parents=True, exist_ok=True)
    mark = time.time()
    try:
        async with page.expect_download(timeout=20000) as dl_info:
            await _click_text(page, "引出")
            extra = page.get_by_text("Excel", exact=False)
            if await extra.count():
                await extra.first.click()
        download = await dl_info.value
        await download.save_as(str(dest))
        if dest.is_file() and dest.stat().st_size > 400:
            return True
    except Exception:
        await _click_text(page, "引出")
        await page.wait_for_timeout(4000)
    found = _newest_xlsx(DOWNLOADS, mark - 1)
    if found:
        dest.write_bytes(found.read_bytes())
        return dest.is_file() and dest.stat().st_size > 400
    return dest.is_file() and dest.stat().st_size > 400


async def export_book(page, key: str, legal: str, header: str, period: str, out_dir: Path) -> dict:
    note = {"book": key, "header": header, "ok": False, "files": []}
    if not await _switch_book(page, legal):
        note["reason"] = "switch_failed"
        return note
    reports = [
        ("account", ACCT_FORM, out_dir / f"{header}_科目余额表.xlsx", _set_acct_filters),
        ("assist", ASSIST_FORM, out_dir / f"{header}_核算项目余额表.xlsx", _set_assist_dept),
        ("profit", PROFIT_FORM, out_dir / f"{header}_利润表.xlsx", None),
    ]
    for kind, url, path, setup in reports:
        try:
            await _dismiss_overlays(page)
            await page.goto(url, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
            await _dismiss_overlays(page)
            if setup:
                await setup(page)
            else:
                await _click_text(page, "查询")
                await page.wait_for_timeout(2000)
            body = await page.locator("body").inner_text()
            if "暂无数据" in body and kind != "profit":
                note.setdefault("empty", []).append(kind)
                continue
            if await _export_current(page, path):
                note["files"].append(path.name)
            else:
                note.setdefault("empty", []).append(kind)
        except Exception as e:
            note.setdefault("errors", []).append(f"{kind}:{type(e).__name__}")
    note["ok"] = bool(note["files"])
    return note


async def _open_context(p):
    if PROFILE.is_dir():
        try:
            ctx = await p.chromium.launch_persistent_context(
                user_data_dir=str(PROFILE),
                channel="chrome",
                headless=False,
                accept_downloads=True,
            )
            return ctx, None, "profile"
        except Exception:
            pass
    browser = await p.chromium.launch(headless=False, channel="chrome")
    ctx = await browser.new_context(
        accept_downloads=True,
        storage_state=str(STATE) if STATE.is_file() else None,
    )
    return ctx, browser, "fresh"


async def main_async(period: str, out_dir: Path, keys: list[str]) -> int:
    from playwright.async_api import async_playwright

    out_dir.mkdir(parents=True, exist_ok=True)
    STATE.parent.mkdir(parents=True, exist_ok=True)
    wanted = set(keys)
    catalog = [(k, legal, header) for k, legal, header in _export_books() if k in wanted]
    async with async_playwright() as p:
        ctx, browser, mode = await _open_context(p)
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        if mode != "profile":
            await _ensure_login(page)
            await ctx.storage_state(path=str(STATE))
            STATE.chmod(0o600)
        results = []
        for key, legal, header in catalog:
            results.append(await export_book(page, key, legal, header, period, out_dir))
        await ctx.close()
        if browser:
            await browser.close()
    ok = sum(1 for r in results if r.get("ok"))
    print(f"period={period} books_ok={ok}/{len(results)} mode={mode}")
    for r in results:
        extra = r.get("reason") or ",".join(r.get("errors") or r.get("empty") or [])
        print(
            f"book={r['book']} ok={int(bool(r.get('ok')))} files={len(r.get('files') or [])} {extra}".strip()
        )
    return 0 if ok else 2


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--period", default="202608")
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--books", default="wenhua,shanghai,hunan_fgs,hunan_zgs")
    args = parser.parse_args()
    keys = [k.strip() for k in args.books.split(",") if k.strip()]
    return asyncio.run(main_async(args.period, Path(args.out_dir).expanduser(), keys))


if __name__ == "__main__":
    raise SystemExit(main())
