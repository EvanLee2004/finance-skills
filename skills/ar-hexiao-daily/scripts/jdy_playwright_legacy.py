"""旧版备份：使用 Playwright 完成智云登录与只读取数。

固定流程：官网登录 → 进入使用 → 财务报表 → 现金流量表 → 选择最大期间
→ 查询 → 逐行点击“本月金额” → 调整列表选择本年/全期间 → 全选 → 导出。

脚本不执行重算、调整、重新指定等业务写操作，也不会把用户名或密码写入文件。
终端只输出计数和状态，不输出金额、客户、摘要等财务明细。

示例：
    python scripts/jdy_login.py
    python scripts/jdy_login.py --username your-name
    python scripts/jdy_login.py --login-only --storage-state 工作区/rpa/jdy-state.json

密码不要放在命令行参数中，优先通过交互式输入；无人值守时可使用
JDY_PASSWORD 环境变量。storage state 包含会话凭据，只有显式指定
--storage-state 时才会保存。
"""

from __future__ import annotations

import argparse
import asyncio
import getpass
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence

from playwright.async_api import (
    BrowserContext,
    Frame,
    Locator,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
)


DEFAULT_URL = "http://www.jdy.com/"
DEFAULT_ARTIFACTS_DIR = Path(__file__).resolve().parents[1] / "工作区" / "rpa"
DEFAULT_BROWSER_CHANNEL = "msedge"
DEFAULT_USER_DATA_DIR = DEFAULT_ARTIFACTS_DIR / "edge-profile"
DEFAULT_DOWNLOAD_ROOT = (
    Path(__file__).resolve().parents[1] / "工作区" / "04_产出" / "现金流量调整导出"
)

USERNAME_SELECTORS = (
    'input[name*="user"]',
    'input[id*="user"]',
    'input[name*="account"]',
    'input[id*="account"]',
    'input[placeholder*="用户名"]',
    'input[placeholder*="账号"]',
    'input[placeholder*="手机号"]',
    'input[aria-label*="用户名"]',
    'input[aria-label*="账号"]',
    'input[type="email"]',
    'input[type="tel"]',
    'input[type="text"]',
)

PASSWORD_SELECTORS = (
    'input[type="password"]',
    'input[name*="password"]',
    'input[id*="password"]',
    'input[name*="pwd"]',
    'input[id*="pwd"]',
    'input[placeholder*="密码"]',
    'input[aria-label*="密码"]',
)

LOGIN_BUTTON_SELECTORS = (
    '#login_btn',
    'input[type="button"][id*="login"]:not(#login_btn_gray):not([id*="_gray"])',
    'input[type="button"][value*="登录"]:not(#login_btn_gray):not([id*="_gray"])',
    'input[type="submit"][value*="登录"]',
    'button:has-text("登录")',
    'button:has-text("登陆")',
    'input[type="submit"]',
    '[role="button"]:has-text("登录")',
    '[role="button"]:has-text("登陆")',
)

LOGIN_LINK_SELECTORS = (
    'a[href="/login"]',
    'a[href="/login/"]',
    'a.toplogin',
)

CONSENT_BUTTON_SELECTORS = (
    'a#agree-protocol',
    'a:has-text("我同意")',
    'a:has-text("同意")',
    'button:has-text("我同意")',
    'button:has-text("同意")',
    'input[type="button"][value="我同意"]',
    'input[type="button"][value="同意"]',
    '[role="button"]:has-text("我同意")',
    '[role="button"]:has-text("同意")',
)

LOGIN_ERROR_MARKERS = (
    "用户名或密码错误",
    "账号或密码错误",
    "密码错误",
    "登录失败",
    "请先登录",
)

LOGIN_FORM_POLL_MS = 2_000
ROW_OPEN_TIMEOUT_MS = 3_000
AMOUNT_RE = re.compile(r"^\(?-?[\d,]+(?:\.\d+)?\)?$")
PERIOD_RE = re.compile(r"(?<!\d)(1[0-2]|0?[1-9])\s*期")


@dataclass
class AdjustmentSurface:
    page: Page
    frame: Frame
    container: Locator | None
    opened_new_page: bool


@dataclass
class CashflowTable:
    frame: Frame
    table: Locator
    amount_column_index: int
    row_selector: str
    cell_selector: str


class LoginError(RuntimeError):
    """登录流程无法安全继续。"""


async def first_visible(page: Page, selectors: Iterable[str]) -> Locator | None:
    """按固定候选选择器返回第一个可见元素。"""

    for selector in selectors:
        locator = page.locator(selector)
        try:
            count = await locator.count()
        except Exception:
            continue
        for index in range(min(count, 10)):
            candidate = locator.nth(index)
            try:
                if await candidate.is_visible():
                    return candidate
            except Exception:
                continue
    return None


async def first_visible_in_frame(
    frame: Frame, selectors: Iterable[str]
) -> Locator | None:
    for selector in selectors:
        locator = frame.locator(selector)
        try:
            count = await locator.count()
        except Exception:
            continue
        for index in range(min(count, 20)):
            candidate = locator.nth(index)
            try:
                if await candidate.is_visible():
                    return candidate
            except Exception:
                continue
    return None


async def first_visible_anywhere(
    page: Page, selectors: Iterable[str]
) -> Locator | None:
    for frame in page.frames:
        candidate = await first_visible_in_frame(frame, selectors)
        if candidate is not None:
            return candidate
    return None


async def visible_text_anywhere(
    page: Page, texts: Sequence[str], *, exact: bool = True
) -> tuple[Frame, Locator] | None:
    for frame in page.frames:
        for text in texts:
            locator = frame.get_by_text(text, exact=exact)
            try:
                count = await locator.count()
            except Exception:
                continue
            for index in range(min(count, 30)):
                candidate = locator.nth(index)
                try:
                    if await candidate.is_visible():
                        return frame, candidate
                except Exception:
                    continue
    return None


async def wait_for_visible_text_anywhere(
    page: Page,
    texts: Sequence[str],
    *,
    timeout_ms: int,
    exact: bool = True,
) -> tuple[Frame, Locator]:
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1_000
    while asyncio.get_running_loop().time() < deadline:
        found = await visible_text_anywhere(page, texts, exact=exact)
        if found is not None:
            return found
        await page.wait_for_timeout(300)
    raise LoginError(f"等待页面元素超时：{' / '.join(texts)}")


async def click_and_follow_new_page(
    context: BrowserContext,
    current_page: Page,
    locator: Locator,
    *,
    wait_ms: int = 2_500,
) -> Page:
    before = list(context.pages)
    await locator.click()
    deadline = asyncio.get_running_loop().time() + wait_ms / 1_000
    while asyncio.get_running_loop().time() < deadline:
        new_pages = [page for page in context.pages if page not in before]
        if new_pages:
            new_page = new_pages[-1]
            await new_page.wait_for_load_state("commit", timeout=10_000)
            return new_page
        await current_page.wait_for_timeout(200)
    return current_page


async def body_text(page: Page) -> str:
    try:
        return await page.locator("body").inner_text(timeout=3_000)
    except Exception:
        return ""


async def save_screenshot(page: Page, artifacts_dir: Path, name: str) -> Path:
    artifacts_dir.mkdir(parents=True, exist_ok=True)
    path = artifacts_dir / name
    await page.screenshot(path=str(path), full_page=True, timeout=5_000)
    return path


async def close_browser_safely(browser: object) -> None:
    """异常页面或浏览器进程卡住时，也要让脚本按时退出。"""

    try:
        await asyncio.wait_for(browser.close(), timeout=5)  # type: ignore[attr-defined]
    except Exception:
        # 关闭失败不应覆盖真正的登录/页面错误；由 Playwright 回收进程。
        pass


async def is_authenticated_surface(page: Page) -> bool:
    """识别已登录工作台或可直接进入业务系统的页面。"""

    if "service.jdy.com/workbench" in page.url:
        return True
    return await visible_text_anywhere(page, ("进入使用", "财务报表"), exact=True)


async def wait_for_login_form(
    page: Page,
) -> tuple[Locator | None, Locator | None]:
    """持续等待用户名和密码输入框同时出现，直到用户中断。"""

    while True:
        if await is_authenticated_surface(page):
            return None, None

        username_input = await first_visible(page, USERNAME_SELECTORS)
        password_input = await first_visible(page, PASSWORD_SELECTORS)
        if username_input is not None and password_input is not None:
            return username_input, password_input

        print("正在寻找账号密码输入框", flush=True)
        await page.wait_for_timeout(LOGIN_FORM_POLL_MS)


async def enter_login_page(
    context: BrowserContext, page: Page
) -> Page:
    """持续寻找官网登录按钮；若已在登录页则直接返回。"""

    while True:
        if await is_authenticated_surface(page):
            return page

        username = await first_visible_anywhere(page, USERNAME_SELECTORS)
        password = await first_visible_anywhere(page, PASSWORD_SELECTORS)
        if username is not None and password is not None:
            return page

        login_link = await first_visible_anywhere(page, LOGIN_LINK_SELECTORS)
        if login_link is not None:
            print("已找到登录按钮，正在进入登录界面。", flush=True)
            page = await click_and_follow_new_page(
                context, page, login_link, wait_ms=5_000
            )
            await page.wait_for_timeout(1_500)
            return page

        print("正在寻找登录按钮", flush=True)
        await page.wait_for_timeout(1_000)


async def accept_consent_interstitials(
    page: Page, wait_first_ms: int = 0
) -> None:
    """按用户授权处理“同意/我同意”提示框，直到遮挡消失。"""

    deadline = asyncio.get_running_loop().time() + wait_first_ms / 1_000
    for _ in range(5):
        button = await first_visible_anywhere(page, CONSENT_BUTTON_SELECTORS)
        if button is None:
            if asyncio.get_running_loop().time() < deadline:
                await page.wait_for_timeout(200)
                continue
            return
        print("检测到提示框，正在点击同意。", flush=True)
        await button.click()
        await page.wait_for_timeout(700)


async def wait_for_login_result(page: Page, before_url: str, timeout_ms: int) -> None:
    """等待 URL 或登录表单状态变化，并识别常见登录错误。"""

    deadline = asyncio.get_running_loop().time() + timeout_ms / 1_000
    while asyncio.get_running_loop().time() < deadline:
        if page.url != before_url:
            return

        text = await body_text(page)
        for marker in LOGIN_ERROR_MARKERS:
            if marker in text:
                raise LoginError(f"页面提示登录未完成：{marker}")

        password = await first_visible(page, PASSWORD_SELECTORS)
        if password is None:
            return

        await page.wait_for_timeout(500)

    raise LoginError(
        "点击登录后页面没有发生可确认的变化；请检查账号密码、提示框或登录选择器。"
    )


async def wait_for_manual_login(
    context: BrowserContext, page: Page
) -> Page:
    """等待用户手动勾选协议、点击登录并处理可能出现的同意弹窗。"""

    print(
        "账号密码已填写，请在浏览器中手动勾选协议并点击登录。",
        flush=True,
    )
    rounds = 0
    while True:
        open_pages = [candidate for candidate in context.pages if not candidate.is_closed()]
        for candidate in reversed(open_pages):
            if await is_authenticated_surface(candidate):
                return candidate

        if not open_pages:
            raise LoginError("浏览器页面已关闭。")
        page = open_pages[-1]
        if rounds % 5 == 0:
            print("正在等待人工完成登录", flush=True)
        rounds += 1
        await page.wait_for_timeout(1_000)


async def ensure_login_agreement(page: Page) -> None:
    selectors = (
        "#reg_agreement",
        'input[type="checkbox"][name*="agreement"]',
        'input[type="checkbox"][id*="agreement"]',
    )
    deadline = asyncio.get_running_loop().time() + 5
    while asyncio.get_running_loop().time() < deadline:
        for frame in page.frames:
            v5_container = frame.locator("#v5_agreement")
            v5_visible = (
                await v5_container.count() > 0
                and await v5_container.first.is_visible()
            )
            for selector in selectors:
                candidates = frame.locator(selector)
                for index in range(await candidates.count()):
                    checkbox = candidates.nth(index)
                    try:
                        if not v5_visible and not await checkbox.is_visible():
                            continue
                        if not await checkbox.is_checked():
                            box = await checkbox.bounding_box()
                            if box is not None:
                                await page.mouse.click(
                                    box["x"] + box["width"] / 2,
                                    box["y"] + box["height"] / 2,
                                )
                                await page.wait_for_timeout(100)
                        if not await checkbox.is_checked():
                            await checkbox.evaluate(
                                """element => {
                                    element.checked = true;
                                    element.dispatchEvent(
                                        new Event("input", { bubbles: true })
                                    );
                                    element.dispatchEvent(
                                        new Event("change", { bubbles: true })
                                    );
                                }"""
                            )
                        if await checkbox.is_checked():
                            return
                    except Exception:
                        continue

        await page.wait_for_timeout(200)

    if await first_visible_anywhere(page, ("#old_agreement",)) is not None:
        return

    raise LoginError("登录协议勾选框已出现，但未能完成勾选。")


async def perform_login(
    context: BrowserContext,
    page: Page,
    args: argparse.Namespace,
    username: str,
    password: str,
) -> Page:
    await page.goto(args.url, wait_until="commit", timeout=args.timeout_ms)
    await page.wait_for_timeout(2_000)
    page = await enter_login_page(context, page)
    if await is_authenticated_surface(page):
        return page

    username_input, password_input = await wait_for_login_form(page)
    if username_input is None or password_input is None:
        return page
    if args.auto_submit_login:
        await accept_consent_interstitials(page)
        username_input, password_input = await wait_for_login_form(page)
        if username_input is None or password_input is None:
            return page

    await username_input.fill(username)
    await username_input.dispatch_event("input")
    await password_input.fill(password)
    await password_input.dispatch_event("input")

    if not args.auto_submit_login:
        if args.headless:
            raise LoginError(
                "当前为人工登录模式，不能使用 --headless；"
                "请显示浏览器，或显式添加 --auto-submit-login。"
            )
        return await wait_for_manual_login(context, page)

    await ensure_login_agreement(page)
    await password_input.dispatch_event("input")
    await page.wait_for_timeout(300)
    await ensure_login_agreement(page)

    before_url = page.url
    login_button = await first_visible_anywhere(page, LOGIN_BUTTON_SELECTORS)
    if login_button is None:
        print("登录按钮尚未显现，正在启用官网登录按钮。", flush=True)
        real_login_button = page.locator("#login_btn")
        if await real_login_button.count() == 0:
            raise LoginError("没有找到官网真实登录按钮。")
        await real_login_button.evaluate(
            """element => {
                element.style.display = "block";
                const gray = document.querySelector("#login_btn_gray");
                if (gray) gray.style.display = "none";
            }"""
        )
        page = await click_and_follow_new_page(
            context, page, real_login_button
        )
    else:
        if not await login_button.is_enabled():
            raise LoginError("登录按钮仍未启用，请检查协议勾选状态。")
        page = await click_and_follow_new_page(context, page, login_button)
    await accept_consent_interstitials(page, wait_first_ms=3_000)
    await wait_for_login_result(page, before_url, args.timeout_ms)
    return page


async def enter_application(
    context: BrowserContext, page: Page, timeout_ms: int
) -> Page:
    while True:
        if await visible_text_anywhere(page, ("财务报表",), exact=True):
            return page

        try:
            _, enter_button = await wait_for_visible_text_anywhere(
                page, ("进入使用",), timeout_ms=min(timeout_ms, 2_000)
            )
            break
        except LoginError:
            print("正在等待“进入使用”按钮加载", flush=True)

    print("正在点击“进入使用”。", flush=True)
    page = await click_and_follow_new_page(context, page, enter_button, wait_ms=5_000)
    await page.wait_for_timeout(2_000)
    return page


async def navigate_to_cashflow(
    context: BrowserContext, page: Page, timeout_ms: int
) -> Page:
    _, finance_reports = await wait_for_visible_text_anywhere(
        page, ("财务报表",), timeout_ms=timeout_ms
    )
    print("正在进入“财务报表”。", flush=True)
    page = await click_and_follow_new_page(
        context, page, finance_reports, wait_ms=3_000
    )
    await page.wait_for_timeout(1_000)

    _, cashflow = await wait_for_visible_text_anywhere(
        page, ("现金流量表",), timeout_ms=timeout_ms
    )
    print("正在进入“现金流量表”。", flush=True)
    page = await click_and_follow_new_page(context, page, cashflow, wait_ms=5_000)
    await page.wait_for_timeout(2_000)
    return page


async def find_period_input(frame: Frame) -> Locator | None:
    return await find_period_input_from_root(frame.locator("body"))


async def find_period_input_from_root(root: Locator) -> Locator | None:
    inputs = root.locator("input")
    for index in range(min(await inputs.count(), 100)):
        candidate = inputs.nth(index)
        try:
            if not await candidate.is_visible():
                continue
            value = (await candidate.input_value()).strip()
            placeholder = (await candidate.get_attribute("placeholder") or "").strip()
            if "期间" in placeholder or PERIOD_RE.search(value):
                return candidate
        except Exception:
            continue
    return None


async def visible_period_options(page: Page) -> list[tuple[int, Locator]]:
    option_selectors = (
        '[role="option"]',
        ".el-select-dropdown__item",
        ".ant-select-item-option",
        ".kd-select-option",
        ".kd-dropdown-menu-item",
        "li",
    )
    found: list[tuple[int, Locator]] = []
    for frame in page.frames:
        for selector in option_selectors:
            options = frame.locator(selector)
            for index in range(min(await options.count(), 100)):
                option = options.nth(index)
                try:
                    if not await option.is_visible():
                        continue
                    text = (await option.inner_text()).strip()
                except Exception:
                    continue
                match = PERIOD_RE.search(text)
                if match:
                    found.append((int(match.group(1)), option))
        if found:
            break
    return found


async def select_max_period(page: Page, frame: Frame) -> int:
    period_input = await find_period_input(frame)
    if period_input is None:
        raise LoginError("现金流量表页面没有找到期间选择框。")

    await period_input.click()
    await page.wait_for_timeout(500)
    options = await visible_period_options(page)
    if options:
        max_period = max(period for period, _ in options)
        target = next(locator for period, locator in options if period == max_period)
        await target.click()
        print(f"已选择最大期间：第 {max_period} 期。", flush=True)
        return max_period

    await period_input.press("End")
    await period_input.press("Enter")
    await page.wait_for_timeout(500)
    value = await period_input.input_value()
    periods = [int(value) for value in PERIOD_RE.findall(value)]
    if not periods:
        raise LoginError("期间下拉列表中没有识别到第 1 至第 12 期。")
    max_period = max(periods)
    print(f"已选择最大期间：第 {max_period} 期。", flush=True)
    return max_period


async def click_frame_button(frame: Frame, texts: Sequence[str]) -> None:
    await click_root_button(frame.locator("body"), texts)


async def click_root_button(root: Locator, texts: Sequence[str]) -> None:
    for text in texts:
        role_button = root.get_by_role("button", name=text, exact=True)
        for index in range(min(await role_button.count(), 20)):
            candidate = role_button.nth(index)
            if await candidate.is_visible():
                await candidate.click()
                return

        selectors = (
            f'input[type="button"][value="{text}"]',
            f'input[type="submit"][value="{text}"]',
            f'button:has-text("{text}")',
        )
        candidate = await first_visible_in_frame_from_root(root, selectors)
        if candidate is not None:
            await candidate.click()
            return

        exact_text = root.get_by_text(text, exact=True)
        for index in range(min(await exact_text.count(), 20)):
            candidate = exact_text.nth(index)
            if await candidate.is_visible():
                await candidate.click()
                return
    raise LoginError(f"没有找到按钮：{' / '.join(texts)}")


async def find_cashflow_frame(page: Page, timeout_ms: int) -> Frame:
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1_000
    while asyncio.get_running_loop().time() < deadline:
        for frame in page.frames:
            monthly = frame.get_by_text("本月金额", exact=True)
            lines = frame.get_by_text("行次", exact=True)
            try:
                if (
                    await monthly.count()
                    and await monthly.first.is_visible()
                    and await lines.count()
                    and await lines.first.is_visible()
                ):
                    return frame
            except Exception:
                continue
        await page.wait_for_timeout(400)
    raise LoginError("没有找到现金流量表数据区域。")


async def discover_cashflow_table(frame: Frame) -> CashflowTable:
    tables = frame.locator("table")

    async def table_with_data(
        header_table: Locator, column_index: int
    ) -> Locator:
        if await header_table.locator("tbody tr").count():
            return header_table

        best_table = header_table
        best_rows = 0
        for body_index in range(await tables.count()):
            candidate = tables.nth(body_index)
            try:
                if not await candidate.is_visible():
                    continue
                rows = candidate.locator("tbody tr")
                row_count = await rows.count()
                if row_count <= best_rows:
                    continue
                max_cells = 0
                for row_index in range(min(row_count, 10)):
                    max_cells = max(
                        max_cells,
                        await rows.nth(row_index).locator("td").count(),
                    )
                if max_cells <= column_index:
                    continue
                best_table = candidate
                best_rows = row_count
            except Exception:
                continue
        return best_table

    for index in range(await tables.count()):
        table = tables.nth(index)
        try:
            if not await table.is_visible():
                continue
            header_rows = table.locator("thead tr")
            if not await header_rows.count():
                continue
            headers = header_rows.last.locator("th,td")
            header_texts = [
                (await headers.nth(i).inner_text()).strip()
                for i in range(await headers.count())
            ]
        except Exception:
            continue
        for column_index, text in enumerate(header_texts):
            if "本月金额" in text:
                data_table = await table_with_data(table, column_index)
                return CashflowTable(
                    frame=frame,
                    table=data_table,
                    amount_column_index=column_index,
                    row_selector="tbody tr",
                    cell_selector="td",
                )

    grid_selectors = (
        '[role="grid"]',
        ".el-table",
        ".ant-table",
        ".vxe-table",
        ".kd-table",
    )
    for selector in grid_selectors:
        grids = frame.locator(selector)
        for index in range(await grids.count()):
            grid = grids.nth(index)
            try:
                if not await grid.is_visible():
                    continue
                headers = grid.locator('[role="columnheader"], thead th')
                header_texts = [
                    (await headers.nth(i).inner_text()).strip()
                    for i in range(await headers.count())
                ]
            except Exception:
                continue
            for column_index, text in enumerate(header_texts):
                if "本月金额" in text:
                    return CashflowTable(
                        frame=frame,
                        table=grid,
                        amount_column_index=column_index,
                        row_selector='[role="row"]:has([role="gridcell"]), tbody tr',
                        cell_selector='[role="gridcell"],[role="cell"],td',
                    )
    raise LoginError("没有识别到包含“本月金额”的现金流量表格。")


async def eligible_cashflow_row_indexes(table: CashflowTable) -> list[int]:
    rows = table.table.locator(table.row_selector)
    eligible: list[int] = []
    for row_index in range(await rows.count()):
        cells = rows.nth(row_index).locator(table.cell_selector)
        if await cells.count() <= table.amount_column_index:
            continue
        try:
            text = (
                (await cells.nth(table.amount_column_index).inner_text())
                .strip()
                .replace(" ", "")
                .replace("￥", "")
                .replace("¥", "")
            )
        except Exception:
            continue
        if text and AMOUNT_RE.fullmatch(text):
            eligible.append(row_index)
    return eligible


async def print_cashflow_structure(frame: Frame) -> None:
    """仅输出表格结构计数，避免在终端暴露任何财务明细。"""

    tables = frame.locator("table")
    visible_index = 0
    for index in range(min(await tables.count(), 20)):
        table = tables.nth(index)
        try:
            if not await table.is_visible():
                continue
            visible_index += 1
            headers = await table.locator("thead th, thead td").count()
            body_rows = await table.locator("tbody tr").count()
            first_cells = 0
            if body_rows:
                first_cells = await table.locator("tbody tr").first.locator("td").count()
            class_name = (await table.get_attribute("class") or "-").strip()
            print(
                "现金流量表结构"
                f" {visible_index}：表头 {headers} 列，数据 {body_rows} 行，"
                f"首行 {first_cells} 列，类型 {class_name}",
                flush=True,
            )
        except Exception:
            continue


async def click_cashflow_amount(table: CashflowTable, row_index: int) -> None:
    rows = table.table.locator(table.row_selector)
    if row_index >= await rows.count():
        raise LoginError("现金流量表行数发生变化。")
    cells = rows.nth(row_index).locator(table.cell_selector)
    cell = cells.nth(table.amount_column_index)
    await cell.scroll_into_view_if_needed()
    nested = cell.locator('a,button,[role="link"]')
    for index in range(min(await nested.count(), 10)):
        candidate = nested.nth(index)
        if await candidate.is_visible():
            await candidate.click()
            return
    await cell.click()


async def wait_for_adjustment_surface(
    context: BrowserContext,
    before_pages: Sequence[Page],
    main_page: Page,
    timeout_ms: int,
) -> AdjustmentSurface | None:
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1_000
    while asyncio.get_running_loop().time() < deadline:
        for candidate_page in context.pages:
            found = await visible_text_anywhere(
                candidate_page, ("现金流量调整列表",), exact=True
            )
            if found is not None:
                frame, title = found
                container: Locator | None = None
                dialog_container = title.locator(
                    "xpath=ancestor::*[contains(@class,'dialog') or "
                    "contains(@class,'modal') or contains(@class,'window')][1]"
                )
                if await dialog_container.count():
                    container = dialog_container
                else:
                    action_container = title.locator(
                        "xpath=ancestor::*[.//button[normalize-space()='导出'] "
                        "or .//input[@value='导出']][1]"
                    )
                    if await action_container.count():
                        container = action_container
                return AdjustmentSurface(
                    page=candidate_page,
                    frame=frame,
                    container=container,
                    opened_new_page=candidate_page not in before_pages,
                )
        await main_page.wait_for_timeout(150)
    return None


async def set_adjustment_full_year(
    surface: AdjustmentSurface, max_period: int
) -> None:
    root = surface.container or surface.frame.locator("body")
    period_input = await find_period_input_from_root(root)
    if period_input is None:
        raise LoginError("调整列表没有找到期间选择框。")

    value = (await period_input.input_value()).strip()
    periods = [int(item) for item in PERIOD_RE.findall(value)]
    if len(periods) >= 2 and min(periods) == 1 and max(periods) >= max_period:
        return

    await period_input.click()
    await surface.page.wait_for_timeout(300)
    full_year = await visible_text_anywhere(surface.page, ("本年",), exact=True)
    if full_year is not None:
        _, option = full_year
        await option.click()
        await surface.page.wait_for_timeout(300)
        return

    if await period_input.is_editable():
        year_match = re.search(r"(20\d{2})年", value)
        year = int(year_match.group(1)) if year_match else datetime.now().year
        await period_input.fill(f"{year}年01期 - {year}年{max_period:02d}期")
        await period_input.press("Enter")
        return

    raise LoginError("调整列表期间无法切换为本年或第 1 期至最后一期。")


async def find_adjustment_table(surface: AdjustmentSurface) -> Locator:
    root = surface.container or surface.frame.locator("body")
    tables = root.locator("table")
    for index in range(await tables.count()):
        table = tables.nth(index)
        try:
            text = await table.inner_text()
            if await table.is_visible() and "日期" in text and "凭证字号" in text:
                return table
        except Exception:
            continue

    grids = root.locator('[role="grid"],.el-table,.ant-table,.vxe-table,.kd-table')
    for index in range(await grids.count()):
        grid = grids.nth(index)
        try:
            text = await grid.inner_text()
            if await grid.is_visible() and "日期" in text and "凭证字号" in text:
                return grid
        except Exception:
            continue
    raise LoginError("调整列表没有识别到凭证明细表格。")


async def select_all_adjustment_rows(surface: AdjustmentSurface) -> None:
    table = await find_adjustment_table(surface)
    selectors = (
        'thead input[type="checkbox"]',
        'thead [role="checkbox"]',
        "thead .el-checkbox",
        "thead .ant-checkbox-wrapper",
        "thead .ant-checkbox",
        "thead .kd-checkbox",
        "thead [class*='checkbox']",
        '.el-table__header input[type="checkbox"]',
        '.ant-table-thead input[type="checkbox"]',
        '.kd-table-header [role="checkbox"]',
    )
    checkbox = await first_visible_in_frame_from_root(table, selectors)
    if checkbox is None:
        raise LoginError("调整列表没有找到表头全选框。")
    try:
        if not await checkbox.is_checked():
            await checkbox.check(force=True)
    except Exception:
        await checkbox.click(force=True)


async def first_visible_in_frame_from_root(
    root: Locator, selectors: Iterable[str]
) -> Locator | None:
    for selector in selectors:
        locator = root.locator(selector)
        for index in range(min(await locator.count(), 20)):
            candidate = locator.nth(index)
            try:
                if await candidate.is_visible():
                    return candidate
            except Exception:
                continue
    return None


def safe_download_name(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return cleaned or "cashflow-adjustment.xlsx"


async def export_adjustment(
    surface: AdjustmentSurface, output_dir: Path, ordinal: int
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    root = surface.container or surface.frame.locator("body")
    export_button = root.get_by_role("button", name="导出", exact=True)
    button: Locator | None = None
    for index in range(min(await export_button.count(), 20)):
        candidate = export_button.nth(index)
        if await candidate.is_visible():
            button = candidate
            break
    if button is None:
        button = await first_visible_in_frame_from_root(
            root,
            (
                'input[type="button"][value="导出"]',
                'button:has-text("导出")',
                '[role="button"]:has-text("导出")',
            ),
        )
    if button is None:
        raise LoginError("调整列表没有找到导出按钮。")

    async with surface.page.expect_download(timeout=60_000) as download_info:
        await button.click()
    download = await download_info.value
    filename = safe_download_name(download.suggested_filename)
    destination = output_dir / f"{ordinal:03d}_{filename}"
    suffix = 1
    while destination.exists():
        destination = output_dir / f"{ordinal:03d}_{suffix}_{filename}"
        suffix += 1
    await download.save_as(str(destination))
    return destination


async def close_adjustment(surface: AdjustmentSurface) -> None:
    if surface.opened_new_page:
        await surface.page.close()
        return

    close_selectors = (
        '[aria-label="关闭"]',
        '[aria-label="Close"]',
        ".el-dialog__headerbtn",
        ".ant-modal-close",
        ".kd-modal-close",
        "[class*='dialog'] [class*='close']",
        "[class*='modal'] [class*='close']",
        'button:has-text("×")',
        'button:has-text("关闭")',
    )
    close_button: Locator | None = None
    titles = surface.frame.get_by_text("现金流量调整列表", exact=True)
    for index in range(min(await titles.count(), 10)):
        title = titles.nth(index)
        if not await title.is_visible():
            continue
        container = surface.container or title.locator(
            "xpath=ancestor::*[contains(@class,'dialog') or "
            "contains(@class,'modal') or contains(@class,'window')][1]"
        )
        if await container.count():
            close_button = await first_visible_in_frame_from_root(
                container, close_selectors
            )
            if close_button is not None:
                break
    if close_button is None:
        close_button = await first_visible_in_frame(surface.frame, close_selectors)
    if close_button is not None:
        await close_button.click()
    else:
        await surface.page.keyboard.press("Escape")
    await surface.page.wait_for_timeout(500)


async def run_cashflow_export(
    context: BrowserContext,
    page: Page,
    args: argparse.Namespace,
) -> tuple[int, int, int, Path]:
    page = await enter_application(context, page, args.timeout_ms)
    page = await navigate_to_cashflow(context, page, args.timeout_ms)
    frame = await find_cashflow_frame(page, args.timeout_ms)
    max_period = await select_max_period(page, frame)
    await click_frame_button(frame, ("查询",))
    await page.wait_for_timeout(1_500)

    frame = await find_cashflow_frame(page, args.timeout_ms)
    table = await discover_cashflow_table(frame)
    eligible = await eligible_cashflow_row_indexes(table)
    if not eligible:
        await print_cashflow_structure(frame)
    if args.max_rows is not None:
        eligible = eligible[: args.max_rows]

    run_dir = args.download_dir or (
        DEFAULT_DOWNLOAD_ROOT / datetime.now().strftime("%Y%m%d_%H%M%S")
    )
    run_dir.mkdir(parents=True, exist_ok=True)
    print(f"本月金额列共有 {len(eligible)} 行可处理。", flush=True)

    exported = 0
    skipped = 0
    failed = 0
    for ordinal, row_index in enumerate(eligible, start=1):
        print(f"正在处理第 {ordinal}/{len(eligible)} 行。", flush=True)
        frame = await find_cashflow_frame(page, args.timeout_ms)
        table = await discover_cashflow_table(frame)
        before_pages = list(context.pages)
        try:
            await click_cashflow_amount(table, row_index)
            surface = await wait_for_adjustment_surface(
                context,
                before_pages,
                page,
                args.row_open_timeout_ms,
            )
            if surface is None:
                skipped += 1
                print(f"第 {ordinal} 行 3 秒内未打开调整列表，已跳过。", flush=True)
                unexpected_pages = [
                    item for item in context.pages if item not in before_pages
                ]
                for unexpected in unexpected_pages:
                    try:
                        await unexpected.close()
                    except Exception:
                        pass
                continue

            try:
                await set_adjustment_full_year(surface, max_period)
                adjustment_root = surface.container or surface.frame.locator("body")
                await click_root_button(adjustment_root, ("查询",))
                await surface.page.wait_for_timeout(1_000)
                await select_all_adjustment_rows(surface)
                await export_adjustment(surface, run_dir, ordinal)
                exported += 1
                print(f"第 {ordinal} 行导出完成。", flush=True)
            finally:
                await close_adjustment(surface)
        except Exception as exc:
            failed += 1
            print(f"第 {ordinal} 行处理失败：{type(exc).__name__}", file=sys.stderr)
            unexpected_pages = [item for item in context.pages if item not in before_pages]
            for unexpected in unexpected_pages:
                try:
                    await unexpected.close()
                except Exception:
                    pass
    return exported, skipped, failed, run_dir


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="智云现金流量调整明细批量导出（只查询和导出）"
    )
    parser.add_argument("--url", default=DEFAULT_URL, help="登录入口 URL")
    parser.add_argument(
        "--username",
        default=os.getenv("JDY_USERNAME"),
        help="用户名；不提供时交互式输入，也可使用 JDY_USERNAME",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="无头运行；默认显示浏览器，便于观察完整流程",
    )
    parser.add_argument(
        "--auto-submit-login",
        action="store_true",
        help="自动勾选协议并点击登录；默认等待人工操作",
    )
    parser.add_argument(
        "--browser-channel",
        choices=("msedge", "chrome", "chromium"),
        default=DEFAULT_BROWSER_CHANNEL,
        help="浏览器通道，默认使用本机 Microsoft Edge",
    )
    parser.add_argument(
        "--user-data-dir",
        type=Path,
        default=DEFAULT_USER_DATA_DIR,
        help="独立浏览器配置目录；默认保留 Edge 缓存和登录会话",
    )
    parser.add_argument(
        "--cdp-url",
        help=(
            "连接已开启远程调试的 Edge，例如 http://127.0.0.1:9222；"
            "不提供时启动独立自动化 Edge"
        ),
    )
    parser.add_argument(
        "--storage-state",
        type=Path,
        help="显式保存登录后的 Playwright 会话文件；文件包含敏感会话凭据",
    )
    parser.add_argument(
        "--artifacts-dir",
        type=Path,
        default=DEFAULT_ARTIFACTS_DIR,
        help="失败截图和诊断文件目录",
    )
    parser.add_argument(
        "--timeout-ms",
        type=int,
        default=30_000,
        help="页面及登录等待超时，默认 30000 毫秒",
    )
    parser.add_argument(
        "--keep-open",
        action="store_true",
        help="流程结束后保持浏览器打开，按 Ctrl+C 退出",
    )
    parser.add_argument(
        "--login-only",
        action="store_true",
        help="只完成登录，不执行现金流量调整明细导出",
    )
    parser.add_argument(
        "--download-dir",
        type=Path,
        help="导出目录；默认写入工作区/04_产出/现金流量调整导出/时间戳",
    )
    parser.add_argument(
        "--row-open-timeout-ms",
        type=int,
        default=ROW_OPEN_TIMEOUT_MS,
        help="点击本月金额后等待调整列表的时间，默认 3000 毫秒",
    )
    parser.add_argument(
        "--max-rows",
        type=int,
        help="仅处理前 N 个可点击数据行，用于现场试跑",
    )
    return parser.parse_args()


async def run(args: argparse.Namespace) -> int:
    username = args.username or input("用户名：").strip()
    password = os.getenv("JDY_PASSWORD") or getpass.getpass("密码：")
    if not username or not password:
        raise LoginError("用户名和密码不能为空。")

    args.artifacts_dir.mkdir(parents=True, exist_ok=True)
    if not args.cdp_url:
        args.user_data_dir.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as playwright:
        owns_context = not args.cdp_url
        if args.cdp_url:
            try:
                attached_browser = await playwright.chromium.connect_over_cdp(
                    args.cdp_url
                )
            except Exception as exc:
                raise LoginError(
                    f"无法连接现有 Edge：{args.cdp_url}；"
                    "请确认浏览器已开启远程调试端口。"
                ) from exc
            if not attached_browser.contexts:
                raise LoginError("现有 Edge 没有可用的浏览器上下文。")
            context = attached_browser.contexts[0]
            page = await context.new_page()
            print("已连接现有 Edge，正在新标签页中执行 RPA。", flush=True)
        else:
            channel = (
                None if args.browser_channel == "chromium" else args.browser_channel
            )
            context = await playwright.chromium.launch_persistent_context(
                user_data_dir=str(args.user_data_dir.resolve()),
                channel=channel,
                headless=args.headless,
                accept_downloads=True,
            )
            page = context.pages[0] if context.pages else await context.new_page()
        page.set_default_timeout(args.timeout_ms)

        try:
            page = await perform_login(context, page, args, username, password)

            if args.storage_state:
                args.storage_state.parent.mkdir(parents=True, exist_ok=True)
                await context.storage_state(path=str(args.storage_state))
                print(f"登录成功，会话已保存到：{args.storage_state}")
            else:
                print(f"登录成功，当前页面：{page.url}")

            if not args.login_only:
                exported, skipped, failed, run_dir = await run_cashflow_export(
                    context, page, args
                )
                print(
                    f"现金流量调整明细处理完成：导出 {exported} 行，"
                    f"跳过 {skipped} 行，失败 {failed} 行。"
                )
                print(f"导出目录：{run_dir}")

            if args.keep_open:
                print("浏览器保持打开，按 Ctrl+C 退出。")
                while True:
                    await page.wait_for_timeout(1_000)
            return 0
        except (PlaywrightTimeoutError, LoginError) as exc:
            try:
                screenshot = await save_screenshot(page, args.artifacts_dir, "jdy-login-failure.png")
                print(f"流程失败：{exc}\n诊断截图：{screenshot}", file=sys.stderr)
            except Exception:
                print(f"流程失败：{exc}", file=sys.stderr)
            return 2
        finally:
            if owns_context:
                await close_browser_safely(context)


def main() -> int:
    args = parse_args()
    try:
        return asyncio.run(run(args))
    except LoginError as exc:
        print(f"流程失败：{exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("已退出。")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
