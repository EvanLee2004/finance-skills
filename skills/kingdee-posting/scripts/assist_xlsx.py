#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""客户核算项目余额表：Playwright 从总部星辰引出。OpenAPI 没有这张表。不捡本地旧 Excel。"""
from __future__ import annotations

import json
import sys
import time
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs

from openpyxl import Workbook, load_workbook

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import lookups as lookup_mod  # noqa: E402

ASSIST_FORM = "https://tf.jdy.com/ierp/index.html?formId=gl_rpt_assistbalance"
EXPORT_LOG = "https://tf.jdy.com/ierp/index.html?formId=bos_exportlog_list"
WORKBENCH = "https://service.jdy.com/workbench/web/index.html"
XINGCHEN_HOME = "https://tf.jdy.com/ierp/index.html?formId=home_page"
HQ_NAME = "甲骨易（北京）语言科技股份有限公司"
HQ_SEARCH = "语言科技"
STATE = Path.home() / ".config" / "finance" / "xingchen.playwright-state.json"
CACHE_DIR = Path.home() / ".cache" / "finance"
_ASSIST_CACHE: tuple[list, str] | None = None

# 与 技能/金蝶/工作区/20260907_客户核算项目余额表/现网探测.md 同一套过滤。
EXPORT_FILTERS = {
    "form_id": "gl_rpt_assistbalance",
    "assist_type": "客户",
    "account": "1131",
    "period": "本年",
    "hide_zero_balance": False,
    "show_detail_accounts": True,
}

MANUAL_EXPORT_STEPS = (
    "请核对本机 ~/.config/finance/xingchen.local.json（金蝶网页账密，与月度损益表同一份），重跑让脚本从总部星辰引出。"
    "不要把旧的核算项目余额表放进材料夹当源。"
)


def assist_xlsx_period_span(path: Path) -> str:
    """表头「期间：202601-202612」。本年必须 01-12。"""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=1, max_row=4, values_only=True):
            for v in row:
                t = str(v or "").strip()
                if t.startswith("期间"):
                    return t.split("：", 1)[-1].split(":", 1)[-1].strip()
    finally:
        wb.close()
    return ""


def assist_xlsx_company(path: Path) -> str:
    """表头「公司名称：xxx」。没有就返回空串。"""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=1, max_row=4, values_only=True):
            for v in row:
                t = str(v or "").strip()
                if t.startswith("公司名称"):
                    return t.split("：", 1)[-1].split(":", 1)[-1].strip()
    finally:
        wb.close()
    return ""


def parse_assist_xlsx_checked(path: Path) -> list[lookup_mod.AssistRow]:
    """解析并把「表不对」说清楚：不是总部账套 / 没有 1131 明细，都停下来问，不要拿空表把每笔都 hold。"""
    rows = parse_assist_xlsx(path)
    company = assist_xlsx_company(path)
    if not company or (HQ_NAME not in company and company not in HQ_NAME):
        raise SystemExit(
            f"网页引出的客户核算项目余额表公司名是「{company or '空'}」，不是总部账套。销项/收款只记总部。{MANUAL_EXPORT_STEPS} 未生成引入表。"
        )
    if not rows:
        raise SystemExit(
            f"网页引出的客户核算项目余额表没有 1131 明细行。多半没切到总部，或过滤不是客户+1131+本年。{MANUAL_EXPORT_STEPS} 未生成引入表。"
        )
    span = assist_xlsx_period_span(path)
    if not span or "-" not in span:
        raise SystemExit(
            f"客户核算项目余额表没有「期间：本年01-12」。{MANUAL_EXPORT_STEPS} 未生成引入表。"
        )
    start, end = (x.strip() for x in span.split("-", 1))
    if not (start.endswith("01") and end.endswith("12")):
        raise SystemExit(
            f"网页引出的期间是 {span}，不是本年（01期-12期）。过滤没点上「本年」。{MANUAL_EXPORT_STEPS} 未生成引入表。"
        )
    months = {lookup_mod._period_key(r.period)[-2:] for r in rows if lookup_mod._period_key(r.period)}
    if "01" not in months or "12" not in months:
        raise SystemExit(
            f"表头写本年，但行里没有 1 月和 12 月（现有 {sorted(months)}）。未生成引入表。"
        )
    return rows


_SCRAPE_JS = """() => {
  function cellsOf(el) {
    return [...el.querySelectorAll(':scope > td, :scope > th, :scope > [class*="cell"]')]
      .map(n => (n.innerText || '').replace(/\\s+/g, ' ').trim());
  }
  const rowEls = [...document.querySelectorAll('tr, [class*="kd-table-row"], [class*="TableRow"]')];
  let map = null;
  const rows = [];
  for (const el of rowEls) {
    const cells = cellsOf(el);
    if (cells.length < 4) continue;
    if (cells.includes('客户编码') && cells.includes('科目编码') && !cells.includes('借方')) {
      map = {};
      cells.forEach((t, i) => { if (t && map[t] === undefined) map[t] = i; });
      continue;
    }
    if (!map || map['科目编码'] === undefined || map['客户编码'] === undefined) continue;
    const acc = (cells[map['科目编码']] || '').replace(/\\.0$/, '');
    const code = (cells[map['客户编码']] || '').replace(/\\.0$/, '');
    if (!code || !String(acc).startsWith('1131')) continue;
    const ytd = map['本年累计'];
    const end = map['期末'];
    rows.push({
      period: cells[map['期间']] || '',
      code,
      name: cells[map['客户名称']] || '',
      account: acc,
      ytd_debit: ytd != null ? cells[ytd] : '',
      ytd_credit: ytd != null ? cells[ytd + 1] : '',
      end_debit: end != null ? cells[end] : '',
      end_credit: end != null ? cells[end + 1] : '',
    });
  }
  return rows;
}"""


async def _scrape_assist_rows(page) -> list[dict]:
    found: list[dict] = []
    for chunk in await _eval_frames(page, _SCRAPE_JS):
        if isinstance(chunk, list):
            found.extend(chunk)
    return found


def _virtual_payload(data):
    if isinstance(data, list) and data:
        block = data[0] if isinstance(data[0], dict) else {}
        params = block.get("p") or []
        head = params[0] if params and isinstance(params[0], dict) else {}
        args = head.get("args") or []
        if args and isinstance(args[0], dict) and "rows" in args[0]:
            return args[0]
        if "rows" in head:
            return head
    if isinstance(data, dict) and "rows" in data:
        return data
    return None


def _rows_from_virtual(data) -> tuple[list[dict], int, int]:
    payload = _virtual_payload(data)
    if not payload:
        return [], 0, 0
    idx = payload.get("dataindex") or {}
    rows = payload.get("rows") or []
    total = int(payload.get("datacount") or 0)
    pi, ci, ni, ai = idx.get("period"), idx.get("f0001number"), idx.get("f0001name"), idx.get("acctnumber")
    ytd_d, ytd_c = idx.get("ytddebit"), idx.get("ytdcredit")
    end_d, end_c = idx.get("enddebit"), idx.get("endcredit")
    out = []
    raw_n = len(rows)
    for row in rows:
        if not isinstance(row, list):
            continue
        acc = str(row[ai] if isinstance(ai, int) and ai < len(row) else "").strip()
        if acc.endswith(".0") and acc[:-2].replace(".", "", 1).isdigit():
            acc = acc[:-2]
        if not acc.startswith("1131"):
            continue
        code = str(row[ci] if isinstance(ci, int) and ci < len(row) else "").strip()
        if code.endswith(".0") and code[:-2].isdigit():
            code = code[:-2]
        if not code:
            continue
        name = str(row[ni] if isinstance(ni, int) and ni < len(row) else "").strip()
        period = str(row[pi] if isinstance(pi, int) and pi < len(row) else "").strip()
        def _num(i):
            if not isinstance(i, int) or i >= len(row):
                return None
            return lookup_mod.money(row[i])
        out.append(
            {
                "period": period,
                "code": code,
                "name": name,
                "account": acc,
                "ytd_debit": _num(ytd_d),
                "ytd_credit": _num(ytd_c),
                "end_debit": _num(end_d),
                "end_credit": _num(end_c),
            }
        )
    return out, total, raw_n


async def _fetch_virtual_pages(page, url: str, post_data: str) -> tuple[list[dict], bool]:
    qs = parse_qs(post_data or "")
    raw_params = (qs.get("params") or [""])[0]
    if not raw_params:
        return [], False
    params = json.loads(raw_params)
    page_id = (qs.get("pageId") or [""])[0]
    app_id = (qs.get("appId") or ["gl"])[0]
    all_rows: list[dict] = []
    start = 0
    size = 1000
    total = None
    fetched = 0
    while True:
        if isinstance(params, list) and params and isinstance(params[0], dict):
            params[0]["args"] = [start, size]
        resp = await page.request.post(
            url,
            form={
                "pageId": page_id,
                "appId": app_id,
                "params": json.dumps(params, ensure_ascii=False, separators=(",", ":")),
            },
        )
        if not resp.ok:
            return all_rows, False
        data = await resp.json()
        chunk, count, raw_n = _rows_from_virtual(data)
        if total is None:
            total = count
        all_rows.extend(chunk)
        if raw_n <= 0:
            break
        start += raw_n
        fetched += raw_n
        if total is not None and start >= total:
            break
        if start > 200000:
            break
    complete = total is not None and total > 0 and fetched >= total
    return all_rows, complete


def _scrape_is_year(rows: list[dict]) -> bool:
    months = set()
    for row in rows:
        key = lookup_mod._period_key(row.get("period") or "")
        if len(key) >= 6 and key[-2:].isdigit():
            months.add(key[-2:])
    return "01" in months and "12" in months


def write_assist_book(path: Path, company: str, span: str, rows: list[dict]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "核算项目余额表"
    ws.append(["核算项目余额表"])
    ws.append([f"公司名称：{company}", "", "", "", "", "", "", f"期间：{span}"])
    ws.append(["期间", "客户编码", "客户名称", "科目编码", "科目名称", "期初", "期初", "本期", "本期", "本年累计", "本年累计", "期末", "期末"])
    ws.append(["期间", "客户编码", "客户名称", "科目编码", "科目名称", "借方", "贷方", "借方", "贷方", "借方", "贷方", "借方", "贷方"])
    for row in rows:
        ws.append(
            [
                row.get("period"),
                row.get("code"),
                row.get("name"),
                row.get("account"),
                "",
                None,
                None,
                None,
                None,
                row.get("ytd_debit"),
                row.get("ytd_credit"),
                row.get("end_debit"),
                row.get("end_credit"),
            ]
        )
    wb.save(path)
    wb.close()
    return path


def parse_assist_xlsx(path: Path) -> list[lookup_mod.AssistRow]:
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        col: dict[str, int] = {}
        ytd_i = None
        end_i = None
        out: list[lookup_mod.AssistRow] = []
        for row in ws.iter_rows(values_only=True):
            vals = list(row)
            texts = [str(c).strip() if c is not None else "" for c in vals]
            if "期间" in texts and "客户编码" in texts and "科目编码" in texts:
                if texts and texts[0] == "期间" and "借方" in texts:
                    continue
                for i, t in enumerate(texts):
                    if t == "期间" and "period" not in col:
                        col["period"] = i
                    elif t in ("客户编码", "客户代码", "核算项目编码"):
                        col["code"] = i
                    elif t in ("客户名称", "核算项目名称"):
                        col["name"] = i
                    elif t in ("科目编码", "科目代码"):
                        col["account"] = i
                    elif t == "本年累计" and ytd_i is None:
                        ytd_i = i
                    elif t == "期末" and end_i is None:
                        end_i = i
                continue
            if not col:
                continue
            period_i = col.get("period")
            code_i = col.get("code")
            name_i = col.get("name")
            acc_i = col.get("account")
            if period_i is None or code_i is None or acc_i is None:
                continue
            if acc_i >= len(vals):
                continue
            code = str(vals[code_i] or "").strip()
            acc = str(vals[acc_i] or "").strip()
            if acc.endswith(".0") and acc[:-2].isdigit():
                acc = acc[:-2]
            if code.endswith(".0") and code[:-2].isdigit():
                code = code[:-2]
            if not code or not acc.startswith("1131"):
                continue
            name = str(vals[name_i] or "").strip() if name_i is not None and name_i < len(vals) else ""
            ytd_debit = ytd_credit = end_debit = end_credit = None
            if ytd_i is not None and ytd_i + 1 < len(vals):
                ytd_debit = lookup_mod.money(vals[ytd_i])
                ytd_credit = lookup_mod.money(vals[ytd_i + 1])
            if end_i is not None and end_i + 1 < len(vals):
                end_debit = lookup_mod.money(vals[end_i])
                end_credit = lookup_mod.money(vals[end_i + 1])
            out.append(
                lookup_mod.AssistRow(
                    period=lookup_mod._period_key(vals[period_i] if period_i < len(vals) else ""),
                    customer_code=code,
                    customer_name=name,
                    account=acc,
                    ending_debit=end_debit,
                    ending_credit=end_credit,
                    ytd_debit=ytd_debit,
                    ytd_credit=ytd_credit,
                )
            )
        return out
    finally:
        wb.close()


async def _click_text(page, label: str) -> bool:
    return bool(
        await page.evaluate(
            """(label) => {
              const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
              let n;
              while ((n = walker.nextNode())) {
                if ((n.textContent || '').trim() === label) {
                  const el = n.parentElement;
                  (el?.closest('label,li,button,div,span') || el)?.click();
                  return true;
                }
              }
              return false;
            }""",
            label,
        )
    )


async def _dismiss_overlays(page) -> None:
    await page.evaluate(
        """() => {
          document.querySelectorAll('iframe').forEach(f => {
            const src = f.src || '';
            if (/piaozone|fpdk|etax/.test(src)) (f.closest('div') || f).remove();
          });
        }"""
    )


async def _open_context(p):
    browser = await p.chromium.launch(headless=False, channel="chrome")
    kwargs = {"accept_downloads": True, "locale": "zh-CN"}
    if STATE.is_file():
        kwargs["storage_state"] = str(STATE)
    ctx = await browser.new_context(**kwargs)
    return ctx, browser, "fresh"


LOGIN_SCRIPT = HERE.parent.parent / "pl-dept-report" / "scripts" / "xingchen_login.py"
EXPORT_SCRIPT = HERE.parent.parent / "pl-dept-report" / "scripts" / "xingchen_export.py"
ASK_LOGIN = (
    "本机没有金蝶网页账密。请把用户名密码写进 ~/.config/finance/xingchen.local.json"
    "（与月度损益表同一份，格式看 pl-dept-report/config/xingchen.local.example.json）。"
    "销项/收款抄 1131 靠脚本从总部星辰引出客户核算项目余额表，不要把旧表放进材料夹。"
)


def _load_sibling(name: str, path: Path):
    if not path.is_file():
        return None
    import importlib.util

    folder = str(path.parent)
    added = folder not in sys.path
    if added:
        sys.path.insert(0, folder)
    try:
        spec = importlib.util.spec_from_file_location(name, path)
        if spec is None or spec.loader is None:
            return None
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    finally:
        if added and sys.path and sys.path[0] == folder:
            sys.path.pop(0)


def _login_mod():
    return _load_sibling("kingdee_posting_xingchen_login", LOGIN_SCRIPT)


def _export_mod():
    """损益表脚本和本技能都有 inspect_inputs.py，加载时先挪开本技能那份。"""
    folder = str(EXPORT_SCRIPT.parent)
    saved = []
    for key in ("inspect_inputs", "common", "layout"):
        mod = sys.modules.get(key)
        origin = str(getattr(mod, "__file__", "") or "") if mod else ""
        if mod is not None and "kingdee-posting" in origin.replace("\\", "/"):
            saved.append((key, sys.modules.pop(key)))
    added = folder not in sys.path
    if added:
        sys.path.insert(0, folder)
    try:
        return _load_sibling("kingdee_posting_xingchen_export", EXPORT_SCRIPT)
    finally:
        if added and sys.path and sys.path[0] == folder:
            sys.path.pop(0)
        for key, mod in saved:
            sys.modules[key] = mod


def _logged_out(page) -> bool:
    url = page.url or ""
    return "logout" in url or "login" in url


async def _wait_assist_form(page) -> None:
    await page.goto(ASSIST_FORM, wait_until="domcontentloaded")
    js = """() => [...document.querySelectorAll('input')]
        .filter(el => el.offsetParent && /年/.test(el.value || '') && /期/.test(el.value || ''))
        .map(el => el.value)"""
    for _ in range(30):
        await _dismiss_overlays(page)
        if _logged_out(page):
            return
        vals = []
        for chunk in await _eval_frames(page, js):
            vals.extend(chunk or [])
        if vals and await page.get_by_text("引出", exact=True).count():
            return
        await page.wait_for_timeout(700)


async def _ensure_xingchen(page) -> bool:
    login = _login_mod()
    if login is None:
        raise SystemExit(ASK_LOGIN)
    creds = login.load_creds()
    last_err = None
    for attempt in range(3):
        try:
            await login.ensure_login(page, creds)
            last_err = None
            break
        except Exception as e:
            last_err = e
            msg = str(e)
            if "ERR_" not in msg and "Timeout" not in msg and "net::" not in msg:
                raise
            await page.wait_for_timeout(2500 * (attempt + 1))
    if last_err is not None:
        raise SystemExit("金蝶网页现在连不上（tf.jdy.com / 工作台）。请检查网络后重跑。")
    try:
        await login.save_state(page.context)
    except Exception:
        pass
    if await page.get_by_text("进入使用").count():
        await page.get_by_text("进入使用").first.click()
        await page.wait_for_timeout(2000)
    exporter = _export_mod()
    if exporter is None or not hasattr(exporter, "_switch_book"):
        raise SystemExit("找不到月度损益表技能的切账套脚本（pl-dept-report/scripts/xingchen_export.py）。两个技能要装在一起。")
    return bool(await exporter._switch_book(page, HQ_NAME, HQ_SEARCH))


async def _eval_frames(page, js, arg=None):
    out = []
    for frame in page.frames:
        try:
            out.append(await frame.evaluate(js) if arg is None else await frame.evaluate(js, arg))
        except Exception:
            continue
    return out


async def _set_this_year_period(page) -> None:
    """期间必须本年（2026年01期-12期）。默认往往是当期一个月。"""
    year = date.today().year
    typed = f"{year}年01期 - {year}年12期"
    boxes = page.locator("input:visible")
    n = await boxes.count()
    for i in range(min(n, 12)):
        try:
            value = await boxes.nth(i).input_value()
        except Exception:
            continue
        if "年" not in value or "期" not in value:
            continue
        box = boxes.nth(i)
        try:
            await box.click(force=True)
            await page.wait_for_timeout(300)
            await box.fill(typed)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(400)
        except Exception:
            pass
        if await _period_is_year(page):
            return
        loc = page.get_by_text("本年", exact=True)
        if await loc.count():
            try:
                await loc.last.click(force=True)
                await page.wait_for_timeout(400)
            except Exception:
                pass
        if await _period_is_year(page):
            return
        one = page.get_by_text("1期", exact=True)
        twelve = page.get_by_text("12期", exact=True)
        if await one.count() and await twelve.count():
            try:
                await one.first.click()
                await page.wait_for_timeout(200)
                await twelve.last.click()
                await page.wait_for_timeout(400)
            except Exception:
                pass
        break
    try:
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(200)
    except Exception:
        pass


async def _period_is_year(page) -> bool:
    js = """() => [...document.querySelectorAll('input')]
        .filter(el => el.offsetParent)
        .map(el => el.value || '')
        .filter(v => v.includes('年') && v.includes('期'))"""
    for vals in await _eval_frames(page, js):
        if any("01期" in v and "12期" in v for v in (vals or [])):
            return True
    return False


async def _set_customer_ar_filters(page) -> None:
    await _click_text(page, "展开过滤")
    await page.wait_for_timeout(400)
    box = page.locator(".kd-table-cell-basedata-container").first
    if await box.count():
        await box.click()
        await page.keyboard.type(EXPORT_FILTERS["assist_type"])
        await page.wait_for_timeout(800)
        if await page.get_by_text("客户", exact=True).count():
            await page.get_by_text("客户", exact=True).last.click()
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(400)
    if not await _click_text(page, "1131 应收账款"):
        await _click_text(page, "应收账款")
    if await page.get_by_text("显示最明细科目").count():
        await _click_text(page, "显示最明细科目")
    checked = page.locator("label").filter(has_text="余额为 0 不显示")
    if await checked.count():
        # 不要勾。若已勾则点掉。
        await page.evaluate(
            """() => {
              const labels = Array.from(document.querySelectorAll('label,span,div'));
              for (const el of labels) {
                if ((el.textContent || '').trim() === '余额为 0 不显示' ||
                    (el.textContent || '').trim() === '余额为0不显示') {
                  const input = el.querySelector('input[type=checkbox]') ||
                    el.parentElement?.querySelector('input[type=checkbox]');
                  if (input && input.checked) el.click();
                  break;
                }
              }
            }"""
        )
    await _set_this_year_period(page)
    await _click_text(page, "查询")
    await page.wait_for_timeout(3500)
    await _set_this_year_period(page)
    if not await _period_is_year(page):
        await _click_text(page, "查询")
        await page.wait_for_timeout(3500)
        await _set_this_year_period(page)
    for _ in range(12):
        body = await page.locator("body").inner_text()
        if "1131" in body and await page.get_by_text("引出", exact=True).count():
            break
        await page.wait_for_timeout(500)


def _assist_export_stamp(name: str) -> str:
    import re

    hit = re.search(r"核算项目余额表-(\d{14})", name or "")
    return hit.group(1) if hit else ""


_NAME_JS = """() => [...document.querySelectorAll('span,a,div,td')]
    .map(el => (el.innerText || '').trim())
    .filter(t => /^核算项目余额表-\\d{14}\\.xlsx$/.test(t))"""


async def _export_log_names(page) -> list[str]:
    found: list[str] = []
    for frame in page.frames:
        try:
            found.extend(await frame.evaluate(_NAME_JS) or [])
        except Exception:
            continue
    return found


async def _click_export(page) -> None:
    try:
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(200)
    except Exception:
        pass
    js = """() => {
      const nodes = [...document.querySelectorAll('button,a,span,div,li')];
      const hits = nodes.filter(el => (el.innerText || '').trim() === '引出' && el.offsetParent);
      const btn = hits.find(el => el.tagName === 'BUTTON' || (el.className || '').toString().toLowerCase().includes('btn')) || hits[0];
      if (!btn) return 0;
      btn.click();
      return hits.length;
    }"""
    for frame in page.frames:
        try:
            if int(await frame.evaluate(js) or 0):
                break
        except Exception:
            continue
    else:
        hit = page.get_by_text("引出", exact=True)
        if await hit.count():
            await hit.first.click()
        else:
            await _click_text(page, "引出")
    for _ in range(12):
        await page.wait_for_timeout(400)
        bodies = []
        for frame in page.frames:
            try:
                bodies.append(await frame.locator("body").inner_text())
            except Exception:
                continue
        blob = "\n".join(bodies)
        if "引出结果查询" in blob or "数据文件正在生成" in blob or "正在生成" in blob:
            break
    if await page.get_by_text("引出结果查询").count():
        await page.get_by_text("引出结果查询").first.click()
        await page.wait_for_timeout(600)


async def _export_via_log(page, dest: Path, before: set[str] | None = None) -> bool:
    dest.parent.mkdir(parents=True, exist_ok=True)
    mark = time.time()
    before = before or set()
    cutoff = time.strftime("%Y%m%d%H%M%S", time.localtime(mark - 5))
    await _click_export(page)
    await page.goto(EXPORT_LOG, wait_until="domcontentloaded")
    newest = None
    for _ in range(50):
        await page.wait_for_timeout(1000)
        names = await _export_log_names(page)
        fresh = [nm for nm in names if _assist_export_stamp(nm) >= cutoff]
        if before:
            fresh = [nm for nm in fresh if nm not in before]
        if fresh:
            newest = max(fresh, key=_assist_export_stamp)
            break
    if not newest:
        return False
    tmp = dest.with_name(dest.name + ".part")
    try:
        if tmp.exists():
            tmp.unlink()
        async with page.expect_download(timeout=60000) as dl_info:
            await page.get_by_text(newest, exact=True).first.click()
        download = await dl_info.value
        await download.save_as(str(tmp))
        if tmp.is_file() and tmp.stat().st_size > 400:
            tmp.replace(dest)
            return True
    except Exception:
        if tmp.exists():
            tmp.unlink()
    return False


def export_customer_assist_xlsx(dest: Path | None = None) -> dict:
    """网页引出客户+1131 核算项目余额表。失败不得假装 OpenAPI 通了。不捡本地旧表。"""
    dest = dest or (CACHE_DIR / "核算项目余额表_客户_1131_本次.xlsx")
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        import playwright  # noqa: F401
    except Exception as e:
        return {"ok": False, "error": f"playwright_missing:{type(e).__name__}", "path": None}

    async def _run():
        from playwright.async_api import async_playwright

        async with async_playwright() as p:
            ctx, browser, mode = await _open_context(p)
            page = await ctx.new_page()
            cap = {"url": None, "post": None}

            def _on_req(req):
                if "getVirtualData" in req.url and req.method == "POST" and req.post_data:
                    cap["url"] = req.url
                    cap["post"] = req.post_data

            page.on("request", _on_req)
            try:
                if not await _ensure_xingchen(page):
                    return {
                        "ok": False,
                        "error": (
                            "星辰已登录，但没切到总部账套「甲骨易（北京）语言科技股份有限公司」。"
                            "不要点空账套。请核对网页账密后重跑。"
                        ),
                        "path": None,
                        "mode": mode,
                    }
                await _dismiss_overlays(page)
                await _wait_assist_form(page)
                await _set_customer_ar_filters(page)
                if _logged_out(page) or await page.get_by_text("引出", exact=True).count() == 0:
                    if not await _ensure_xingchen(page):
                        return {
                            "ok": False,
                            "error": "设过滤时金蝶把会话挤掉了，切总部失败。请重跑。",
                            "path": None,
                            "mode": mode,
                        }
                    await _wait_assist_form(page)
                    await _set_customer_ar_filters(page)
                if not await _period_is_year(page):
                    js = """() => [...document.querySelectorAll('input')].filter(el=>el.offsetParent).map(el=>el.value||'').filter(v=>v.includes('期'))"""
                    shown = []
                    for chunk in await _eval_frames(page, js):
                        shown.extend(chunk or [])
                    return {
                        "ok": False,
                        "error": f"期间没设成本年（01期-12期），当前输入={shown}。未取表。",
                        "path": None,
                        "mode": mode,
                    }
                year = date.today().year
                for _ in range(25):
                    if cap.get("post") and cap.get("url"):
                        break
                    await page.wait_for_timeout(200)
                virtual_rows: list[dict] = []
                complete = False
                cap_url, cap_post = cap.get("url"), cap.get("post")
                if cap_url and cap_post:
                    try:
                        virtual_rows, complete = await _fetch_virtual_pages(page, cap_url, cap_post)
                    except Exception:
                        virtual_rows, complete = [], False
                if complete and _scrape_is_year(virtual_rows):
                    keys = [lookup_mod._period_key(r.get("period") or "") for r in virtual_rows]
                    keys = [k for k in keys if len(k) >= 6]
                    span = f"{min(keys)}-{max(keys)}" if keys else f"{year}01-{year}12"
                    write_assist_book(dest, HQ_NAME, span, virtual_rows)
                    parse_assist_xlsx_checked(dest)
                    return {"ok": True, "path": str(dest), "mode": mode + "+api"}
                if _logged_out(page) or await page.get_by_text("引出", exact=True).count() == 0:
                    return {
                        "ok": False,
                        "error": "页面表格不是本年 1131，也没有「引出」。请重跑。",
                        "path": None,
                        "mode": mode,
                    }
                ok = await _export_via_log(page, dest)
                if ok:
                    try:
                        parse_assist_xlsx_checked(dest)
                    except SystemExit as e:
                        return {"ok": False, "error": str(e), "path": None, "mode": mode}
                    return {"ok": True, "path": str(dest), "mode": mode}
                return {"ok": False, "error": "未能从页面表格或引出拿到本年核算项目余额表", "path": None, "mode": mode}
            finally:
                await ctx.close()
                if browser:
                    await browser.close()

    import asyncio

    try:
        return asyncio.run(_run())
    except SystemExit as e:
        code = str(e)
        human = {
            "missing_creds": ASK_LOGIN,
            "bad_password": "金蝶网页账号或密码错（xingchen.local.json），请核对后重跑",
            "need_captcha": "金蝶网页登录要滑块/短信验证，脚本过不去。请先关掉验证码后再跑，不要用手导的旧表凑。",
            "login_failed": "金蝶网页登录没成功。请核对 xingchen.local.json 后重跑。",
        }.get(code, code)
        return {"ok": False, "error": human, "path": None}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {str(e)[:200]}", "path": None}


def ensure_assist_xlsx(extra_dirs: list[Path] | None = None) -> Path:
    del extra_dirs  # 生产不认材料夹/Downloads/技能家里的旧表
    dest = CACHE_DIR / "核算项目余额表_客户_1131_本次.xlsx"
    got = export_customer_assist_xlsx(dest)
    if got.get("ok") and got.get("path"):
        path = Path(got["path"])
        if path.is_file() and path.stat().st_size > 400:
            parse_assist_xlsx_checked(path)
            return path
    reason = got.get("error") or "没有客户核算项目余额表"
    raise SystemExit(
        f"{reason}。这张表 OpenAPI 没有，销项/收款的应收科目都靠网页从总部引出。"
        "请写好 ~/.config/finance/xingchen.local.json 后重跑。不要把旧核算项目余额表放进材料夹。未生成引入表。"
    )


def clear_assist_cache() -> None:
    global _ASSIST_CACHE
    _ASSIST_CACHE = None


def fetch_hq_assist() -> tuple[list, str]:
    """本进程只引出一次。返回 (行, 文件名)。"""
    global _ASSIST_CACHE
    if _ASSIST_CACHE is not None:
        return _ASSIST_CACHE
    path = ensure_assist_xlsx()
    rows = parse_assist_xlsx_checked(path)
    _ASSIST_CACHE = (rows, path.name)
    return _ASSIST_CACHE
