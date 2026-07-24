"""智云登录 + 下单表 GetFilterRows（精简自利润看板，独立可跑，禁止 sys.path 指看板）。

日期筛：filterType=11 + dateRange=18 + minValue/maxValue（智云心法）。
客户端再滤一次防边界。
"""

from __future__ import annotations

import json
from datetime import date, datetime
from typing import Any, Callable

# 内置默认（与看板一致，非密钥）
ZHIYUN_DEFAULTS: dict[str, Any] = {
    "base_url": "http://192.168.10.167:18880",
    "app_id": "6ff4fb2e-e68c-4ee9-83a0-836de8f72c11",
    "orders_worksheet_id": "6501688ebf25d7b91abdb465",
}

PAGE_SIZE = 1000
MAX_PAGES = 500
DATE_COL_NAME = "下单日期"

# login selectors (board login_zhiyun)
ACCOUNT_SEL = "#txtMobilePhone"
PASSWORD_SEL = "input[type=password]"
LOGIN_BTN_SELECTORS = ["text=登 录", "text=登录", ".loginBtn"]
LOGIN_TIMEOUT_MS = 30000
POST_LOGIN_WAIT_MS = 6000


class LoginError(RuntimeError):
    pass


class FetchError(RuntimeError):
    pass


def parse_cell(cell, ctrl: dict) -> str:
    if cell in (None, ""):
        return ""
    if isinstance(cell, (list, dict)):
        v, s = cell, json.dumps(cell, ensure_ascii=False)
    else:
        s = str(cell)
        if s[:1] not in ("[", "{"):
            return s
        try:
            v = json.loads(s)
        except (ValueError, TypeError):
            return s
    if not isinstance(v, list):
        return s
    if v and isinstance(v[0], str):
        m = {o["key"]: o["value"] for o in (ctrl.get("options") or [])}
        return "/".join(m.get(k, k) for k in v)
    out = []
    for x in v:
        if isinstance(x, dict):
            out.append(
                x.get("fullname")
                or x.get("departmentName")
                or x.get("name")
                or x.get("organizeName")
                or x.get("sourcevalue")
                or ""
            )
        else:
            out.append(str(x))
    return "/".join(o for o in out if o)


def rows_to_records(rows: list[dict], controls: list[dict]) -> list[dict[str, str]]:
    cols = [(c["controlName"], c) for c in controls if c.get("controlName")]
    out = []
    for row in rows:
        rec: dict[str, str] = {}
        for name, c in cols:
            val = parse_cell(row.get(c["controlId"]), c)
            if name not in rec or (not rec[name] and val):
                rec[name] = val
        out.append(rec)
    return out


def build_date_range_filter(
    controls: list[dict],
    date_col_name: str,
    start: date | str,
    end: date | str,
) -> list[dict]:
    """Construct filterType=11 dateRange=18 min/max filter body (pure, offline-testable)."""
    start_s = _as_ymd(start)
    end_s = _as_ymd(end)
    matches = [c for c in (controls or []) if c.get("controlName") == date_col_name]
    if not matches:
        return []
    ctrl = matches[0]
    return [
        {
            "controlId": ctrl["controlId"],
            "dataType": ctrl.get("type", 15),
            "spliceType": 1,
            "filterType": 11,
            "dateRange": 18,
            "minValue": start_s,
            "maxValue": end_s,
            "value": "",
            "values": [],
            "spec": {},
        }
    ]


def _as_ymd(d: date | str) -> str:
    if isinstance(d, date):
        return d.isoformat()
    s = str(d).strip()[:10].replace("/", "-").replace(".", "-")
    datetime.strptime(s, "%Y-%m-%d")
    return s


def client_filter_by_date(
    records: list[dict[str, str]],
    start: date | str,
    end: date | str,
    date_col: str = DATE_COL_NAME,
) -> list[dict[str, str]]:
    """Inclusive client-side date filter (boundary safety net)."""
    start_s = _as_ymd(start)
    end_s = _as_ymd(end)
    out = []
    for r in records:
        raw = r.get(date_col) or ""
        head = str(raw).strip()[:10].replace("/", "-").replace(".", "-")
        try:
            datetime.strptime(head, "%Y-%m-%d")
        except ValueError:
            continue
        if start_s <= head <= end_s:
            out.append(r)
    return out


def fetch_all_rows(
    post: Callable[[str, dict], dict],
    worksheet_id: str,
    app_id: str,
    filter_controls: list[dict] | None = None,
) -> list[dict]:
    """翻页拉全部命中行。

    以服务端第 1 页返回的 count 为总数依据翻页并做一致性校验：
    若服务端声明命中 N 行、实际只拉到 < N 行（例如服务端悄悄封顶每页条数、
    或翻页丢页），**直接 FetchError 报错中止**——财务数字宁可报错，绝不静默少算。
    服务端若未给 count，则退回“短页即止”的旧行为。
    """
    fc = filter_controls or []
    out: list[dict] = []
    server_total: int | None = None
    page = 1
    while page <= MAX_PAGES:
        body = {
            "worksheetId": worksheet_id,
            "appId": app_id,
            "pageSize": PAGE_SIZE,
            "pageIndex": page,
            "status": 1,
            "sortControls": [],
            "notGetTotal": page > 1,
            "searchType": 1,
            "keyWords": "",
            "filterControls": fc,
            "fastFilters": [],
            "navGroupFilters": [],
        }
        d = post("Worksheet/GetFilterRows", body).get("data") or {}
        rows = d.get("data") or []
        if page == 1:
            total = d.get("count")
            if isinstance(total, int) and total >= 0:
                server_total = total
        out.extend(rows)
        if server_total is not None:
            # 已知总数：拉够总数、或某页空了就停
            if len(out) >= server_total or not rows:
                break
        elif len(rows) < PAGE_SIZE:
            # 未知总数：退回短页即止
            break
        page += 1
    else:
        raise FetchError(f"翻页超过安全上限 {MAX_PAGES} 页仍未拉完")
    # 一致性闸：服务端声明总数 ≠ 实抓行数 → 报错，绝不静默上报错误数字
    if server_total is not None and len(out) != server_total:
        raise FetchError(
            f"抓取行数与服务端声明不一致：服务端 count={server_total}，实抓={len(out)} 行。"
            "疑似分页封顶/翻页丢数，已中止以免上报少算的下单额。"
        )
    return out


def login(zy: dict, headless: bool = True) -> tuple[str, str | None]:
    """Playwright 无头登录，返回 (md_pss_id, account_id|None)。"""
    base = zy.get("base_url") or ZHIYUN_DEFAULTS["base_url"]
    user = zy.get("username")
    pwd = zy.get("password")
    if not (base and user and pwd):
        raise LoginError("智云配置缺 base_url/username/password，无法自动登录")
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise LoginError(
            "未安装 playwright。请执行：pip install playwright && playwright install chromium"
        ) from e

    try:
        with sync_playwright() as p:
            br = p.chromium.launch(headless=headless)
            try:
                ctx = br.new_context(ignore_https_errors=True)
                pg = ctx.new_page()
                pg.goto(base, wait_until="networkidle", timeout=LOGIN_TIMEOUT_MS)
                pg.fill(ACCOUNT_SEL, user)
                pg.fill(PASSWORD_SEL, pwd)
                clicked = False
                for sel in LOGIN_BTN_SELECTORS:
                    try:
                        pg.click(sel, timeout=2500)
                        clicked = True
                        break
                    except Exception:  # noqa: BLE001
                        continue
                if not clicked:
                    pg.keyboard.press("Enter")
                pg.wait_for_timeout(POST_LOGIN_WAIT_MS)
                token = None
                for c in ctx.cookies():
                    if c["name"] == "md_pss_id" and c.get("value"):
                        token = c["value"]
                        break
                if not token:
                    raise LoginError(
                        f"登录后未取到 md_pss_id（当前地址 {pg.url}，可能账号密码错或需验证码）"
                    )
                account_id = None
                try:
                    account_id = pg.evaluate(
                        "() => { try { return md.global.Account.accountId || null; } catch(e) { return null; } }"
                    )
                except Exception:  # noqa: BLE001
                    account_id = None
                return token, account_id
            finally:
                br.close()
    except LoginError:
        raise
    except Exception as e:  # noqa: BLE001
        raise LoginError(f"登录过程异常（{type(e).__name__}: {e}）") from e


def _is_auth_expired(j) -> bool:
    if not isinstance(j, dict) or j.get("state") not in (0, "0"):
        return False
    msg = str(j.get("exception") or j.get("message") or "")
    return ("登录" in msg) or ("退出" in msg) or ("登陆" in msg)


def _make_post(zy: dict):
    import requests

    state = {"token": zy.get("md_pss_id", ""), "login_failed": False}

    def _headers():
        return {
            "Content-Type": "application/json",
            "Authorization": f"md_pss_id {state['token']}",
            "AccountId": zy.get("account_id", "") or "",
            "X-Requested-With": "XMLHttpRequest",
        }

    def post(path: str, body: dict) -> dict:
        url = f"{zy['base_url'].rstrip('/')}/wwwapi/{path}"
        r = requests.post(url, headers=_headers(), json=body, timeout=120)
        need_relogin = r.status_code == 401
        if not need_relogin and r.status_code == 200:
            try:
                need_relogin = _is_auth_expired(r.json())
            except ValueError:
                need_relogin = False
        if need_relogin and not state["login_failed"]:
            try:
                token, account_id = login(zy)
                state["token"] = token
                zy["md_pss_id"] = token
                if account_id:
                    zy["account_id"] = account_id
            except Exception:
                state["login_failed"] = True
                raise
            r = requests.post(url, headers=_headers(), json=body, timeout=120)
        r.raise_for_status()
        return r.json()

    return post


def get_controls(post: Callable[[str, dict], dict], worksheet_id: str, app_id: str) -> list[dict]:
    """Get worksheet controls template."""
    # Try common paths used by Mingdao HAP
    for path, body in (
        (
            "Worksheet/GetWorksheetInfo",
            {"worksheetId": worksheet_id, "getTemplate": True, "getViews": False, "appId": app_id},
        ),
        (
            "Worksheet/Get",
            {"worksheetId": worksheet_id, "getTemplate": True, "appId": app_id},
        ),
    ):
        try:
            j = post(path, body)
        except Exception:  # noqa: BLE001
            continue
        data = j.get("data") or {}
        # template.controls or controls
        tpl = data.get("template") or data
        controls = tpl.get("controls") or data.get("controls") or []
        if controls:
            return controls
    raise FetchError("无法获取下单表字段模板（controls 为空）")


def fetch_orders(
    start: date | str,
    end: date | str,
    zy_cfg: dict,
) -> list[dict[str, str]]:
    """登录（如需）→ 区间筛 → 翻页 → 中文列名记录 → 客户端再滤。"""
    zy = {
        "base_url": zy_cfg.get("base_url") or ZHIYUN_DEFAULTS["base_url"],
        "app_id": zy_cfg.get("app_id") or ZHIYUN_DEFAULTS["app_id"],
        "orders_worksheet_id": zy_cfg.get("orders_worksheet_id")
        or zy_cfg.get("worksheet_id")
        or ZHIYUN_DEFAULTS["orders_worksheet_id"],
        "username": zy_cfg.get("username") or "",
        "password": zy_cfg.get("password") or "",
        "md_pss_id": zy_cfg.get("md_pss_id") or "",
        "account_id": zy_cfg.get("account_id") or "",
    }
    if not zy["md_pss_id"]:
        if not (zy["username"] and zy["password"]):
            raise LoginError("缺少凭据：请配置 username/password 或 md_pss_id")
        token, account_id = login(zy)
        zy["md_pss_id"] = token
        if account_id:
            zy["account_id"] = account_id

    post = _make_post(zy)
    ws_id = zy["orders_worksheet_id"]
    app_id = zy["app_id"]
    controls = get_controls(post, ws_id, app_id)
    fc = build_date_range_filter(controls, DATE_COL_NAME, start, end)
    raw_rows = fetch_all_rows(post, ws_id, app_id, filter_controls=fc)
    records = rows_to_records(raw_rows, controls)
    return client_filter_by_date(records, start, end, DATE_COL_NAME)
