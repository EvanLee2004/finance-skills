#!/usr/bin/env python3
"""登录智云，按客户核对合同和下单，写出 sheet「智云核对」。

密码只从本机 json 读，不写进源码，也不进 stdout。
stdout 只打 status=、rows=、miss=、out=。
不打开回款记录。合同和订单用登录后的只读查询，表从「智云」应用里
CMS/客户管理 下的「合同管理 / 合同」「订单 / 下单」「订单 / 订单明细」认。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from months import clean_name, is_umbrella  # noqa: E402

CONFIG_PATH = Path.home() / ".config" / "finance" / "zhiyun.local.json"
COLUMNS = ["销售", "客户", "订单号", "合同归档号", "订单状态", "说明"]
PAGE_SIZE = 300
MAX_PAGES = 500
LIST_CAP = 30
SUFFIXES = ("股份有限公司", "有限责任公司", "有限公司")
SHEET_NAMES = ("合同", "下单", "订单明细")


class ProbeError(RuntimeError):
    pass


def norm_name(value: str) -> str:
    text = str(value or "").strip()
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)
    text = text.replace("中华人民共和国公安部", "公安部").replace("中国公安部", "公安部")
    return text


def stem_name(value: str) -> str:
    text = norm_name(value)
    changed = True
    while changed:
        changed = False
        for suffix in SUFFIXES:
            if text.endswith(suffix) and len(text) - len(suffix) >= 4:
                text = text[: -len(suffix)]
                changed = True
    return text


def load_config(path: Path) -> dict:
    if not path.is_file():
        raise ProbeError("config")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ProbeError("config")
    for key in ("username", "password", "base_url", "app_id"):
        if not str(data.get(key) or "").strip():
            raise ProbeError("config")
    return data


def load_customers(path: Path) -> list[dict]:
    grouped: dict[str, dict] = {}
    order: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        parts = line.split("\t")
        while len(parts) < 3:
            parts.append("")
        name, so, _month = (parts[0].strip(), parts[1].strip(), parts[2].strip())
        if not name:
            continue
        if name not in grouped:
            cleaned, _sos = clean_name(name)
            grouped[name] = {"name": name, "sos": [], "by_order": is_umbrella(cleaned)}
            order.append(name)
        if so and so not in grouped[name]["sos"]:
            grouped[name]["sos"].append(so)
    if not order:
        raise ProbeError("customers")
    return [grouped[name] for name in order]


def login(base: str, username: str, password: str) -> tuple[str, str | None]:
    """Playwright：首页登录，等到工作台，再点进智云。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise ProbeError("playwright") from exc
    try:
        with sync_playwright() as playwright:
            try:
                browser = playwright.chromium.launch(channel="chrome", headless=True)
            except Exception:
                browser = playwright.chromium.launch(headless=True)
            try:
                context = browser.new_context(ignore_https_errors=True)
                page = context.new_page()
                page.goto(base, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_selector("#txtMobilePhone", timeout=30000)
                page.fill("#txtMobilePhone", username)
                page.fill("input[type=password]", password)
                page.click(".btnForLogin")
                page.get_by_text("工作台").first.wait_for(timeout=45000)
                page.get_by_role("link", name="智云").first.click()
                page.wait_for_timeout(2000)
                token = ""
                for cookie in context.cookies():
                    if cookie.get("name") == "md_pss_id" and cookie.get("value"):
                        token = str(cookie["value"])
                        break
                if not token:
                    raise ProbeError("login")
                account_id = None
                try:
                    account_id = page.evaluate(
                        "() => { try { return md.global.Account.accountId || null } catch (e) { return null } }"
                    )
                except Exception:
                    account_id = None
                return token, str(account_id) if account_id else None
            finally:
                browser.close()
    except ProbeError:
        raise
    except Exception as exc:
        raise ProbeError(type(exc).__name__) from exc


class Client:
    def __init__(self, base: str, token: str, account_id: str | None, app_id: str, blocked: set[str]):
        self.base = base.rstrip("/")
        self.app_id = app_id
        self.blocked = {item for item in blocked if item}
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Content-Type": "application/json",
                "Authorization": f"md_pss_id {token}",
                "AccountId": account_id or "",
                "X-Requested-With": "XMLHttpRequest",
            }
        )

    def post(self, path: str, body: dict) -> dict:
        worksheet_id = str(body.get("worksheetId") or "")
        if worksheet_id and worksheet_id in self.blocked:
            raise ProbeError("blocked")
        url = f"{self.base}/wwwapi/{path}"
        last_error: Exception | None = None
        for attempt in range(3):
            try:
                response = self.session.post(url, json=body, timeout=(10, 180))
                if response.status_code in (502, 503, 504):
                    last_error = ProbeError("http")
                    time.sleep(1 + attempt)
                    continue
                if response.status_code == 401:
                    raise ProbeError("auth")
                response.raise_for_status()
                payload = response.json()
                if isinstance(payload, dict) and payload.get("state") in (0, "0"):
                    raise ProbeError("api")
                return payload
            except ProbeError:
                raise
            except (requests.Timeout, requests.ConnectionError) as exc:
                last_error = exc
                time.sleep(1 + attempt)
            except requests.RequestException as exc:
                raise ProbeError(type(exc).__name__) from exc
        raise ProbeError(type(last_error).__name__ if last_error else "http")

    def controls(self, worksheet_id: str) -> list[dict]:
        payload = self.post(
            "Worksheet/GetWorksheetInfo",
            {
                "worksheetId": worksheet_id,
                "getTemplate": True,
                "getViews": False,
                "appId": self.app_id,
            },
        )
        data = payload.get("data") or {}
        template = data.get("template") or data
        found = template.get("controls") or data.get("controls") or []
        if not found:
            raise ProbeError("controls")
        return found

    def fetch_rows(self, worksheet_id: str) -> list[dict]:
        rows: list[dict] = []
        seen: set[str] = set()
        server_total: int | None = None
        page = 1
        while page <= MAX_PAGES:
            payload = self.post(
                "Worksheet/GetFilterRows",
                {
                    "worksheetId": worksheet_id,
                    "appId": self.app_id,
                    "pageSize": PAGE_SIZE,
                    "pageIndex": page,
                    "status": 1,
                    "sortControls": [],
                    "notGetTotal": page > 1,
                    "searchType": 1,
                    "keyWords": "",
                    "filterControls": [],
                    "fastFilters": [],
                    "navGroupFilters": [],
                },
            )
            data = payload.get("data") or {}
            batch = data.get("data") if isinstance(data, dict) else data
            batch = batch if isinstance(batch, list) else []
            if page == 1 and isinstance(data, dict) and isinstance(data.get("count"), int):
                server_total = data["count"]
            fresh = 0
            for row in batch:
                row_id = str(row.get("rowid") or row.get("_id") or "")
                if row_id and row_id in seen:
                    continue
                if row_id:
                    seen.add(row_id)
                rows.append(row)
                fresh += 1
            if server_total is not None and len(rows) >= server_total:
                break
            if not batch or fresh == 0:
                break
            page += 1
        else:
            raise ProbeError("pages")
        if server_total is not None and len(rows) != server_total:
            raise ProbeError("count")
        return rows

    def search_rows(self, worksheet_id: str, keyword: str) -> list[dict]:
        rows: list[dict] = []
        page = 1
        while page <= 20:
            payload = self.post(
                "Worksheet/GetFilterRows",
                {
                    "worksheetId": worksheet_id,
                    "appId": self.app_id,
                    "pageSize": 50,
                    "pageIndex": page,
                    "status": 1,
                    "sortControls": [],
                    "notGetTotal": False,
                    "searchType": 1,
                    "keyWords": keyword,
                    "filterControls": [],
                    "fastFilters": [],
                    "navGroupFilters": [],
                },
            )
            data = payload.get("data") or {}
            batch = data.get("data") if isinstance(data, dict) else data
            batch = batch if isinstance(batch, list) else []
            rows.extend(batch)
            if len(batch) < 50:
                break
            page += 1
        return rows


def discover_sheets(client: Client) -> dict[str, str]:
    payload = client.post(
        "HomeApp/GetApp",
        {"appId": client.app_id, "getSection": True, "getManager": False, "getLang": False},
    )
    found: dict[str, list[tuple[str, str]]] = {name: [] for name in SHEET_NAMES}

    def walk(node: dict, path: list[str]) -> None:
        current = path + [str(node.get("name") or "")]
        for sheet in node.get("workSheetInfo") or []:
            sheet_name = str(sheet.get("workSheetName") or "")
            if "回款" in sheet_name:
                continue
            if sheet_name in found and sheet.get("type") == 0 and sheet.get("workSheetId"):
                found[sheet_name].append(("/".join(current + [sheet_name]), str(sheet["workSheetId"])))
        for child in node.get("childSections") or []:
            if isinstance(child, dict):
                walk(child, current)

    for section in (payload.get("data") or {}).get("sections") or []:
        if isinstance(section, dict):
            walk(section, [])
    chosen: dict[str, str] = {}
    for name, items in found.items():
        if not items:
            raise ProbeError("menu")
        items.sort(key=lambda item: (0 if ("客户" in item[0] or "CMS" in item[0]) else 1, len(item[0])))
        chosen[name] = items[0][1]
    return chosen


def control_map(controls: list[dict]) -> dict[str, dict]:
    found = {}
    for control in controls:
        name = control.get("controlName")
        if name and name not in found:
            found[name] = control
    return found


def option_map(control: dict) -> dict[str, str]:
    found = {}
    for option in control.get("options") or []:
        key = option.get("key")
        if key:
            found[str(key)] = str(option.get("value") or "").strip()
    return found


def json_list(raw) -> list:
    if isinstance(raw, list):
        return raw
    if isinstance(raw, str):
        text = raw.strip()
        if not text or text[0] not in "[{":
            return []
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            return []
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            return [parsed]
    return []


def relation_names(raw) -> list[str]:
    names = []
    for item in json_list(raw):
        if not isinstance(item, dict):
            continue
        name = str(item.get("name") or item.get("fullname") or "").strip()
        if name and name not in names:
            names.append(name)
    return names


def plain_names(raw) -> list[str]:
    if raw is None:
        return []
    if isinstance(raw, str):
        text = raw.strip()
        if not text:
            return []
        related = relation_names(text)
        if related:
            return related
        if text[0] in "[{":
            return []
        return [text]
    return relation_names(raw)


def linked_ids(raw) -> list[str]:
    found = []
    for item in json_list(raw):
        if not isinstance(item, dict):
            continue
        source = item.get("sourcevalue")
        if isinstance(source, str) and source[:1] in "{[":
            try:
                source = json.loads(source)
            except json.JSONDecodeError:
                source = None
        row_id = ""
        if isinstance(source, dict):
            row_id = str(source.get("rowid") or source.get("_id") or "")
        if row_id and row_id not in found:
            found.append(row_id)
    return found


def control_text(raw, mapping: dict[str, str]) -> str:
    if mapping:
        label = option_label(raw, mapping)
        if label and label != "看不出来":
            return label
    if isinstance(raw, str):
        text = raw.strip()
        if text and not text.startswith("[") and text not in {"None", "[]"}:
            return text
    return ""


def option_label(raw, mapping: dict[str, str]) -> str:
    keys = []
    if isinstance(raw, list):
        keys = [str(item) for item in raw]
    elif isinstance(raw, str) and raw.strip().startswith("["):
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            parsed = None
        if isinstance(parsed, list):
            keys = [str(item) for item in parsed]
    if not keys:
        return "看不出来"
    labels = []
    for key in keys:
        label = mapping.get(key, "")
        if label and label not in labels:
            labels.append(label)
    return "、".join(labels) if labels else "看不出来"


def amount_text(raw) -> str:
    if raw is None or isinstance(raw, bool):
        return "看不出来"
    if isinstance(raw, (int, float)):
        if isinstance(raw, float):
            if raw != raw:
                return "看不出来"
            if abs(raw - round(raw)) < 1e-6:
                return str(int(round(raw)))
            text = format(raw, "f").rstrip("0").rstrip(".")
            return text or "看不出来"
        return str(raw)
    text = str(raw).strip()
    if not text or text in ("[]", "{}"):
        return "看不出来"
    return text


def parse_contracts(rows: list[dict], controls: dict[str, dict]) -> list[dict]:
    needed = ("客户", "合同编号", "合同类型", "合同金额", "合同审批")
    for name in needed:
        if name not in controls:
            raise ProbeError("controls")
    type_map = option_map(controls["合同类型"])
    approval_map = option_map(controls["合同审批"])
    customer_id = controls["客户"]["controlId"]
    full_id = (controls.get("客户全称") or {}).get("controlId")
    number_id = controls["合同编号"]["controlId"]
    type_id = controls["合同类型"]["controlId"]
    amount_id = controls["合同金额"]["controlId"]
    approval_id = controls["合同审批"]["controlId"]
    archive_id = (controls.get("合同归档号") or {}).get("controlId")
    parsed = []
    for row in rows:
        names = relation_names(row.get(customer_id))
        if full_id:
            names.extend(name for name in plain_names(row.get(full_id)) if name not in names)
        number = str(row.get(number_id) or "").strip()
        archive = str(row.get(archive_id) or "").strip() if archive_id else ""
        if archive in ("None", "[]"):
            archive = ""
        parsed.append(
            {
                "id": str(row.get("rowid") or row.get("_id") or ""),
                "no": number,
                "type": option_label(row.get(type_id), type_map),
                "amount": amount_text(row.get(amount_id)),
                "approval": option_label(row.get(approval_id), approval_map),
                "customers": names,
                "archive": archive,
            }
        )
    return parsed


def parse_orders(rows: list[dict], controls: dict[str, dict]) -> list[dict]:
    for name in ("SO", "客户", "订单合同"):
        if name not in controls:
            raise ProbeError("controls")
    so_id = controls["SO"]["controlId"]
    customer_id = controls["客户"]["controlId"]
    customer_rel = (controls.get("客户名称") or {}).get("controlId")
    contract_id = controls["订单合同"]["controlId"]
    archive_id = (controls.get("合同归档号") or {}).get("controlId")
    status_ctrl = controls.get("订单状态")
    status_id = status_ctrl["controlId"] if status_ctrl else ""
    status_map = option_map(status_ctrl) if status_ctrl else {}
    parsed = []
    for row in rows:
        names = plain_names(row.get(customer_id))
        if customer_rel:
            for name in relation_names(row.get(customer_rel)):
                if name not in names:
                    names.append(name)
        so = str(row.get(so_id) or "").strip()
        archive = str(row.get(archive_id) or "").strip() if archive_id else ""
        parsed.append(
            {
                "so": so,
                "customers": names,
                "contract_ids": linked_ids(row.get(contract_id)),
                "archive": archive,
                "status": control_text(row.get(status_id), status_map) if status_id else "",
            }
        )
    return parsed


def _join(values: list[str]) -> str:
    cleaned = []
    for value in values:
        text = str(value or "").strip()
        if text and text not in cleaned:
            cleaned.append(text)
    return "、".join(cleaned)


def _sentence(parts: list[str]) -> str:
    kept = []
    for part in parts:
        text = str(part or "").strip("。")
        if text and text not in kept:
            kept.append(text)
    if "未对照本月应收" not in kept:
        kept.append("未对照本月应收")
    return "。".join(kept) + "。"


def _blank(contract_no: str, contract_type: str, amount: str, approval: str) -> dict:
    return {
        "合同编号": contract_no,
        "合同类型": contract_type,
        "合同金额": amount,
        "合同审批": approval,
    }


def assemble(customers: list[dict], contracts: list[dict], orders: list[dict], detail_found: set[str]) -> tuple[list[dict], int]:
    by_id = {}
    contracts_by_name: dict[str, list[dict]] = defaultdict(list)
    contract_stem: dict[str, set[str]] = defaultdict(set)
    for contract in contracts:
        if contract.get("id"):
            by_id[contract["id"]] = contract
        seen_names = set()
        for customer in contract.get("customers") or []:
            key = norm_name(customer)
            if not key or key in seen_names:
                continue
            seen_names.add(key)
            contracts_by_name[key].append(contract)
            contract_stem[stem_name(key)].add(key)

    orders_by_name: dict[str, list[dict]] = defaultdict(list)
    orders_by_so: dict[str, list[dict]] = defaultdict(list)
    orders_by_contract: dict[str, list[dict]] = defaultdict(list)
    order_stem: dict[str, set[str]] = defaultdict(set)
    for order in orders:
        seen_names = set()
        for customer in order.get("customers") or []:
            key = norm_name(customer)
            if not key or key in seen_names:
                continue
            seen_names.add(key)
            orders_by_name[key].append(order)
            order_stem[stem_name(key)].add(key)
        if order.get("so"):
            orders_by_so[order["so"].upper()].append(order)
        for contract_id in order.get("contract_ids") or []:
            orders_by_contract[contract_id].append(order)

    rows: list[dict] = []
    miss = 0
    for customer in customers:
        built = _rows_for_customer(
            customer,
            by_id,
            contracts_by_name,
            contract_stem,
            orders_by_name,
            orders_by_so,
            orders_by_contract,
            order_stem,
            detail_found,
        )
        if built and all(item["合同编号"] == "未找到" and item["订单号"] == "未找到" for item in built):
            miss += 1
        rows.extend(built)
    return rows, miss


def _lookup_names(key: str, exact: dict, stemmed: dict[str, set[str]]) -> tuple[set[str], str]:
    if key in exact:
        return {key}, ""
    candidates = stemmed.get(stem_name(key), set())
    if len(candidates) == 1:
        only = next(iter(candidates))
        if only != key:
            return {only}, "suffix"
    if len(candidates) > 1 and stem_name(key) != key:
        return set(), "ambiguous"
    return set(), ""


def _names_close(feedback: str, contract_customers: list[str]) -> bool:
    key = norm_name(feedback)
    stemmed = stem_name(key)
    for customer in contract_customers:
        if norm_name(customer) == key or stem_name(customer) == stemmed:
            return True
    return False


def _rows_for_customer(
    customer: dict,
    by_id: dict,
    contracts_by_name,
    contract_stem,
    orders_by_name,
    orders_by_so,
    orders_by_contract,
    order_stem,
    detail_found: set[str],
) -> list[dict]:
    name = customer["name"]
    feedback = list(customer.get("sos") or [])
    by_order = bool(customer.get("by_order"))
    key = norm_name(name)
    contract_names, contract_how = _lookup_names(key, contracts_by_name, contract_stem)
    order_names, order_how = _lookup_names(key, orders_by_name, order_stem)
    how = contract_how or order_how
    name_contracts = []
    seen_contract = set()
    for matched in contract_names:
        for contract in contracts_by_name.get(matched, []):
            if contract["id"] not in seen_contract:
                seen_contract.add(contract["id"])
                name_contracts.append(contract)
    name_contracts.sort(key=lambda item: (item.get("no") or "", item.get("id") or ""))
    name_ids = {item["id"] for item in name_contracts}

    own_orders = []
    seen_order = set()
    for matched in order_names | contract_names | ({key} if key in orders_by_name else set()):
        for order in orders_by_name.get(matched, []):
            marker = id(order)
            if marker not in seen_order:
                seen_order.add(marker)
                own_orders.append(order)

    so_contracts: dict[str, list[str]] = {}
    unlinked: list[str] = []
    detail_sos: list[str] = []
    missing: list[str] = []
    for so in feedback:
        hits = orders_by_so.get(so.upper(), [])
        if not hits:
            if so.upper() in detail_found or so in detail_found:
                detail_sos.append(so)
            else:
                missing.append(so)
            continue
        contract_ids = []
        for hit in hits:
            for contract_id in hit.get("contract_ids") or []:
                if contract_id not in contract_ids:
                    contract_ids.append(contract_id)
        if contract_ids:
            so_contracts[so] = contract_ids
        else:
            unlinked.append(so)

    # 公安部按订单号看，不把名下每一份合同都摊开。
    emit_ids = [] if by_order else [item["id"] for item in name_contracts]
    for contract_ids in so_contracts.values():
        for contract_id in contract_ids:
            if contract_id not in emit_ids:
                emit_ids.append(contract_id)

    rows = []
    suffix_note = "客户名和页面只差公司后缀，只对上这一家" if how == "suffix" else ""
    ambiguous_note = "相近客户名不止一家，合同不替你挑" if how == "ambiguous" and not contract_names else ""

    for contract_id in emit_ids:
        contract = by_id.get(contract_id)
        linked_feedback = [so for so, contract_ids in so_contracts.items() if contract_id in contract_ids]
        extras = []
        seen_so = {so.upper() for so in linked_feedback}
        for order in orders_by_contract.get(contract_id, []):
            so = str(order.get("so") or "").strip()
            if so and so.upper() not in seen_so:
                seen_so.add(so.upper())
                extras.append(so)
        if contract is None:
            rows.append(
                {
                    "客户": name,
                    "订单号": _join(linked_feedback) or "看不出来",
                    "合同编号": "看不出来",
                    "合同类型": "看不出来",
                    "合同金额": "看不出来",
                    "合同审批": "看不出来",
                    "说明": _sentence(["单号关联的合同记录对不上", suffix_note]),
                }
            )
            continue
        bits = ["找到合同"]
        if contract.get("type") == "年度框架协议":
            bits.append("年度框架协议")
        if suffix_note and contract_id in name_ids:
            bits.append(suffix_note)
        if linked_feedback:
            bits.append(f"销售反馈单号在下单里对上{len(linked_feedback)}个")
            show = list(linked_feedback)
            hidden = 0
            for so in extras:
                if len(show) >= LIST_CAP:
                    hidden += 1
                else:
                    show.append(so)
            if hidden:
                bits.append(f"这份合同还关联另外{hidden}笔下单，没有逐条写入")
            order_cell = _join(show)
        elif feedback:
            bits.append("销售反馈的单号没有挂在这份合同上")
            if extras:
                bits.append(f"这份合同另外关联下单{len(extras)}笔")
            order_cell = "未找到"
        elif not extras:
            bits.append("下单里没有挂上单号")
            order_cell = "未找到"
        elif len(extras) <= LIST_CAP:
            bits.append(f"下单里对上{len(extras)}个单号")
            order_cell = _join(extras)
        else:
            bits.append(f"这份合同关联下单{len(extras)}笔，看不出本次要哪几笔")
            order_cell = "看不出来"
        if contract_id not in name_ids and not _names_close(name, contract.get("customers") or []):
            bits.append("单号挂到的合同，客户名和销售反馈不一致")
        rows.append(
            {
                "客户": name,
                "订单号": order_cell,
                "合同编号": contract.get("no") or "看不出来",
                "合同类型": contract.get("type") or "看不出来",
                "合同金额": contract.get("amount") or "看不出来",
                "合同审批": contract.get("approval") or "看不出来",
                "说明": _sentence(bits),
            }
        )

    if by_order and not emit_ids and not feedback:
        rows.append(
            {
                "客户": name,
                "订单号": "看不出来",
                **_blank("看不出来", "看不出来", "看不出来", "看不出来"),
                "说明": _sentence(["按订单号看", "销售反馈没有单号，合同不整户列出"]),
            }
        )
    elif not emit_ids and not feedback:
        own_sos = []
        for order in own_orders:
            so = str(order.get("so") or "").strip()
            if so and so not in own_sos:
                own_sos.append(so)
        if ambiguous_note:
            order_cell = _join(own_sos[:LIST_CAP]) if own_sos and len(own_sos) <= LIST_CAP else ("看不出来" if own_sos else "未找到")
            rows.append(
                {
                    "客户": name,
                    "订单号": order_cell,
                    "合同编号": "看不出来",
                    "合同类型": "看不出来",
                    "合同金额": "看不出来",
                    "合同审批": "看不出来",
                    "说明": _sentence([ambiguous_note, f"客户名下有下单{len(own_sos)}笔" if own_sos else "订单也未找到"]),
                }
            )
        elif not own_sos:
            rows.append(
                {
                    "客户": name,
                    "订单号": "未找到",
                    **_blank("未找到", "未找到", "看不出来", "看不出来"),
                    "说明": _sentence(["合同和订单都未找到"]),
                }
            )
        elif len(own_sos) <= LIST_CAP:
            rows.append(
                {
                    "客户": name,
                    "订单号": _join(own_sos),
                    **_blank("未找到", "未找到", "看不出来", "看不出来"),
                    "说明": _sentence([suffix_note, f"合同未找到。客户名下有下单{len(own_sos)}笔"]),
                }
            )
        else:
            rows.append(
                {
                    "客户": name,
                    "订单号": "看不出来",
                    **_blank("未找到", "未找到", "看不出来", "看不出来"),
                    "说明": _sentence([suffix_note, f"合同未找到。客户名下有下单{len(own_sos)}笔，看不出本次要哪几笔"]),
                }
            )

    if unlinked:
        rows.append(
            {
                "客户": name,
                "订单号": _join(unlinked),
                **_blank("未找到", "未找到", "看不出来", "看不出来"),
                "说明": _sentence([f"这{len(unlinked)}个单号在下单里有，没有关联合同"]),
            }
        )
    if detail_sos:
        rows.append(
            {
                "客户": name,
                "订单号": _join(detail_sos),
                **_blank("看不出来", "看不出来", "看不出来", "看不出来"),
                "说明": _sentence(["下单里没有，订单明细里有这些单号", "合同看不出来"]),
            }
        )
    if missing:
        rows.append(
            {
                "客户": name,
                "订单号": _join(missing),
                **_blank("未找到", "未找到", "看不出来", "看不出来"),
                "说明": _sentence(["这些单号在下单和订单明细里都未找到"]),
            }
        )
    if not rows:
        rows.append(
            {
                "客户": name,
                "订单号": "未找到",
                **_blank("未找到", "未找到", "看不出来", "看不出来"),
                "说明": _sentence([ambiguous_note or "合同和订单都未找到"]),
            }
        )

    if by_order and rows:
        hidden = [item for item in name_contracts if item["id"] not in set(emit_ids)]
        if hidden:
            note = f"按订单号看。名下另有{len(hidden)}份合同没有挂这些单号，没有逐份写入"
            rows[0]["说明"] = rows[0]["说明"].replace("未对照本月应收。", note + "。未对照本月应收。", 1)

    present = set()
    for item in rows:
        for part in str(item["订单号"]).split("、"):
            text = part.strip()
            if text:
                present.add(text.upper())
    absent = [so for so in feedback if so.upper() not in present]
    if absent:
        rows.append(
            {
                "客户": name,
                "订单号": _join(absent),
                **_blank("未找到", "未找到", "看不出来", "看不出来"),
                "说明": _sentence(["这些单号没有写进上面的行，在下单和订单明细里未找到"]),
            }
        )
    return rows


def detail_hits_for(client: Client, worksheet_id: str, controls: dict[str, dict], missing: list[str]) -> set[str]:
    if "SO" not in controls or not missing:
        return set()
    so_id = controls["SO"]["controlId"]
    found = set()
    for so in missing:
        try:
            rows = client.search_rows(worksheet_id, so)
        except ProbeError:
            continue
        for row in rows:
            value = str(row.get(so_id) or "").strip()
            if value.upper() == so.upper():
                found.add(so.upper())
                break
    return found


HEADER_NOTES = {
    "合同归档号": "这张销售单在智云里挂上的合同归档号。没挂上就写未找到。",
    "订单状态": "智云「下单」页这一列。下单里没有这张单就写未找到。",
    "说明": "归档号和订单状态都有了，这里是空的。只在对不上时写原因。",
}
MISSING_FILL = PatternFill("solid", fgColor="FCE4B3")


def _clean_token(value) -> str:
    text = str(value or "").strip()
    if text in {"", "None", "[]", "{}", "null", "NULL", "看不出来", "未找到"}:
        return ""
    return text


def write_xlsx(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    if sheet is None:
        raise ProbeError("xlsx")
    sheet.title = "智云核对"
    header_font = Font(name="微软雅黑", bold=True, size=11)
    body_font = Font(name="微软雅黑", size=11)
    wrap = Alignment(wrap_text=False, vertical="center")
    for col, title in enumerate(COLUMNS, start=1):
        cell = sheet.cell(1, col, title)
        cell.font = header_font
        cell.alignment = wrap
        note = HEADER_NOTES.get(title)
        if note:
            cell.comment = Comment(note, "应收抽查", width=240, height=48)
    for index, item in enumerate(rows, start=2):
        for col, title in enumerate(COLUMNS, start=1):
            value = item.get(title, "")
            cell = sheet.cell(index, col, "" if value is None else str(value))
            cell.font = body_font
            cell.alignment = wrap
            if title in {"合同归档号", "订单状态"} and cell.value == "未找到":
                cell.fill = MISSING_FILL
    widths = [14, 36, 22, 18, 24, 42]
    for col, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    last = get_column_letter(len(COLUMNS))
    sheet.auto_filter.ref = f"A1:{last}{max(1, len(rows) + 1)}"
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)
    workbook.close()


def _hit_archives(hit: dict, by_id: dict[str, dict]) -> list[str]:
    linked: list[str] = []
    for contract_id in hit.get("contract_ids") or []:
        archive = _clean_token((by_id.get(contract_id) or {}).get("archive"))
        if archive and archive not in linked:
            linked.append(archive)
    if linked:
        return linked
    own = _clean_token(hit.get("archive"))
    return [own] if own else []


def finance_rows(
    customers: list[dict],
    contracts: list[dict],
    orders: list[dict],
    detail_found: set[str] | None = None,
) -> list[dict]:
    """销售反馈里一个单号一行。归档号是这张单挂上的合同，状态是下单页那一列。"""
    by_id = {contract["id"]: contract for contract in contracts if contract.get("id")}
    orders_by_so: dict[str, list[dict]] = defaultdict(list)
    for order in orders:
        so = _clean_token(order.get("so"))
        if so:
            orders_by_so[so.upper()].append(order)
    known_detail = {str(so).strip().upper() for so in (detail_found or set()) if str(so).strip()}
    rows: list[dict] = []
    for customer in customers:
        name = customer["name"]
        sos = [str(so).strip() for so in (customer.get("sos") or []) if str(so).strip()]
        if not sos:
            rows.append(_finance_row(name, "未找到", "未找到", "未找到", "销售反馈没有单号。"))
            continue
        for so in sos:
            hits = orders_by_so.get(so.upper(), [])
            if not hits:
                if so.upper() in known_detail:
                    note = "下单里没有这个单号，订单明细里有。"
                else:
                    note = "下单里没有这个单号。"
                rows.append(_finance_row(name, so, "未找到", "未找到", note))
                continue
            statuses: list[str] = []
            archives: list[str] = []
            for hit in hits:
                status = _clean_token(hit.get("status"))
                if status and status not in statuses:
                    statuses.append(status)
                for archive in _hit_archives(hit, by_id):
                    if archive not in archives:
                        archives.append(archive)
            archive_text = "、".join(archives) if archives else "未找到"
            status_text = "、".join(statuses) if statuses else "未找到"
            notes: list[str] = []
            if len(archives) > 1:
                notes.append("这张单挂了多份合同。")
            if len(statuses) > 1:
                notes.append("下单里这张单有多个状态。")
            if not archives and not statuses:
                notes.append("下单有这张单，没有订单状态，也没有合同归档号。")
            elif not archives:
                notes.append("下单有这张单，没有合同归档号。")
            elif not statuses:
                notes.append("有合同归档号，下单没有订单状态。")
            rows.append(_finance_row(name, so, archive_text, status_text, "".join(notes)))
    return rows


def _finance_row(name: str, so: str, archive: str, status: str, note: str) -> dict:
    return {
        "销售": "",
        "客户": name,
        "订单号": so,
        "合同归档号": archive,
        "订单状态": status,
        "说明": note,
    }


def annotate_rows(rows: list[dict], contracts: list[dict], orders: list[dict]) -> None:
    """结果页要归档号和订单状态，不把合同编号写给她。"""
    archive_by_no: dict[str, str] = {}
    for contract in contracts:
        number = str(contract.get("no") or "").strip()
        archive = str(contract.get("archive") or "").strip()
        if number and archive and number not in archive_by_no:
            archive_by_no[number] = archive
    status_by_so: dict[str, str] = {}
    for order in orders:
        so = str(order.get("so") or "").upper()
        status = str(order.get("status") or "").strip()
        if so and status and so not in status_by_so:
            status_by_so[so] = status
    for item in rows:
        number = str(item.get("合同编号") or "").strip()
        if number in {"", "未找到"}:
            item["合同归档号"] = "未找到"
        elif number == "看不出来":
            item["合同归档号"] = "看不出来"
        else:
            item["合同归档号"] = archive_by_no.get(number) or "看不出来"
        found = []
        for so in re.findall(r"SO\d+", str(item.get("订单号") or ""), flags=re.I):
            status = status_by_so.get(so.upper())
            if status and status not in found:
                found.append(status)
        if found:
            item["订单状态"] = "、".join(found)
        elif "没有挂在这份合同" in str(item.get("说明") or ""):
            item["订单状态"] = "没挂在这份合同上"
        elif str(item.get("订单号") or "") in {"", "未找到"}:
            item["订单状态"] = "未找到"
        else:
            item["订单状态"] = "看不出来"
        item.setdefault("销售", "")


def attach_archive(contracts: list[dict], orders: list[dict]) -> None:
    """订单没关联合同、但归档号只对应一份合同时，用归档号接上。"""
    index: dict[str, set[str]] = defaultdict(set)
    for contract in contracts:
        archive = str(contract.get("archive") or "").strip()
        if archive and contract.get("id"):
            index[archive].add(contract["id"])
    unique = {archive: next(iter(ids)) for archive, ids in index.items() if len(ids) == 1}
    for order in orders:
        if order.get("contract_ids"):
            continue
        contract_id = unique.get(str(order.get("archive") or "").strip())
        if contract_id:
            order["contract_ids"] = [contract_id]


def missing_feedback_sos(customers: list[dict], orders: list[dict]) -> list[str]:
    known = {str(order.get("so") or "").upper() for order in orders if order.get("so")}
    missing = []
    for customer in customers:
        for so in customer.get("sos") or []:
            if so.upper() not in known and so not in missing:
                missing.append(so)
    return missing


def run(customers_path: Path, out_path: Path, config_path: Path, limit: int | None) -> tuple[int, int]:
    config = load_config(config_path)
    customers = load_customers(customers_path)
    if limit is not None:
        customers = customers[:limit]
    token, account_id = login(str(config["base_url"]).rstrip("/"), str(config["username"]), str(config["password"]))
    del config["password"]
    client = Client(
        str(config["base_url"]),
        token,
        account_id or str(config.get("account_id") or ""),
        str(config["app_id"]),
        {str(config.get("receipts_worksheet_id") or "")},
    )
    sheets = discover_sheets(client)
    contract_controls = control_map(client.controls(sheets["合同"]))
    order_controls = control_map(client.controls(sheets["下单"]))
    detail_controls = control_map(client.controls(sheets["订单明细"]))
    contracts = parse_contracts(client.fetch_rows(sheets["合同"]), contract_controls)
    orders = parse_orders(client.fetch_rows(sheets["下单"]), order_controls)
    attach_archive(contracts, orders)
    detail_found = detail_hits_for(client, sheets["订单明细"], detail_controls, missing_feedback_sos(customers, orders))
    rows = finance_rows(customers, contracts, orders, detail_found)
    names = {item["name"] for item in customers}
    written = {item["客户"] for item in rows}
    if names - written:
        raise ProbeError("rows")
    miss = 0
    by_customer: dict[str, list[dict]] = defaultdict(list)
    for item in rows:
        by_customer[item["客户"]].append(item)
    for items in by_customer.values():
        if all(item["合同归档号"] == "未找到" and item["订单状态"] == "未找到" for item in items):
            miss += 1
    write_xlsx(out_path, rows)
    return len(rows), miss


def self_test() -> None:
    contracts = [
        {
            "id": "c1",
            "no": "HT1",
            "type": "年度框架协议",
            "amount": "100",
            "approval": "财务归档",
            "customers": ["甲有限公司"],
            "archive": "1",
        },
        {
            "id": "c2",
            "no": "HT2",
            "type": "项目制协议",
            "amount": "看不出来",
            "approval": "审批中",
            "customers": ["东方测试股份有限公司"],
            "archive": "",
        },
    ]
    orders = [
        {"so": "SO00000001", "customers": ["甲有限公司"], "contract_ids": ["c1"], "archive": "1"},
        {"so": "SO00000002", "customers": ["甲有限公司"], "contract_ids": [], "archive": ""},
        {"so": "SO00000003", "customers": ["东方测试股份有限公司"], "contract_ids": ["c2"], "archive": ""},
    ]
    customers = [
        {"name": "甲有限公司", "sos": ["SO00000001"]},
        {"name": "东方测试有限公司", "sos": []},
        {"name": "丙公司", "sos": ["SO00000009"]},
        {"name": "丁有限公司", "sos": []},
    ]
    rows, miss = assemble(customers, contracts, orders, set())
    by_customer = defaultdict(list)
    for row in rows:
        by_customer[row["客户"]].append(row)
    assert "SO00000001" in by_customer["甲有限公司"][0]["订单号"]
    assert by_customer["甲有限公司"][0]["合同编号"] == "HT1"
    assert "年度框架协议" in by_customer["甲有限公司"][0]["说明"]
    assert by_customer["东方测试有限公司"][0]["合同编号"] == "HT2"
    assert "SO00000003" in by_customer["东方测试有限公司"][0]["订单号"]
    assert "后缀" in by_customer["东方测试有限公司"][0]["说明"]
    assert by_customer["丙公司"][0]["合同编号"] == "未找到"
    assert by_customer["丙公司"][0]["订单号"] == "SO00000009"
    assert by_customer["丁有限公司"][0]["订单号"] == "未找到"
    assert by_customer["丁有限公司"][0]["合同编号"] == "未找到"
    assert miss == 1
    rows2, _miss2 = assemble(customers[:1], contracts, orders, {"SO00000009"})
    assert rows2[0]["合同编号"] == "HT1"
    umbrella_contracts = [
        {
            "id": "u1",
            "no": "HU1",
            "type": "项目制协议",
            "amount": "1",
            "approval": "财务归档",
            "customers": ["公安部"],
            "archive": "",
        },
        {
            "id": "u2",
            "no": "HU2",
            "type": "年度框架协议",
            "amount": "看不出来",
            "approval": "财务归档",
            "customers": ["公安部"],
            "archive": "",
        },
    ]
    umbrella_orders = [
        {"so": "SO10000001", "customers": ["公安部"], "contract_ids": ["u1"], "archive": ""},
    ]
    umbrella_rows, umbrella_miss = assemble(
        [{"name": "公安部", "sos": ["SO10000001"], "by_order": True}],
        umbrella_contracts,
        umbrella_orders,
        set(),
    )
    assert [item["合同编号"] for item in umbrella_rows] == ["HU1"]
    assert "SO10000001" in umbrella_rows[0]["订单号"]
    assert "按订单号看" in umbrella_rows[0]["说明"]
    assert "1份" in umbrella_rows[0]["说明"]
    assert umbrella_miss == 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="按客户核对智云合同和下单")
    parser.add_argument("--customers", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--config", type=Path, default=CONFIG_PATH)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    if args.self_test:
        pass
    elif not args.customers or not args.out:
        print("status=fail")
        print("rows=0")
        print("miss=0")
        print("out=")
        return 2
    if args.self_test:
        try:
            self_test()
        except Exception:
            print("status=fail")
            print("rows=0")
            print("miss=0")
            print("out=")
            return 1
        print("status=ok")
        print("rows=0")
        print("miss=0")
        print("out=")
        return 0
    if not args.config.exists():
        print("status=ask")
        print("ask=本机没有智云登录。把账号写进配置文件，或这次把账号发我，再查合同归档号和订单状态。")
        print("rows=0")
        print("miss=0")
        print("out=")
        return 2
    try:
        count, miss = run(args.customers, args.out, args.config, args.limit)
    except Exception as exc:
        print("status=fail")
        print("rows=0")
        print("miss=0")
        print("out=")
        label = str(exc) if isinstance(exc, ProbeError) else type(exc).__name__
        if not label.isascii() or len(label) > 40:
            label = type(exc).__name__
        sys.stderr.write(label + "\n")
        return 1
    print("status=ok")
    print(f"rows={count}")
    print(f"miss={miss}")
    print(f"out={args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
