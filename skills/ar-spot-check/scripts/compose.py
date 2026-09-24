#!/usr/bin/env python3
"""把事实、新闻和智云核对合成亮晶要的四页。规则只执行判断.md 里写死的那几条。"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

POOL_SHEET = "待抽查清单"
SUGGEST_SHEET = "建议本次抽"
EXEMPT_SHEET = "豁免与已回款"
NEWS_SHEET = "风险提示"
ZHIYUN_SHEET = "智云核对"
POOL_HEADERS = ["销售", "客户", "订单号", "交付月份", "账龄", "抽查原因", "已回款笔数", "订单数", "台账确认", "旁注"]
SUGGEST_HEADERS = ["销售", "客户", "订单号", "交付月份", "账龄", "应收金额", "档", "原因", "旁注"]
SUGGEST_BANNER = (
    "这周的短名单，不是全部待抽。先丢掉下面四类，再按账龄和金额进线，进不了线的销售补1条。\n"
    "不进：已豁免；已拿到盖章或对公邮件；销售解释里写了坏账；销售和智云都已回款。\n"
    "高危：账龄满24个月，金额满1000。次危：账龄6到23个月，金额满10万。\n"
    "信用期内但要看：账龄1到2个月，金额满1000，并且有一条：新闻风险高、没合同、合同还在流程、季结、半年结，或预计回款日已经过了还没回。\n"
    "补位：这个销售按上面一条都进不去，但仍有待抽的行，硬放1条。不是按金额进来的。\n"
    "账龄0、账龄空、账龄3到5个月、金额不满1000，留在待抽。梁玲玲和梁玲玲-高美杰算一个人。同一档里金额大的在前。"
)
EXEMPT_HEADERS = ["销售", "客户", "订单号", "交付月份", "账龄", "原因", "豁免原因"]
NEWS_HEADERS = ["客户", "新闻摘要", "链接", "说明", "风险等级", "判断原因", "检索日期"]
ZHIYUN_HEADERS = ["销售", "客户", "订单号", "合同归档号", "订单状态", "说明", "销售结算阶段", "智云订单状态核对"]
STATUS_VERDICTS = {"不冲突", "不一致"}
SUGGEST_ORDER = {"高危": 0, "次危": 1, "信用期内但要看": 2, "补位": 3}
ELIGIBLE = {"没查过", "未提供或未反馈", "确认为空", "待你定"}
KEYWORD_MEANINGS = {"客户名称", "销售人员"}
RISK_ORDER = {"高": 0, "中": 1, "低": 2, "无": 3}
RISK_STYLE = {
    "高": ("F4C7C3", "9C1B1B"),
    "中": ("FCE4B3", "8A5A00"),
    "低": ("D9EAD3", "1E6B3A"),
    "无": ("EEEEEE", "666666"),
}
MISSING_FILL = PatternFill("solid", fgColor="FCE4B3")
ZHIYUN_NOTES = {
    "合同归档号": "这张销售单在智云里挂上的合同归档号。没挂上就写未找到。",
    "订单状态": "智云「下单」页这一列。下单里没有这张单就写未找到。",
    "说明": "归档号和订单状态都有了，这里是空的。只在对不上时写原因。",
}
ZHIYUN_WIDTHS = {"销售": 12, "客户": 36, "订单号": 22, "合同归档号": 18, "订单状态": 24, "说明": 42, "销售结算阶段": 18, "智云订单状态核对": 20}
POOL_WIDTHS = {
    "销售": 12, "客户": 36, "订单号": 28, "交付月份": 12, "账龄": 8,
    "抽查原因": 16, "已回款笔数": 12, "订单数": 10, "台账确认": 22, "旁注": 28,
}
EXEMPT_WIDTHS = {"销售": 12, "客户": 36, "订单号": 28, "交付月份": 12, "账龄": 8, "原因": 24, "豁免原因": 36}
NEWS_WIDTHS = {"客户": 36, "新闻摘要": 46, "链接": 28, "说明": 28, "风险等级": 10, "判断原因": 36, "检索日期": 14}
SHEET_WIDTHS = {
    POOL_SHEET: POOL_WIDTHS,
    EXEMPT_SHEET: EXEMPT_WIDTHS,
    NEWS_SHEET: NEWS_WIDTHS,
    ZHIYUN_SHEET: ZHIYUN_WIDTHS,
    SUGGEST_SHEET: {
        "销售": 12, "客户": 36, "订单号": 28, "交付月份": 12, "账龄": 8,
        "应收金额": 14, "档": 16, "原因": 28, "旁注": 28,
    },
}
HEADER_NOTES = {
    POOL_SHEET: {
        "订单号": "这个客户这个月的单号。格子里放不下时，点进去能看全。",
        "抽查原因": "这行为什么还在清单里。待你定在后一段，已豁免在最下面。",
        "已回款笔数": "销售反馈里，这个月已经标了已回款的笔数。没进豁免清单、又整月回完的，不在这一页。",
        "订单数": "这个客户这个月有几个不同的单号。同一单号多行只算一笔。",
        "台账确认": "合规台账里正式确认的原话。没查过就是空的。",
        "旁注": "按订单抽、回了一部分、整月都已回款，或台账里有写不成月份的旧记录。",
    },
    EXEMPT_SHEET: {
        "原因": "已豁免是豁免清单对上的。销售和智云都已回款，是销售结算阶段和智云订单状态两边都写了已回款。",
        "豁免原因": "豁免清单里的原话，一个字不改。",
    },
    NEWS_SHEET: {
        "新闻摘要": "近半年公开报道的一句话。没查到就写未查到。",
        "链接": "点蓝色字打开原文。没有链接就是没查到。",
        "风险等级": "只看这条新闻会不会直接影响把这笔钱收回来。",
        "判断原因": "为什么是这个等级。没查到的不用写。",
    },
    ZHIYUN_SHEET: {
        **ZHIYUN_NOTES,
        "销售结算阶段": "这张单在销售反馈里写的结算阶段。",
        "智云订单状态核对": "模型看销售结算阶段和智云订单状态冲不冲突。没核到就空着。不一致标红。不对银行流水。",
    },
    SUGGEST_SHEET: {
        "档": "模型按技能里的标准定的档。补位表示这个销售没有别的能进线的行，仍抽一条。",
        "应收金额": "这一组销售反馈的应收金额合计。同一档里金额大的在前。",
        "原因": "这一行为什么进这次建议。",
    },
}
PAID_NOTE = "这个月每笔都已回款"
BOTH_PAID = "销售和智云都已回款"



def as_int(value, default=0) -> int:
    if value in (None, ""):
        return default
    if isinstance(value, float) and value == int(value):
        value = int(value)
    return int(value)


def note_marker(text: str) -> bool:
    return "口径" in text or "供参考" in text


def seq_is_number(value) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value or "").strip()
    if not text:
        return False
    try:
        float(text)
        return True
    except ValueError:
        return False


def load_keywords(path: Path | None) -> list[tuple[str, str]]:
    if path is None:
        return []
    if not path.exists():
        raise KeyError("ledger-missing")
    wb = load_workbook(path, read_only=True, data_only=True)
    if "豁免清单" not in wb.sheetnames:
        wb.close()
        raise KeyError("ledger-sheet")
    rows = wb["豁免清单"].iter_rows(values_only=True)
    header = ["" if cell is None else str(cell) for cell in next(rows)]
    key_at = next((i for i, name in enumerate(header) if "关键词" in name or name == "客户"), None)
    why_at = next((i for i, name in enumerate(header) if "原因" in name), None)
    seq_at = next((i for i, name in enumerate(header) if "序号" in name), None)
    found = []
    if key_at is None:
        wb.close()
        raise KeyError("ledger-column")
    for row in rows:
        seq_raw = row[seq_at] if seq_at is not None and seq_at < len(row) else None
        seq = "" if seq_raw is None else str(seq_raw).strip()
        key = str(row[key_at] or "").strip() if key_at < len(row) else ""
        why = str(row[why_at] or "").strip() if why_at is not None and why_at < len(row) else ""
        if note_marker(seq) or note_marker(key):
            break
        if seq_at is not None and seq and not seq_is_number(seq_raw):
            continue
        if key:
            found.append((key, why))
    wb.close()
    return found


def load_config(path: Path) -> dict:
    section = None
    exempt, aliases, drafts = [], [], []
    for line in path.read_text(encoding="utf-8").splitlines():
        raw = line.strip()
        if raw.startswith("## "):
            title = raw[3:].strip()
            if "已确认豁免" in title:
                section = "exempt"
            elif "已确认别名" in title:
                section = "alias"
            elif "草稿" in title:
                section = "draft"
            else:
                section = None
            continue
        if not raw or raw.startswith("#"):
            continue
        if section == "exempt":
            if raw != "客户名称":
                exempt.append(raw)
        elif section == "alias":
            if raw.startswith("写法") or "|" not in raw:
                continue
            left, right = [part.strip() for part in raw.split("|", 1)]
            if left and right and set(left) != {"-"} and set(right) != {"-"}:
                aliases.append((left, right))
        elif section == "draft":
            if raw.startswith("还没人点头") or raw.startswith("待抽"):
                continue
            drafts.append(raw)
    return {"exempt": set(exempt), "aliases": aliases, "drafts": drafts}


def read_sheet(path: Path, title: str) -> tuple[list[str], list[tuple]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    if title not in wb.sheetnames:
        wb.close()
        raise KeyError(title)
    ws = wb[title]
    rows = ws.iter_rows(values_only=True)
    header = ["" if cell is None else str(cell) for cell in next(rows)]
    data = [row for row in rows if any(cell not in (None, "") for cell in row)]
    wb.close()
    return header, data


def read_facts(path: Path) -> list[dict]:
    header, data = read_sheet(path, "事实")
    index = {name: pos for pos, name in enumerate(header)}
    rows = []
    for item in data:
        row = {name: item[index[name]] if index[name] < len(item) else None for name in header}
        rows.append(row)
    return rows


def seen_this_month(raw, check_month: int) -> bool:
    prefix = str(check_month)
    for part in str(raw or "").split("；"):
        digits = "".join(ch for ch in part if ch.isdigit())
        if digits.startswith(prefix):
            return True
    return False


def token_passes(token: str) -> bool:
    if token == "对公邮件":
        return True
    return "盖章" in token and "未" not in token


def judge(hits: int, confirm) -> str:
    if hits == 0:
        return "没查过"
    tokens = [part.strip() for part in str(confirm or "").split("；")]
    if any(token_passes(token) for token in tokens):
        return "已拿到"
    if all(token in {"", "未提供", "未反馈"} for token in tokens):
        if all(token == "" for token in tokens):
            return "确认为空"
        return "未提供或未反馈"
    return "待你定"


def fully_paid(row: dict) -> bool:
    orders = as_int(row.get("订单数"))
    paid = as_int(row.get("已回款订单数"))
    return orders > 0 and paid == orders


def sort_key(row: dict, check_month: int, seen_names: set[str]):
    blank_age = row.get("账龄") in (None, "")
    age = 0 if blank_age else as_int(row.get("账龄"))
    if blank_age:
        band = 3
    elif 1 <= age <= 5:
        band = 1
    elif age >= 6:
        band = 0
    else:
        band = 2
    forward = 0 if str(row.get("客户") or "").strip() not in seen_names else 1
    return (band, forward, -age, as_int(row.get("交付月份")), str(row.get("客户") or ""), str(row.get("订单号") or ""))


def side_note(row: dict) -> str:
    notes = []
    if row.get("粒度") == "订单":
        notes.append("按订单抽")
    orders = as_int(row.get("订单数"))
    paid = as_int(row.get("已回款订单数"))
    if row.get("回款未证实"):
        notes.append("销售标了已回款，智云没有")
    elif orders > 0 and paid == orders:
        notes.append(PAID_NOTE)
    elif 0 < paid < orders:
        notes.append(f"已回款{paid}/{orders}")
    if as_int(row.get("同客户无法识别的台账行")) > 0:
        notes.append("台账里有写不成月份的历史")
    return "；".join(notes)


def load_news(directory: Path) -> list[dict]:
    rows = []
    for path in sorted(directory.glob("*.jsonl")):
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.strip():
                rows.append(json.loads(line))
    return rows


def news_bad(row: dict) -> str:
    blob = " ".join(str(row.get(key) or "") for key in ("summary", "url", "note"))
    if "未检索" in blob or "本次未" in blob:
        return "placeholder"
    url = str(row.get("url") or "").strip()
    if url.startswith("http"):
        return ""
    if "未查到" in blob:
        return ""
    return "nolink"


def canon_name(name: str, aliases: list[tuple[str, str]]) -> str:
    found = str(name or "").strip()
    for source, target in aliases:
        if found == source:
            return target
    return found


def resolve_risk(item: dict) -> str:
    value = str(item.get("risk") or "").strip()
    url = str(item.get("url") or "").strip()
    blob = " ".join(str(item.get(key) or "") for key in ("summary", "note"))
    if value in RISK_ORDER:
        if url.startswith("http") and value == "无":
            return ""
        if not url.startswith("http") and value != "无":
            return ""
        return value
    if not url.startswith("http") and "未查到" in blob:
        return "无"
    return ""


def merge_news(fact_names: list[str], news_rows: list[dict], aliases: list[tuple[str, str]], retrieved: str) -> list[dict]:
    by_customer = {}
    for row in news_rows:
        by_customer.setdefault(str(row.get("customer") or "").strip(), []).append(row)
    grouped: dict[str, list[dict]] = {}
    sources: dict[str, list[str]] = {}
    order = []
    for name in fact_names:
        key = canon_name(name, aliases)
        if key not in grouped:
            grouped[key] = []
            sources[key] = []
            order.append(key)
        if name not in sources[key]:
            sources[key].append(name)
        grouped[key].extend(by_customer.get(name, []))
    merged = []
    for position, key in enumerate(order):
        items = grouped[key]
        summaries, urls, notes = [], [], []
        risk = "无"
        judged = []
        for item in items:
            summary = re.sub(r"https?://\S+", "", str(item.get("summary") or "")).strip(" ，,;；")
            url = str(item.get("url") or "").strip()
            note = str(item.get("note") or "").strip()
            item_risk = resolve_risk(item)
            if not item_risk:
                return []
            if RISK_ORDER[item_risk] < RISK_ORDER[risk]:
                risk = item_risk
            judged.append((item_risk, str(item.get("reason") or "").strip()))
            if summary and summary not in summaries:
                summaries.append(summary)
            if url and url not in urls:
                urls.append(url)
            if note and note not in notes:
                notes.append(note)
        covered = sources[key]
        if covered != [key]:
            notes.insert(0, "\n".join(covered))
        if not urls and summaries == ["未查到"] and notes in ([], ["未查到"]):
            note_text = "未查到"
        else:
            note_text = "；".join(notes)
        merged.append(
            {
                "客户": key,
                "新闻摘要": "未查到" if not urls and summaries == ["未查到"] else "；".join(summaries),
                "链接": "\n".join(urls),
                "说明": note_text,
                "风险等级": risk,
                "判断原因": "" if risk == "无" else "；".join(dict.fromkeys(text for level, text in judged if level == risk and text)),
                "检索日期": retrieved,
                "_序": position,
            }
        )
    merged.sort(key=lambda row: (RISK_ORDER[row["风险等级"]], row["_序"]))
    for row in merged:
        row.pop("_序", None)
    return merged


def ask(text: str) -> int:
    print("status=ask")
    print(f"ask={text}")
    return 2


def write_sheet(ws, header: list[str], rows: list[dict]) -> None:
    bold = Font(name="微软雅黑", bold=True, size=11, color="000000")
    body = Font(name="微软雅黑", size=11, color="000000")
    link_font = Font(name="微软雅黑", size=11, color="0563C1", underline="single")
    header_row = 1
    if ws.title == SUGGEST_SHEET:
        header_row = 2
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(header))
        banner = ws.cell(1, 1, SUGGEST_BANNER)
        banner.font = Font(name="微软雅黑", size=11, color="000000")
        banner.alignment = Alignment(wrap_text=True, vertical="center")
        banner.fill = PatternFill("solid", fgColor="FFF2CC")
        ws.row_dimensions[1].height = 96
    for col, title in enumerate(header, start=1):
        cell = ws.cell(header_row, col, title)
        cell.font = bold
        cell.alignment = Alignment(wrap_text=False, vertical="center")
        note = HEADER_NOTES.get(ws.title, {}).get(title)
        if note:
            cell.comment = Comment(note, "应收抽查", width=240, height=48)
    for index, row in enumerate(rows, start=header_row + 1):
        for col, title in enumerate(header, start=1):
            value = row.get(title, "")
            text = "" if value is None else value
            cell = ws.cell(index, col, text)
            cell.alignment = Alignment(
                wrap_text=ws.title == NEWS_SHEET and title == "新闻摘要",
                vertical="center",
            )
            cell.font = body
            if title == "链接" and isinstance(text, str) and text.startswith("http"):
                cell.hyperlink = text.split("\n", 1)[0]
                cell.font = link_font
            style = RISK_STYLE.get(str(value)) if title == "风险等级" else None
            if style:
                cell.fill = PatternFill("solid", fgColor=style[0])
                cell.font = Font(name="微软雅黑", color=style[1], bold=True, size=11)
            elif ws.title == ZHIYUN_SHEET and title in {"合同归档号", "订单状态"} and text == "未找到":
                cell.fill = MISSING_FILL
            elif ws.title == ZHIYUN_SHEET and title == "智云订单状态核对" and text == "不一致":
                cell.fill = PatternFill("solid", fgColor="F4C7C3")
            elif ws.title == SUGGEST_SHEET and title == "档" and text == "高危":
                cell.fill = PatternFill("solid", fgColor="F4C7C3")
    if ws.title == SUGGEST_SHEET:
        ws.freeze_panes = "C3"
        last_row = max(header_row, len(rows) + header_row)
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(header))}{last_row}"
    else:
        ws.freeze_panes = "B2" if ws.title == NEWS_SHEET else "C2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{max(1, len(rows) + 1)}"
    widths = SHEET_WIDTHS.get(ws.title, {})
    for col, title in enumerate(header, start=1):
        ws.column_dimensions[get_column_letter(col)].width = widths.get(title, 18)


def paid_confirmed(row: dict, pay_index: dict) -> bool:
    if not fully_paid(row):
        return False
    name = str(row.get("客户") or "").strip()
    sos = [part.strip().upper() for part in str(row.get("订单号") or "").split("；") if part.strip()]
    if not sos:
        sos = [part.strip().upper() for part in str(row.get("已回款订单号") or "").split("；") if part.strip()]
    if not sos:
        return False
    return all("已回款" in str(pay_index.get((name, so), "")) for so in sos)


def pay_index_from(header: list, rows: list) -> dict:
    def column(name: str):
        return header.index(name) if name in header else None

    customer_at = column("客户")
    order_at = column("订单号")
    status_at = column("订单状态")
    found = {}
    if customer_at is None or order_at is None:
        return found
    for row in rows:
        customer = str(row[customer_at] or "").strip() if customer_at < len(row) else ""
        order = str(row[order_at] or "").strip().upper() if order_at < len(row) else ""
        status = ""
        if status_at is not None and status_at < len(row) and row[status_at] is not None:
            status = str(row[status_at])
        if customer and order:
            found[(customer, order)] = status
    return found


def so_stage_index(facts: list[dict]) -> dict:
    found = {}
    for row in facts:
        customer = str(row.get("客户") or "").strip()
        packed = str(row.get("单号阶段") or "")
        if packed:
            for part in packed.split("；"):
                if "=" not in part:
                    continue
                order, stage = part.split("=", 1)
                if customer and order.strip():
                    found[(customer, order.strip().upper())] = stage
            continue
        stage = str(row.get("结算阶段") or "")
        for order in str(row.get("订单号") or "").split("；"):
            if customer and order.strip():
                found[(customer, order.strip().upper())] = stage
    return found


def needs_status_check(stage: str, status: str) -> bool:
    status_text = str(status or "").strip()
    stage_text = str(stage or "").strip()
    if not status_text or status_text == "未找到":
        return False
    if not stage_text:
        return False
    return True


def load_judgment(path: Path | None) -> dict:
    if path is None:
        raise KeyError("judgment-missing")
    if not path.exists():
        raise KeyError("judgment-missing")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        raise KeyError("judgment-json")
    if not isinstance(data, dict):
        raise KeyError("judgment-shape")
    alias = {}
    for item in data.get("同一人") or []:
        if not isinstance(item, dict):
            raise KeyError("judgment-shape")
        canon = str(item.get("算作") or "").strip()
        names = item.get("写成") or []
        if not canon or not isinstance(names, list) or not names:
            raise KeyError("judgment-shape")
        for name in names:
            alias[str(name).strip()] = canon
    keywords = []
    for item in data.get("关键词") or []:
        if not isinstance(item, dict):
            raise KeyError("judgment-shape")
        word = str(item.get("词") or "").strip()
        meaning = str(item.get("含义") or "").strip()
        reason = str(item.get("豁免原因") or "").strip()
        if not word or meaning not in KEYWORD_MEANINGS:
            raise KeyError("judgment-keyword")
        keywords.append((word, meaning, reason))
    suggest = []
    for item in data.get("建议") or []:
        if not isinstance(item, dict):
            raise KeyError("judgment-shape")
        band = str(item.get("档") or "").strip()
        if band not in SUGGEST_ORDER:
            raise KeyError("judgment-band")
        suggest.append(
            {
                "销售": str(item.get("销售") or "").strip(),
                "客户": str(item.get("客户") or "").strip(),
                "交付月份": as_int(item.get("交付月份")),
                "订单号": str(item.get("订单号") or "").strip(),
                "档": band,
                "原因": str(item.get("原因") or "").strip(),
            }
        )
    checks = []
    for item in data.get("订单状态核对") or []:
        if not isinstance(item, dict):
            raise KeyError("judgment-shape")
        verdict = str(item.get("核对") or "").strip()
        customer = str(item.get("客户") or "").strip()
        order = str(item.get("订单号") or "").strip().upper()
        if not customer or not order or verdict not in STATUS_VERDICTS:
            raise KeyError("judgment-status")
        checks.append({"客户": customer, "订单号": order, "核对": verdict})
    return {"alias": alias, "keywords": keywords, "suggest": suggest, "checks": checks}


def judgment_matches_ledger(judgment: dict, ledger_keywords: list[tuple[str, str]]) -> str:
    got = {word: (meaning, reason) for word, meaning, reason in judgment["keywords"]}
    expected = {word: reason for word, reason in ledger_keywords}
    if set(got) != set(expected):
        return "判断.json 的关键词要和豁免清单正式行一一对应。口径那几行不要写进去。"
    for word, reason in expected.items():
        if got[word][1] != reason:
            return "判断.json 里的豁免原因必须和豁免清单原话一样。"
    return ""


def keyword_hits(customer: str, sales: str, word: str, meaning: str) -> bool:
    needle = word.casefold()
    if not needle:
        return False
    if meaning == "客户名称":
        return needle in customer.casefold()
    if meaning == "销售人员":
        return needle in sales.casefold()
    return False


def match_exemption(customer: str, sales: str, keywords: list[tuple[str, str, str]]) -> str:
    hits = [
        (len(word), reason)
        for word, meaning, reason in keywords
        if keyword_hits(customer, sales, word, meaning)
    ]
    if not hits:
        return ""
    hits.sort(reverse=True)
    reasons = []
    for _, reason in hits:
        if reason and reason not in reasons:
            reasons.append(reason)
    return "；".join(reasons) or "已确认豁免"


def canon_sales(name: str, alias: dict[str, str]) -> str:
    text = str(name or "").strip()
    return alias.get(text, text)


def fact_matches(row: dict, item: dict) -> bool:
    if str(row.get("销售") or "").strip() != item["销售"]:
        return False
    if str(row.get("客户") or "").strip() != item["客户"]:
        return False
    if as_int(row.get("交付月份")) != item["交付月份"]:
        return False
    if item.get("订单号") and str(row.get("订单号") or "").strip() != item["订单号"]:
        return False
    return True


def suggest_problem(facts: list[dict], pool: list[dict], judgment: dict) -> str:
    alias = judgment["alias"]
    seen = set()
    for item in judgment["suggest"]:
        key = (item["销售"], item["客户"], item["交付月份"], item.get("订单号") or "")
        if key in seen:
            return "建议本次抽里有重复的一行。同一销售、同一客户、同一交付月、同一订单号只留一条。"
        seen.add(key)
        matches = [row for row in facts if fact_matches(row, item)]
        if len(matches) != 1:
            return "建议本次抽里有对不上的一行，或对上了多张单。写上订单号再跑。"
        pool_hit = [row for row in pool if fact_matches(row, item)]
        if len(pool_hit) != 1 or pool_hit[0].get("规则") not in ELIGIBLE:
            return "建议本次抽只能从待抽里还要跟的行里挑。已豁免、已拿到、坏账、两边都已回款不要放进来。"
    covered = {canon_sales(item["销售"], alias) for item in judgment["suggest"]}
    needed = set()
    for row in pool:
        if row.get("规则") in ELIGIBLE:
            needed.add(canon_sales(str(row.get("销售") or ""), alias))
    missing = needed - covered
    if missing:
        print(f"uncovered={len(missing)}")
        return "还有销售一条都没进建议本次抽。同一人按判断.json 合并后再补。"
    return ""


CANNED_CREDIT_REASON = "还在信用期，但有经营风险、没合同、季结或预计回款已过"
CREDIT_MARKS = (
    ("没有合同", "没合同"),
    ("合同还在流程", "合同还在流程"),
    ("季结", "季结"),
    ("半年结", "半年结"),
)


def credit_hits(row: dict, check_month: int, news_high: set[str]) -> list[str]:
    text = str(row.get("标记") or "") + str(row.get("销售解释") or "")
    found = []
    if str(row.get("客户") or "").strip() in news_high:
        found.append("新闻风险高")
    for needle, label in CREDIT_MARKS:
        if needle in text:
            found.append(label)
    expect = str(row.get("预计回款日") or "")
    if len(expect) >= 6 and expect[:6].isdigit() and int(expect[:6]) < check_month:
        found.append("预计回款日已过还没回")
    if row.get("回款未证实"):
        found.append("销售写了已回款，智云没有")
    return found


def specific_reason(row: dict, band: str, written: str, check_month: int, news_high: set[str]) -> str:
    age = as_int(row.get("账龄"))
    if band == "高危" and age >= 24:
        return "账龄满24个月"
    if band == "次危" and 6 <= age <= 23:
        return "账龄6到23个月，金额满10万"
    if band == "补位":
        return "这个销售按线进不去，补1条"
    if band == "信用期内但要看":
        found = credit_hits(row, check_month, news_high)
        if found:
            return "；".join(found)
        text = str(written or "").strip()
        if text and text != CANNED_CREDIT_REASON and "或" not in text:
            return text
        return ""
    return str(written or "").strip()


def materialize_suggest(facts: list[dict], judgment: dict, check_month: int, news_high: set[str]) -> tuple[list[dict], str]:
    picked = []
    missing = 0
    for order, item in enumerate(judgment["suggest"]):
        row = next(row for row in facts if fact_matches(row, item))
        reason = specific_reason(row, item["档"], item.get("原因") or "", check_month, news_high)
        if not reason:
            missing += 1
        amount = row.get("应收金额")
        try:
            amount = round(float(amount or 0), 2)
        except (TypeError, ValueError):
            amount = 0
        picked.append(
            {
                "销售": row.get("销售") or "",
                "客户": row.get("客户") or "",
                "订单号": row.get("订单号") or "",
                "交付月份": row.get("交付月份") or "",
                "账龄": row.get("账龄"),
                "应收金额": amount,
                "档": item["档"],
                "原因": reason,
                "旁注": side_note(row),
                "_序": order,
            }
        )
    if missing:
        print(f"reason_unspecified={missing}")
        return [], "信用期内但要看有行写不出具体原因。在判断.json 里给这一行写真正的原因，不要整档共用一句。"
    picked.sort(key=lambda item: (SUGGEST_ORDER[item["档"]], item["_序"]))
    for item in picked:
        item.pop("_序", None)
    return picked, ""


def build(facts: list[dict], config: dict, check_month: int, keywords: list[tuple[str, str, str]], pay_index: dict | None = None):
    sales_order = []
    for row in facts:
        if row.get("销售") not in sales_order:
            sales_order.append(row.get("销售"))
    pool, exempt, kept = [], [], []
    for row in facts:
        name = str(row.get("客户") or "").strip()
        sales_name = str(row.get("销售") or "").strip()
        if fully_paid(row) and not paid_confirmed(row, pay_index or {}):
            row["回款未证实"] = True
        why = match_exemption(name, sales_name, keywords)
        if why or name in config["exempt"]:
            her = why or "已确认豁免"
            exempt.append((row, "已豁免", str(row.get("订单号") or row.get("已回款订单号") or ""), her))
            copied = dict(row)
            copied["规则"] = "已豁免"
            pool.append(copied)
            continue
        if "坏账" in str(row.get("标记") or "") or "坏账" in str(row.get("销售解释") or ""):
            exempt.append((row, "坏账", str(row.get("订单号") or ""), "销售解释里写了坏账"))
            continue
        if fully_paid(row) and paid_confirmed(row, pay_index or {}):
            exempt.append((row, BOTH_PAID, str(row.get("已回款订单号") or row.get("订单号") or ""), ""))
            continue
        rule = judge(as_int(row.get("台账命中条数")), row.get("台账确认"))
        if rule == "已拿到":
            kept.append(row)
            continue
        copied = dict(row)
        copied["规则"] = rule
        pool.append(copied)
    seen_names = {
        str(row.get("客户") or "").strip()
        for row in facts
        if seen_this_month(row.get("台账抽查日"), check_month)
    }
    ordered = []
    for sales in sales_order:
        own = [row for row in pool if row.get("销售") == sales]
        main = [row for row in own if row["规则"] not in ("待你定", "已豁免")]
        ordered.extend(sorted(main, key=lambda item: sort_key(item, check_month, seen_names)))
    for sales in sales_order:
        own = [row for row in pool if row.get("销售") == sales and row["规则"] == "待你定"]
        ordered.extend(sorted(own, key=lambda item: sort_key(item, check_month, seen_names)))
    for sales in sales_order:
        own = [row for row in pool if row.get("销售") == sales and row["规则"] == "已豁免"]
        ordered.extend(sorted(own, key=lambda item: sort_key(item, check_month, seen_names)))
    seen_draft = set()
    for name in config["drafts"]:
        if name in config["exempt"] or name in seen_draft:
            continue
        seen_draft.add(name)
        exempt.append(({"销售": "", "客户": name, "交付月份": "", "账龄": ""}, "待确认", "", ""))
    matched = set()
    for row in facts:
        customer = str(row.get("客户") or "")
        sales_name = str(row.get("销售") or "")
        for key, meaning, _reason in keywords:
            if keyword_hits(customer, sales_name, key, meaning):
                matched.add(key)
    for key, _meaning, reason in keywords:
        if key not in matched:
            exempt.append(({"销售": "", "客户": key, "交付月份": "", "账龄": ""}, "已豁免", "", str(reason or "").strip()))
    exempt.sort(key=lambda item: (str(item[0].get("销售") or "￿"), str(item[0].get("客户") or ""), as_int(item[0].get("交付月份"), 0)))
    return ordered, exempt, kept


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="合成交给亮晶的四页工作簿")
    parser.add_argument("--facts", type=Path, required=True)
    parser.add_argument("--news", type=Path, required=True)
    parser.add_argument("--zhiyun", type=Path, required=True)
    parser.add_argument("--check-month", type=int, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--config", type=Path, default=Path(__file__).resolve().parents[1] / "config" / "豁免与别名.md")
    parser.add_argument("--retrieved", default="", help="新闻检索日期，写进风险提示")
    parser.add_argument("--ledger", type=Path, default=None, help="她维护的台账，里面有豁免清单")
    parser.add_argument("--judgment", type=Path, default=None, help="模型写的判断.json")
    args = parser.parse_args(argv)
    month = args.check_month
    if not (200001 <= month <= 209912 and 1 <= month % 100 <= 12):
        return ask("抽查月要写成 YYYYMM。")
    try:
        facts = read_facts(args.facts)
    except KeyError:
        return ask("事实表里没有「事实」这一页。")
    if any(str(row.get("账龄冲突") or "") == "是" for row in facts):
        return ask("有客户月账龄不一致。先问亮晶，不要取平均，也不要排进待抽。")
    config = load_config(args.config)
    try:
        news_rows = load_news(args.news)
    except json.JSONDecodeError:
        return ask("新闻 jsonl 有一行读不了。")
    bad = sum(1 for row in news_rows if news_bad(row))
    if bad:
        print(f"news_bad={bad}")
        return ask("新闻还有没检索、或既没有链接也没有写未查到的行。")
    try:
        zhiyun_header, zhiyun_rows = read_sheet(args.zhiyun, ZHIYUN_SHEET)
    except KeyError:
        return ask("智云核对文件里没有「智云核对」这一页。")
    fact_names = []
    for row in facts:
        name = str(row.get("客户") or "").strip()
        if name and name not in fact_names:
            fact_names.append(name)
    news_names = {str(row.get("customer") or "").strip() for row in news_rows}
    missing_news = [name for name in fact_names if name not in news_names]
    extra_news = news_names - set(fact_names)
    if missing_news or extra_news:
        print(f"news_missing={len(missing_news)}")
        print(f"news_extra={len(extra_news)}")
        return ask("风险提示要覆盖这期销售反馈里的每个客户，不多也不少。别名只在合成时并成一行。")
    customer_col = next((i for i, name in enumerate(zhiyun_header) if "客户" in name), 0)
    zhiyun_names = {str(row[customer_col]).strip() for row in zhiyun_rows if customer_col < len(row) and row[customer_col]}
    missing_zhiyun = [name for name in fact_names if name not in zhiyun_names]
    if missing_zhiyun:
        print(f"zhiyun_missing={len(missing_zhiyun)}")
        return ask("智云核对还缺销售反馈里的客户。用 Playwright 把这些客户的合同和下单补上，对不上就写未找到。")
    try:
        keywords = load_keywords(args.ledger)
    except KeyError as exc:
        label = str(exc)
        if label == "ledger-missing":
            return ask("找不到台账文件，豁免清单没读到。先别出待抽。")
        if label == "ledger-sheet":
            return ask("台账里没有「豁免清单」这一页。先别出待抽。")
        if label == "ledger-column":
            return ask("豁免清单里没有关键词这一列。先别出待抽。")
        raise
    print(f"exempt_keywords={len(keywords)}")
    try:
        judgment = load_judgment(args.judgment)
    except KeyError as exc:
        label = str(exc)
        if label == "judgment-missing":
            return ask("还没有判断.json。先按技能读事实和豁免清单，写出同一人、关键词含义和本周建议，再跑。")
        if label == "judgment-json":
            return ask("判断.json 读不了。")
        if label == "judgment-keyword":
            return ask("判断.json 里每个关键词都要写含义：客户名称，或销售人员。")
        if label == "judgment-band":
            return ask("建议本次抽的档只能是高危、次危、信用期内但要看、补位。")
        if label == "judgment-status":
            return ask("订单状态核对只能写不冲突或不一致，并且要有客户和订单号。")
        return ask("判断.json 的格式不对。按技能里的样子重写。")
    mismatch = judgment_matches_ledger(judgment, keywords)
    if mismatch:
        return ask(mismatch)
    pay_index = pay_index_from(zhiyun_header, zhiyun_rows)
    pool, exempt, kept = build(facts, config, month, judgment["keywords"], pay_index)
    problem = suggest_problem(facts, pool, judgment)
    if problem:
        return ask(problem)
    fact_customer_order = fact_names
    news_out = merge_news(fact_customer_order, news_rows, config["aliases"], str(args.retrieved))
    if len(news_out) != len({canon_name(name, config["aliases"]) for name in fact_names}):
        return ask("新闻风险等级没标全。有链接的写高、中或低，没查到的写无。")
    pool_out = []
    for row in pool:
        pool_out.append(
            {
                "销售": row.get("销售") or "",
                "客户": row.get("客户") or "",
                "订单号": row.get("订单号") or "",
                "交付月份": row.get("交付月份") or "",
                "账龄": "" if row.get("账龄") in (None, "") else row.get("账龄"),
                "抽查原因": row.get("规则") or "",
                "已回款笔数": as_int(row.get("已回款订单数")),
                "订单数": as_int(row.get("订单数")),
                "台账确认": row.get("台账确认") or "",
                "旁注": side_note(row),
            }
        )
    exempt_out = []
    for row, reason, order_no, why in exempt:
        exempt_out.append(
            {
                "销售": row.get("销售") or "",
                "客户": row.get("客户") or "",
                "订单号": order_no,
                "交付月份": row.get("交付月份") or "",
                "账龄": "" if row.get("账龄") in (None, "") else row.get("账龄"),
                "原因": reason,
                "豁免原因": why,
            }
        )
    zhiyun_out = []
    for row in zhiyun_rows:
        item = {}
        for title in ZHIYUN_HEADERS:
            if title not in zhiyun_header:
                item[title] = ""
                continue
            source = zhiyun_header.index(title)
            item[title] = "" if source >= len(row) or row[source] is None else row[source]
        if not item.get("销售"):
            item["销售"] = ""
        zhiyun_out.append(item)
    sales_of: dict[str, dict[str, int]] = {}
    so_sales: dict[tuple[str, str], str] = {}
    for row in facts:
        customer = str(row.get("客户") or "").strip()
        sales_name = str(row.get("销售") or "").strip()
        if customer and sales_name:
            sales_of.setdefault(customer, {})
            sales_of[customer][sales_name] = sales_of[customer].get(sales_name, 0) + 1
        for so in str(row.get("订单号") or "").split("；"):
            token = so.strip().upper()
            if customer and token and sales_name:
                so_sales[(customer, token)] = sales_name
    for item in zhiyun_out:
        customer = str(item.get("客户") or "").strip()
        token = str(item.get("订单号") or "").strip().upper()
        if not item.get("销售") and (customer, token) in so_sales:
            item["销售"] = so_sales[(customer, token)]
        if not item.get("销售") and customer in sales_of:
            counts = sales_of[customer]
            item["销售"] = max(counts, key=lambda name: counts[name])
    stages = so_stage_index(facts)
    for item in zhiyun_out:
        customer = str(item.get("客户") or "").strip()
        token = str(item.get("订单号") or "").strip().upper()
        stage = stages.get((customer, token), "")
        item["销售结算阶段"] = stage
        status = str(item.get("订单状态") or "")
        if needs_status_check(stage, status):
            item["智云订单状态核对"] = ""
            item["_待核"] = True
        else:
            item["智云订单状态核对"] = ""
    check_index = {}
    for item in judgment["checks"]:
        key = (item["客户"], item["订单号"])
        if key in check_index:
            return ask("订单状态核对里同一张单写了两次。")
        check_index[key] = item["核对"]
    missing = 0
    for item in zhiyun_out:
        item.pop("_待核", None)
        if not needs_status_check(item.get("销售结算阶段"), item.get("订单状态")):
            item["智云订单状态核对"] = ""
            continue
        key = (str(item.get("客户") or "").strip(), str(item.get("订单号") or "").strip().upper())
        verdict = check_index.get(key)
        if verdict not in STATUS_VERDICTS:
            missing += 1
            continue
        item["智云订单状态核对"] = verdict
    if missing:
        print(f"status_unchecked={missing}")
        return ask("还有智云和销售两边都写了状态、但判断.json 没给核对结果的单。先判断再跑。")
    zhiyun_out.sort(key=lambda item: (str(item.get("销售") or "￿"), str(item.get("客户") or ""), str(item.get("订单号") or "")))
    news_high = {row["客户"] for row in news_out if row.get("风险等级") == "高"}
    suggest_out, reason_problem = materialize_suggest(facts, judgment, month, news_high)
    if reason_problem:
        return ask(reason_problem)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    sheets = [
        (POOL_SHEET, POOL_HEADERS, pool_out),
        (SUGGEST_SHEET, SUGGEST_HEADERS, suggest_out),
        (EXEMPT_SHEET, EXEMPT_HEADERS, exempt_out),
        (NEWS_SHEET, NEWS_HEADERS, news_out),
        (ZHIYUN_SHEET, ZHIYUN_HEADERS, zhiyun_out),
    ]
    first = wb.active
    if first is None:
        raise RuntimeError("workbook has no sheet")
    first.title = sheets[0][0]
    write_sheet(first, sheets[0][1], sheets[0][2])
    for title, header, rows in sheets[1:]:
        write_sheet(wb.create_sheet(title), header, rows)
    wb.save(args.out)
    wb.close()
    print("status=ok")
    print(f"groups={len(facts)}")
    print(f"pool={len(pool_out)}")
    print(f"suggest={len(suggest_out)}")
    print(f"exempt={sum(1 for _row, reason, _no, _why in exempt if reason != '待确认')}")
    print(f"draft={sum(1 for _row, reason, _no, _why in exempt if reason == '待确认')}")
    print(f"kept={len(kept)}")
    print(f"news={len(news_out)}")
    for label in ("高", "中", "低", "无"):
        print(f"risk_{label}={sum(1 for row in news_out if row['风险等级'] == label)}")
    print(f"zhiyun_rows={len(zhiyun_out)}")
    print(f"zhiyun_customers={len(zhiyun_names)}")
    print(f"out={args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
