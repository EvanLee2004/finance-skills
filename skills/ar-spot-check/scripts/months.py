"""台账交付月份的拆法。只做这一件确定的事，不下抽查结论。"""

from __future__ import annotations

import re

SO_RE = re.compile(r"SO\d+", re.I)
PAREN_SO_RE = re.compile(
    r"[（(]\s*SO\d+(?:\s*[、,，]\s*SO\d+)*\s*[)）]",
    re.I,
)


def yy(token: str) -> int:
    n = int(token)
    return n if n >= 1000 else 2000 + n


def check_ym(date) -> tuple[int, int] | None:
    if date is None or str(date).strip() == "":
        return None
    if isinstance(date, float) and date == int(date):
        date = int(date)
    s = str(int(date)) if isinstance(date, int) else str(date).strip()
    if re.fullmatch(r"20\d{6}", s):
        return int(s[:4]), int(s[4:6])
    return None


def year_for_month(month: int, cy: int, cm: int) -> int:
    return cy - 1 if month > cm else cy


def months_between(y1: int, m1: int, y2: int, m2: int) -> list[int]:
    out = []
    y, m = y1, m1
    guard = 0
    while (y, m) <= (y2, m2):
        out.append(y * 100 + m)
        m += 1
        if m == 13:
            m = 1
            y += 1
        guard += 1
        if guard > 48:
            raise ValueError(f"range too long {y1}-{m1} .. {y2}-{m2}")
    return out


def span_without_year(start: int, end: int, cy: int, cm: int) -> list[int]:
    """连续月份。整段结束月不晚于抽查月，取能满足这一点的最近一段。"""
    if end >= start:
        year = cy if end <= cm else cy - 1
        return [year * 100 + m for m in range(start, end + 1)]
    end_year = cy if end <= cm else cy - 1
    return months_between(end_year - 1, start, end_year, end)


def parse_month(raw, date):
    """返回 (months, kind)。months 为 None 表示留原值，不当成任何一个月已查过。"""
    if raw is None or str(raw).strip() == "":
        return None, "empty"
    if isinstance(raw, float) and raw == int(raw):
        raw = int(raw)
    if isinstance(raw, int):
        s = str(raw)
    else:
        s = str(raw).strip().replace("\n", "")
    if s in ("全部订单", "折扣进展"):
        return None, "keep"
    pieces = [part.strip() for part in re.split(r"[；;]", s) if part.strip()]
    if len(pieces) > 1 and all(re.fullmatch(r"20\d{4}", part) for part in pieces):
        return [int(part) for part in pieces], "list"
    if re.fullmatch(r"20\d{6}", s):
        return [int(s[:6])], "yyyymm"
    if re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", s):
        return [int(s)], "yyyymm"

    info = check_ym(date)
    matched = re.fullmatch(r"(\d{2,4})年(\d{1,2})月，(\d{2,4})年(\d{1,2})月交付订单", s)
    if matched:
        return [
            yy(matched.group(1)) * 100 + int(matched.group(2)),
            yy(matched.group(3)) * 100 + int(matched.group(4)),
        ], "list"
    matched = re.fullmatch(r"(\d{2,4})年(\d{1,2})月-(\d{2,4})年(\d{1,2})月交付订单", s)
    if matched:
        return months_between(
            yy(matched.group(1)),
            int(matched.group(2)),
            yy(matched.group(3)),
            int(matched.group(4)),
        ), "explicit"
    matched = re.fullmatch(r"(\d{2,4})年(\d{1,2})、(\d{1,2})月交付订单", s)
    if matched:
        year = yy(matched.group(1))
        return [year * 100 + int(matched.group(2)), year * 100 + int(matched.group(3))], "list"
    matched = re.fullmatch(r"(\d{2,4})年(\d{1,2})-(\d{1,2})月交付订单", s)
    if matched:
        year = yy(matched.group(1))
        start, end = int(matched.group(2)), int(matched.group(3))
        if end < start:
            raise ValueError(f"year span goes backwards: {s}")
        return [year * 100 + month for month in range(start, end + 1)], "explicit"
    matched = re.fullmatch(r"(\d{2,4})年(\d{1,2})月交付订单", s)
    if matched:
        return [yy(matched.group(1)) * 100 + int(matched.group(2))], "explicit"
    if info is None:
        return None, "keep"
    cy, cm = info
    matched = re.fullmatch(r"(\d{1,2})、(\d{1,2})月(份)?交付订单", s)
    if matched:
        return [
            year_for_month(int(matched.group(i)), cy, cm) * 100 + int(matched.group(i))
            for i in (1, 2)
        ], "list"
    matched = re.fullmatch(r"(\d{1,2})月及(\d{1,2})月交付订单", s)
    if matched:
        return [
            year_for_month(int(matched.group(i)), cy, cm) * 100 + int(matched.group(i))
            for i in (1, 2)
        ], "list"
    matched = re.fullmatch(r"(\d{1,2})-(\d{1,2})月(交付订单|客服订单)", s)
    if matched:
        return span_without_year(int(matched.group(1)), int(matched.group(2)), cy, cm), "span"
    matched = re.fullmatch(r"(\d{1,2})月交付订单", s)
    if matched:
        month = int(matched.group(1))
        return [year_for_month(month, cy, cm) * 100 + month], "single"
    return None, "keep"


def clean_name(name):
    if name is None or str(name).strip() == "":
        return "", []
    raw = str(name).strip()
    sos = SO_RE.findall(raw)
    text = raw
    if sos:
        text = PAREN_SO_RE.sub("", text)
        text = SO_RE.sub("", text)
        text = re.sub(r"[（(]\s*[)）]", "", text)
        text = re.sub(r"\s{2,}", " ", text).strip("、，, ")
    text = text.replace("中华人民共和国公安部", "公安部").replace("中国公安部", "公安部")
    return text, sos


def is_umbrella(name: str) -> bool:
    """公安部按订单看。分局、公安局仍按客户加交付月。"""
    return name == "公安部" or (name.startswith("公安部") and "分局" not in name and "公安局" not in name)
