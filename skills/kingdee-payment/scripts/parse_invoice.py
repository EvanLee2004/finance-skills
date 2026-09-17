#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""电子发票 PDF：专/普、销售方、价税合计、税额。标题带空格也能认。"""
from __future__ import annotations

import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

TWOPLACES = Decimal("0.01")


def money(v):
    if v is None or v == "":
        return None
    if isinstance(v, Decimal):
        return v.quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return Decimal(v).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    if isinstance(v, float):
        return Decimal(str(round(v, 2))).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    s = str(v).strip().replace(",", "")
    if not s or s.startswith("="):
        return None
    try:
        return Decimal(s).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    except Exception:
        return None


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def parse_invoice_text(text: str) -> dict:
    raw = text or ""
    packed = compact(raw)
    kind = ""
    if "专用发票" in packed:
        kind = "专票"
    elif "普通发票" in packed:
        kind = "普票"
    seller = ""
    for key in ("销售方名称：", "销售方名称:", "销售方名称"):
        if key in packed:
            rest = packed.split(key, 1)[1]
            rest = re.split(r"统一社会|纳税人识别|购买方|项目名称", rest, maxsplit=1)[0]
            seller = rest.strip("：: ")
            break
    if not seller:
        for key in ("销售方名称：", "销售方名称:", "名称：", "名称:"):
            if key in raw:
                seller = raw.split(key, 1)[1].splitlines()[0].strip()
                if seller:
                    break
    total = None
    m = re.search(r"价税合计.*?小写[¥￥]([-0-9,]+\.\d{2})", packed)
    if m:
        total = money(m.group(1))
    if total is None:
        m = re.search(r"价税合计.*?[¥￥]([-0-9,]+\.\d{2})", packed)
        if m:
            total = money(m.group(1))
    tax = None
    m = re.search(r"合计[¥￥]([-0-9,]+\.\d{2})[¥￥]([-0-9,]+\.\d{2})", packed)
    if m:
        tax = money(m.group(2))
    if tax is None:
        m = re.search(r"(?:合计)?税额[:：]?[¥￥]?([-0-9,]+\.\d{2})", packed)
        if m:
            tax = money(m.group(1))
    return {"kind": kind, "seller": seller, "total": total, "tax": tax}


NO_PDF_READER_ASK = (
    "本机没有 pdfplumber，读不了发票 PDF，未生成引入表。"
    "请先装：pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pdfplumber，再重跑。"
)


def parse_invoice_pdf(path: Path) -> dict:
    empty = {"kind": "", "seller": "", "total": None, "tax": None}
    try:
        import pdfplumber
    except ImportError:
        return {**empty, "error": "no_pdfplumber"}
    try:
        with pdfplumber.open(str(path)) as pdf:
            text = "\n".join((page.extract_text() or "") for page in pdf.pages[:4])
    except Exception:
        return {**empty, "error": "unreadable"}
    if not compact(text):
        return {**empty, "error": "unreadable"}
    return parse_invoice_text(text)
