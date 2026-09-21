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


def norm_seller(s: str) -> str:
    """销方名清洗：去内部空白（票面常见「有 限公司」断字），括号统一全角。"""
    s = re.sub(r"\s+", "", (s or "").strip())
    return s.replace("(", "（").replace(")", "）")


# 版面惯例：购买方在左、销售方在右；「名称」二字不会出现在公司名里。
NAME_LABEL = re.compile(r"名称\s*[：:]")
ENTITY_SPLIT = re.compile(r"\s+")
HAN = re.compile(r"[\u4e00-\u9fff]")
NON_NAME_HINT = (
    "项目名称", "规格型号", "价税合计", "合计", "备注", "开户行", "地址", "电话", "账号",
    "发票号码", "开票日期", "税务", "监制", "下载",
)


def name_like(tok: str) -> bool:
    """像实体名：至少两个汉字、不含「名称」、数字占比不过半（挡日期与代码行）。"""
    if not tok or "名称" in tok:
        return False
    if len(HAN.findall(tok)) < 2:
        return False
    digits = sum(1 for ch in tok if ch.isdigit())
    return digits * 2 < len(tok)


# 销方名被排版换行截断时（名字太长折行），把下一行的续字接回来。
COMPLETE_TAIL = ("公司", "企业", "中心", "工作室", "店", "部", "社", "厂", "所", "行", "队", "组", "站", "室", "）", ")")
CONT_START = ("限", "责任", "公司", "中心", "工作室", "店", "部", "社", "厂", "室", "站")


def join_continuation(got: str, lines: list[str], i: int) -> str:
    if not got or got.endswith(COMPLETE_TAIL) or i + 1 >= len(lines):
        return got
    nxt = norm_seller(lines[i + 1])
    if not nxt or len(nxt) > 10:
        return got
    if any(nxt.startswith(p) for p in CONT_START):
        return got + nxt
    return got


def buyer_of(text: str) -> str:
    """购买方（我方）名称，供兜底与防呆。像公司名才认，否则视为未知。"""
    for ln in (text or "").splitlines():
        m = re.search(r"[购买]\s*名称\s*[：:]\s*(.+?)\s*销\s*名称", ln)
        if m:
            got = norm_seller(m.group(1))
            if name_like(got):
                return got
        m = re.search(r"[购买]\s*名称\s*[：:]\s*(\S.*)$", ln)
        if m:
            got = norm_seller(m.group(1))
            if name_like(got):
                return got
    return ""


def parse_seller(text: str) -> str:
    """销方（开票抬头）公司名。四级取值，取不到返回空串，由调用方判待确认。

    1 旧版式独立字段「销售方名称：」
    2 2026 版电子发票「购 名称：X 销 名称：Y」双栏同行
    3 同行出现两个「名称：」时取最后一个
    4 名称被排版挤到别行：4a 与买家名同行则切掉买家；4b 无标签的纯名称行取最右一段
    """
    lines = (text or "").splitlines()
    buyer = buyer_of(text)
    for i, ln in enumerate(lines):
        m = re.search(r"销售方名称\s*[：:]\s*(\S.*)", ln)
        if m and m.group(1).strip():
            return join_continuation(norm_seller(m.group(1)), lines, i)
    for i, ln in enumerate(lines):
        m = re.search(r"销\s*名称\s*[：:]\s*(.+)$", ln)
        if m and m.group(1).strip():
            return join_continuation(norm_seller(m.group(1)), lines, i)
    for ln in lines:
        if len(NAME_LABEL.findall(ln)) >= 2:
            tail = NAME_LABEL.split(ln)[-1].strip()
            if tail:
                return norm_seller(tail)
    for ln in lines:
        c = re.sub(r"\s+", "", ln)
        if buyer and buyer in c:
            rest = norm_seller(c.replace(buyer, "", 1))
            rest = re.sub(r"^(购买|销售)?方?名称[：:]?", "", rest)
            if len(rest) >= 4 and "名称" not in rest:
                return rest
    for ln in lines:
        if "：" in ln or ":" in ln or any(h in ln for h in NON_NAME_HINT):
            continue
        parts = [norm_seller(p) for p in ENTITY_SPLIT.split(ln.strip()) if name_like(norm_seller(p))]
        if len(parts) >= 2:
            return parts[-1]          # 左买右销，取最右
    return ""


def parse_invoice_text(text: str) -> dict:
    raw = text or ""
    packed = compact(raw)
    kind = ""
    if "专用发票" in packed:
        kind = "专票"
    elif "普通发票" in packed:
        kind = "普票"
    seller = parse_seller(raw)
    buyer = buyer_of(raw)
    if seller and (
        seller == buyer
        or (buyer and buyer in seller)
        or "名称" in seller
        or not name_like(seller)
    ):
        seller = ""                   # 抽成买家、残串或不像实体名 → 判空，交给调用方 hold
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
