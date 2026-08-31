#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""预处理查找：业务线→科目、回款/下单销售。合成测试注入，不访问网络。"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Iterable

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


def norm_name(name: str) -> str:
    s = (name or "").strip().replace("(", "（").replace(")", "）")
    return "".join(s.split())


def period_month(day) -> str:
    s = str(day or "").strip().replace("/", "-").replace(".", "-")[:10]
    return s[:7] if len(s) >= 7 else ""


def income_of(ar: str) -> str:
    ar = str(ar or "").strip()
    if len(ar) < 2:
        return ""
    return "5101" + ar[-2:]


def amt_key(value) -> str:
    got = money(value)
    return str(got) if got is not None else ""


@dataclass
class LookupBox:
    customer_lines: dict[str, list[str]] = field(default_factory=dict)
    period_debit: dict[tuple[str, str, str], Decimal] = field(default_factory=dict)
    receipt_sales: dict[tuple[str, str, str], list[str]] = field(default_factory=dict)
    order_sales: dict[str, list[str]] = field(default_factory=dict)
    applicant_dept: dict[str, str] = field(default_factory=dict)
    line_accounts: dict[str, str] = field(default_factory=dict)

    def lines_for(self, customer: str) -> list[str]:
        return list(self.customer_lines.get(norm_name(customer)) or [])

    def order_sales_for(self, customer: str) -> list[str]:
        return list(self.order_sales.get(norm_name(customer)) or [])

    def receipt_sales_for(self, customer: str, day: str, amount) -> list[str]:
        key = (norm_name(customer), str(day or "")[:10], amt_key(amount))
        return list(self.receipt_sales.get(key) or [])


def box_from_dict(raw: dict | None, line_accounts: dict, applicant_dept: dict) -> LookupBox:
    raw = raw or {}
    lines = {}
    for name, vals in (raw.get("customer_lines") or {}).items():
        cleaned = []
        for v in vals or []:
            s = str(v).strip()
            if s and s not in cleaned:
                cleaned.append(s)
        lines[norm_name(name)] = cleaned
    debit = {}
    for item in raw.get("period_debit") or []:
        cus = str(item.get("customer_code") or "").strip()
        acc = str(item.get("account") or "").strip()
        period = str(item.get("period") or "").strip()
        val = money(item.get("debit"))
        if cus and acc and period and val is not None:
            debit[(cus, acc, period)] = val
    rec = {}
    for item in raw.get("receipt_sales") or []:
        name = norm_name(str(item.get("customer") or ""))
        day = str(item.get("date") or "")[:10]
        key_amt = amt_key(item.get("amount"))
        sales = [str(s).strip() for s in (item.get("sales") or []) if str(s).strip()]
        if name and day and key_amt:
            existing = rec.setdefault((name, day, key_amt), [])
            for s in sales:
                if s not in existing:
                    existing.append(s)
    orders = {}
    for name, vals in (raw.get("order_sales") or {}).items():
        sales = []
        for v in vals or []:
            s = str(v).strip()
            if s and s not in sales:
                sales.append(s)
        orders[norm_name(name)] = sales
    return LookupBox(
        customer_lines=lines,
        period_debit=debit,
        receipt_sales=rec,
        order_sales=orders,
        applicant_dept={k: str(v).strip() for k, v in applicant_dept.items() if not str(k).startswith("_") and v},
        line_accounts={k: str(v).strip() for k, v in line_accounts.items() if not str(k).startswith("_") and v},
    )


def pick_line_account(lines: Iterable[str], mapping: dict[str, str]) -> tuple[str | None, str | None]:
    accs = []
    for line in lines:
        acc = mapping.get(line)
        if not acc:
            return None, "业务线无科目对照"
        if acc not in accs:
            accs.append(acc)
    if not accs:
        return None, "客户无业务线"
    if len(accs) == 1:
        return accs[0], None
    return None, "multi"


def resolve_ar(customer: str, invoice_day: str, customer_code: str, box: LookupBox) -> tuple[str, str, str]:
    """返回 (应收, 收入, 失败原因)。失败时前两空。"""
    lines = box.lines_for(customer)
    if not lines:
        return "", "", "客户无业务线"
    ar, why = pick_line_account(lines, box.line_accounts)
    if ar:
        rev = income_of(ar)
        if not rev:
            return "", "", "收入科目无法从应收推导"
        return ar, rev, ""
    if why != "multi":
        return "", "", why or "客户无业务线"
    period = period_month(invoice_day)
    if not period or not customer_code:
        return "", "", "多业务线且本期借方不可用"
    scored = []
    for line in lines:
        acc = box.line_accounts.get(line)
        if not acc:
            return "", "", "业务线无科目对照"
        debit = box.period_debit.get((customer_code, acc, period))
        if debit is None:
            return "", "", "多业务线且本期借方不可用"
        scored.append((debit, acc))
    scored.sort(key=lambda x: x[0], reverse=True)
    if not scored:
        return "", "", "多业务线且本期借方不可用"
    if len(scored) > 1 and scored[0][0] == scored[1][0]:
        return "", "", "多业务线本期借方不唯一"
    ar = scored[0][1]
    return ar, income_of(ar), ""


def resolve_sales(customer: str, day: str, amount, box: LookupBox) -> tuple[str, str]:
    """返回 (销售, 失败原因)。"""
    rec = box.receipt_sales_for(customer, day, amount)
    if len(rec) == 1:
        return rec[0], ""
    if len(rec) > 1:
        return "", "回款销售不唯一"
    orders = box.order_sales_for(customer)
    if len(orders) == 1:
        return orders[0], ""
    if len(orders) > 1:
        return "", "下单对上两个销售，请斯佳单独处理"
    return "", "找不到销售"


def applicant_dept_code(name: str, box: LookupBox) -> str:
    return box.applicant_dept.get(str(name or "").strip()) or ""
