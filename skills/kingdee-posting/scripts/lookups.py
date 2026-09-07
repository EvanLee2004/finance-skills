#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""预处理查找：销项抄客户核算项目余额表；收款仍用 1131xx 往来。合成测试注入，不访问网络。"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
TWOPLACES = Decimal("0.01")

HOLD_ASSIST_MISSING = "客户核算项目余额表没有此抬头，请斯佳确认是否新建"
HOLD_ASSIST_MULTI = "客户核算项目余额表有多条应收，请斯佳确认记哪条"


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
class AssistRow:
    period: str
    customer_code: str
    customer_name: str
    account: str
    ending_debit: Decimal | None = None
    ending_credit: Decimal | None = None
    ytd_debit: Decimal | None = None
    ytd_credit: Decimal | None = None


def prev_completed_month(day) -> str:
    """开票月的上一已过完公历月，YYYYMM。"""
    s = period_month(day)
    if len(s) < 7:
        return ""
    year = int(s[:4])
    month = int(s[5:7]) - 1
    if month <= 0:
        year -= 1
        month = 12
    return f"{year}{month:02d}"


def _period_key(raw) -> str:
    s = str(raw or "").strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits[:6] if len(digits) >= 6 else ""


def _has_amt(debit, credit) -> bool:
    for val in (debit, credit):
        if val is not None and val != 0:
            return True
    return False


def assist_records(rows: list[AssistRow]) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    for row in rows or []:
        item = (str(row.customer_code or "").strip(), str(row.customer_name or "").strip())
        if item[0] and item[1] and item not in out:
            out.append(item)
    return out


def select_period_rows(rows: list[AssistRow], invoice_day: str) -> list[AssistRow]:
    rows = [r for r in (rows or []) if r.account]
    if not rows:
        return []
    periods = sorted({r.period for r in rows if r.period})
    if len(periods) <= 1:
        return list(rows)
    target = prev_completed_month(invoice_day)
    if not target:
        return []
    exact = [r for r in rows if r.period == target]
    if exact:
        return exact
    older = [p for p in periods if p <= target]
    if not older:
        return []
    best = max(older)
    return [r for r in rows if r.period == best]


def pick_assist_account(customer_code: str, invoice_day: str, rows: list[AssistRow]) -> tuple[str, str, str]:
    """返回 (应收, 收入, 失败原因)。销项科目只抄客户核算项目余额表。"""
    code = str(customer_code or "").strip()
    if not code:
        return "", "", HOLD_ASSIST_MISSING
    mine = [r for r in (rows or []) if str(r.customer_code or "").strip() == code and str(r.account or "").startswith("1131")]
    selected = select_period_rows(mine, invoice_day)
    by_acc: dict[str, list[AssistRow]] = {}
    for row in selected:
        acc = str(row.account or "").strip()
        if acc:
            by_acc.setdefault(acc, []).append(row)
    accounts = list(by_acc)
    if not accounts:
        return "", "", HOLD_ASSIST_MISSING
    if len(accounts) == 1:
        ar = accounts[0]
        rev = income_of(ar)
        if not rev:
            return "", "", "收入科目无法从应收推导"
        return ar, rev, ""
    with_end = [acc for acc in accounts if any(_has_amt(r.ending_debit, r.ending_credit) for r in by_acc[acc])]
    if len(with_end) >= 2:
        return "", "", HOLD_ASSIST_MULTI
    if len(with_end) == 1:
        ar = with_end[0]
        rev = income_of(ar)
        if not rev:
            return "", "", "收入科目无法从应收推导"
        return ar, rev, ""
    with_ytd = [acc for acc in accounts if any(_has_amt(r.ytd_debit, r.ytd_credit) for r in by_acc[acc])]
    if len(with_ytd) == 1:
        ar = with_ytd[0]
        rev = income_of(ar)
        if not rev:
            return "", "", "收入科目无法从应收推导"
        return ar, rev, ""
    return "", "", HOLD_ASSIST_MULTI


def assist_rows_from_dicts(items) -> list[AssistRow]:
    out: list[AssistRow] = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        acc = str(item.get("account") or "").strip()
        code = str(item.get("customer_code") or item.get("code") or "").strip()
        if not acc or not code:
            continue
        out.append(
            AssistRow(
                period=_period_key(item.get("period")),
                customer_code=code,
                customer_name=str(item.get("customer_name") or item.get("name") or "").strip(),
                account=acc,
                ending_debit=money(item.get("ending_debit") if "ending_debit" in item else item.get("balance")),
                ending_credit=money(item.get("ending_credit")),
                ytd_debit=money(item.get("ytd_debit")),
                ytd_credit=money(item.get("ytd_credit")),
            )
        )
    return out


@dataclass
class LookupBox:
    customer_lines: dict[str, list[str]] = field(default_factory=dict)
    ar_accounts: list[str] = field(default_factory=list)
    ar_balance: dict[tuple[str, str], Decimal] = field(default_factory=dict)
    period_debit: dict[tuple[str, str, str], Decimal] = field(default_factory=dict)
    receipt_sales: dict[tuple[str, str, str], list[str]] = field(default_factory=dict)
    order_sales: dict[str, list[str]] = field(default_factory=dict)
    applicant_dept: dict[str, str] = field(default_factory=dict)
    line_accounts: dict[str, str] = field(default_factory=dict)
    assist_rows: list[AssistRow] = field(default_factory=list)
    assist_supplied: bool = False

    def lines_for(self, customer: str) -> list[str]:
        n = norm_name(customer)
        if n in self.customer_lines:
            return list(self.customer_lines[n])
        peeled = n
        for suf in ("股份有限公司", "有限责任公司", "有限公司"):
            if peeled.endswith(suf) and len(peeled) > len(suf) + 1:
                peeled = peeled[: -len(suf)]
                break
        keys = [k for k in self.customer_lines if k == peeled or k.endswith(peeled) or peeled.endswith(k)]
        if len(keys) == 1:
            return list(self.customer_lines[keys[0]])
        return []

    def order_sales_for(self, customer: str) -> list[str]:
        return list(self.order_sales.get(norm_name(customer)) or [])

    def receipt_sales_for(self, customer: str, day: str, amount) -> list[str]:
        key = (norm_name(customer), str(day or "")[:10], amt_key(amount))
        return list(self.receipt_sales.get(key) or [])


def box_from_dict(raw: dict | None, ar_or_lines, applicant_dept: dict) -> LookupBox:
    raw = raw or {}
    if isinstance(ar_or_lines, dict):
        ar_accounts = []
        for v in ar_or_lines.values():
            acc = str(v or "").strip()
            if acc and acc not in ar_accounts:
                ar_accounts.append(acc)
    else:
        ar_accounts = [str(x).strip() for x in (ar_or_lines or []) if str(x).strip()]
    extra = raw.get("ar_accounts") or []
    for acc in extra:
        acc = str(acc).strip()
        if acc and acc not in ar_accounts:
            ar_accounts.append(acc)
    lines = {}
    for name, vals in (raw.get("customer_lines") or {}).items():
        cleaned = []
        for v in vals or []:
            s = str(v).strip()
            if s and s not in cleaned:
                cleaned.append(s)
        lines[norm_name(name)] = cleaned
    balances = {}
    for item in raw.get("ar_balance") or []:
        cus = str(item.get("customer_code") or "").strip()
        acc = str(item.get("account") or "").strip()
        val = money(item.get("balance"))
        if cus and acc and val is not None:
            balances[(cus, acc)] = val
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
    assist_supplied = "assist_rows" in raw
    assist_rows = assist_rows_from_dicts(raw.get("assist_rows") or [])
    return LookupBox(
        customer_lines=lines,
        ar_accounts=ar_accounts,
        ar_balance=balances,
        period_debit=debit,
        receipt_sales=rec,
        order_sales=orders,
        applicant_dept={k: str(v).strip() for k, v in applicant_dept.items() if not str(k).startswith("_") and v},
        line_accounts={},
        assist_rows=assist_rows,
        assist_supplied=assist_supplied,
    )


def _pick_by_period_debit(code: str, invoice_day: str, box: LookupBox, accounts: list[str]) -> tuple[str, str, str]:
    period = period_month(invoice_day)
    if not period:
        return "", "", "金蝶往来本期借方不可用"
    scored = []
    for acc in accounts:
        debit = box.period_debit.get((code, acc, period))
        if debit is not None and debit != 0:
            scored.append((debit, acc))
    if not scored:
        return "", "", "金蝶往来没有此客户应收"
    scored.sort(key=lambda x: x[0], reverse=True)
    if len(scored) > 1 and scored[0][0] == scored[1][0]:
        return "", "", "金蝶往来本期借方不唯一"
    ar = scored[0][1]
    rev = income_of(ar)
    if not rev:
        return "", "", "收入科目无法从应收推导"
    return ar, rev, ""


def resolve_ar(customer_code: str, invoice_day: str, box: LookupBox) -> tuple[str, str, str]:
    """返回 (应收, 收入, 失败原因)。科目只看金蝶该客户 1131xx 往来。"""
    code = str(customer_code or "").strip()
    if not code:
        return "", "", "金蝶往来没有此客户应收"
    accounts = [str(a).strip() for a in (box.ar_accounts or []) if str(a).strip()]
    nonzero = []
    for acc in accounts:
        bal = box.ar_balance.get((code, acc))
        if bal is not None and bal != 0:
            nonzero.append(acc)
    if len(nonzero) == 1:
        ar = nonzero[0]
        rev = income_of(ar)
        if not rev:
            return "", "", "收入科目无法从应收推导"
        return ar, rev, ""
    if not nonzero:
        # 当月借方有、余额被回款冲平：仍用本期借方判断科目。
        return _pick_by_period_debit(code, invoice_day, box, accounts)
    scored = []
    period = period_month(invoice_day)
    if not period:
        return "", "", "金蝶往来本期借方不可用"
    for acc in nonzero:
        debit = box.period_debit.get((code, acc, period))
        if debit is None:
            return "", "", "金蝶往来本期借方不可用"
        scored.append((debit, acc))
    scored.sort(key=lambda x: x[0], reverse=True)
    if len(scored) > 1 and scored[0][0] == scored[1][0]:
        return "", "", "金蝶往来本期借方不唯一"
    ar = scored[0][1]
    return ar, income_of(ar), ""


def resolve_ar_any(customers, invoice_day: str, customer_code: str, box: LookupBox) -> tuple[str, str, str]:
    return resolve_ar(customer_code, invoice_day, box)


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


def resolve_sales_any(customers, day: str, amount, box: LookupBox) -> tuple[str, str]:
    last = "找不到销售"
    seen = []
    for name in customers:
        n = norm_name(name)
        if not n or n in seen:
            continue
        seen.append(n)
        sales, why = resolve_sales(name, day, amount, box)
        if sales:
            return sales, ""
        last = why or last
    return "", last


def applicant_dept_code(name: str, box: LookupBox) -> str:
    return box.applicant_dept.get(str(name or "").strip()) or ""
