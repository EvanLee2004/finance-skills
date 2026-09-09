#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""预处理查找：销项/收款科目只抄客户核算项目余额表。合成测试注入，不访问网络。"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
TWOPLACES = Decimal("0.01")
_PAY_TAIL = re.compile(r"第[一二三四五六七八九十百千零〇两0-9]+(?:收缴户|账户|收款户|专户)$")

HOLD_ASSIST_MISSING = "客户核算项目余额表没有此抬头，请斯佳确认是否新建"
HOLD_ASSIST_MULTI = "客户核算项目余额表有多条应收，请斯佳确认记哪条"
HOLD_NEW_NO_LINE = "新建客户没有申请人科目，请斯佳确认记哪条"


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


SUFFIXES = ("股份有限公司", "有限责任公司", "有限公司")


def norm_name(name: str) -> str:
    s = (name or "").strip().replace("(", "（").replace(")", "）")
    return "".join(s.split())


def peel_name(name: str) -> str:
    n = norm_name(name)
    for suf in SUFFIXES:
        if n.endswith(suf) and len(n) > len(suf) + 1:
            return n[: -len(suf)]
    return n


def peel_pay_tail(name: str) -> str:
    """银行抬头常带「第二收缴户」；智云开票客户没有这段。"""
    n = peel_name(name)
    stripped = _PAY_TAIL.sub("", n)
    return stripped if stripped else n


def name_variants(name: str) -> list[str]:
    out: list[str] = []
    for raw in (norm_name(name), peel_name(name), peel_pay_tail(name)):
        if raw and raw not in out:
            out.append(raw)
    return out


def names_overlap(left: str, right: str) -> bool:
    return bool(set(name_variants(left)) & set(name_variants(right)))


def matching_name_keys(query: str, keys) -> list[str]:
    """先变体相交，再唯一包含。多个包含键都留下，销售是否唯一交给调用方。"""
    keys = [str(k) for k in (keys or []) if str(k).strip()]
    exact = [k for k in keys if names_overlap(query, k)]
    if exact:
        return exact
    qn = peel_pay_tail(query) or peel_name(query) or norm_name(query)
    if len(qn) < 4:
        return []
    contain = []
    for k in keys:
        kn = peel_pay_tail(k) or peel_name(k) or norm_name(k)
        if qn in kn or kn in qn:
            contain.append(k)
    return contain


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


def group_assist_accounts(customer_code: str, invoice_day: str, rows: list[AssistRow]) -> dict[str, list[AssistRow]]:
    code = str(customer_code or "").strip()
    if not code:
        return {}
    mine = [r for r in (rows or []) if str(r.customer_code or "").strip() == code and str(r.account or "").startswith("1131")]
    selected = select_period_rows(mine, invoice_day)
    by_acc: dict[str, list[AssistRow]] = {}
    for row in selected:
        acc = str(row.account or "").strip()
        if acc:
            by_acc.setdefault(acc, []).append(row)
    return by_acc


def list_assist_accounts(customer_code: str, invoice_day: str, rows: list[AssistRow]) -> list[str]:
    return list(group_assist_accounts(customer_code, invoice_day, rows))


def preferred_ar_code(raw: str) -> str:
    """07 / 113107 / 7 → 113107。空或不像应收则空串。"""
    s = str(raw or "").strip()
    if not s:
        return ""
    digits = "".join(ch for ch in s if ch.isdigit())
    if s.startswith("1131") and len(digits) >= 6:
        return "1131" + digits[4:6]
    if len(digits) >= 2:
        return "1131" + digits[-2:]
    if len(digits) == 1:
        return "1131" + digits.zfill(2)
    return ""


def pick_assist_account(
    customer_code: str,
    invoice_day: str,
    rows: list[AssistRow],
    preferred_ar: str = "",
) -> tuple[str, str, str]:
    """返回 (应收, 收入, 失败原因)。销项科目只抄客户核算项目余额表。

    一条 1131 仍抄表。多条时若 preferred_ar 落在该客户已有科目上则用它（申请人/销售拆业务线），
    否则仍待确认，禁止取最大。档案有、余额表没有时，第一笔可用申请人科目。
    """
    code = str(customer_code or "").strip()
    if not code:
        return "", "", HOLD_ASSIST_MISSING
    by_acc = group_assist_accounts(code, invoice_day, rows)
    accounts = list(by_acc)
    wanted = preferred_ar_code(preferred_ar)
    if not accounts:
        if wanted:
            rev = income_of(wanted)
            if not rev:
                return "", "", "收入科目无法从应收推导"
            return wanted, rev, ""
        return "", "", HOLD_NEW_NO_LINE
    if len(accounts) == 1:
        ar = accounts[0]
        rev = income_of(ar)
        if not rev:
            return "", "", "收入科目无法从应收推导"
        return ar, rev, ""
    if wanted and wanted in by_acc:
        rev = income_of(wanted)
        if not rev:
            return "", "", "收入科目无法从应收推导"
        return wanted, rev, ""
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
        found: list[str] = []
        for cname in matching_name_keys(customer, (self.order_sales or {}).keys()):
            for person in (self.order_sales or {}).get(cname) or []:
                if person and person not in found:
                    found.append(person)
        return found

    def receipt_sales_for(self, customer: str, day: str, amount) -> list[str]:
        day_s = str(day or "")[:10]
        amt = amt_key(amount)
        named: list[str] = []
        by_money: list[str] = []
        named_row = False
        for (cname, rec_day, rec_amt), sales in (self.receipt_sales or {}).items():
            if rec_day != day_s or rec_amt != amt:
                continue
            for person in sales or []:
                if person and person not in by_money:
                    by_money.append(person)
            if matching_name_keys(customer, [cname]):
                named_row = True
                for person in sales or []:
                    if person and person not in named:
                        named.append(person)
        if named_row:
            return named
        return by_money

    def receipt_history_sales_for(self, customer: str) -> list[str]:
        """这笔到账还没登时：同一开票客户历史回款仍唯一的销售。"""
        found: list[str] = []
        for (cname, _day, _amt), sales in (self.receipt_sales or {}).items():
            if not matching_name_keys(customer, [cname]):
                continue
            for person in sales or []:
                if person and person not in found:
                    found.append(person)
        return found


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
    sales, why, _cands = resolve_sales_detail(customer, day, amount, box)
    return sales, why


def resolve_sales_detail(customer: str, day: str, amount, box: LookupBox) -> tuple[str, str, list[str]]:
    rec = box.receipt_sales_for(customer, day, amount)
    if len(rec) == 1:
        return rec[0], "", rec
    if len(rec) > 1:
        return "", "回款销售不唯一", rec
    orders = box.order_sales_for(customer)
    if len(orders) == 1:
        return orders[0], "", orders
    if len(orders) > 1:
        return "", "下单对上两个销售，请斯佳单独处理", orders
    hist = box.receipt_history_sales_for(customer)
    if len(hist) == 1:
        return hist[0], "", hist
    if len(hist) > 1:
        return "", "回款销售不唯一", hist
    return "", "找不到销售", []


def resolve_sales_any(customers, day: str, amount, box: LookupBox) -> tuple[str, str]:
    sales, why, _cands = resolve_sales_any_detail(customers, day, amount, box)
    return sales, why


def resolve_sales_any_detail(customers, day: str, amount, box: LookupBox) -> tuple[str, str, list[str]]:
    last = "找不到销售"
    last_cands: list[str] = []
    seen = []
    for name in customers:
        n = norm_name(name)
        if not n or n in seen:
            continue
        seen.append(n)
        sales, why, cands = resolve_sales_detail(name, day, amount, box)
        if sales:
            return sales, "", cands
        last = why or last
        if cands:
            last_cands = cands
    return "", last, last_cands


def applicant_dept_code(name: str, box: LookupBox) -> str:
    return box.applicant_dept.get(str(name or "").strip()) or ""


def rows_from_balance_box(box) -> list[AssistRow]:
    existing = list(getattr(box, "assist_rows", None) or [])
    if existing:
        return existing
    out: list[AssistRow] = []
    for item, val in (getattr(box, "ar_balance", None) or {}).items():
        if not isinstance(item, tuple) or len(item) != 2:
            continue
        cus, acc = str(item[0] or "").strip(), str(item[1] or "").strip()
        if not cus or not acc:
            continue
        out.append(
            AssistRow(
                period="",
                customer_code=cus,
                customer_name="",
                account=acc,
                ending_debit=val if isinstance(val, Decimal) else money(val),
            )
        )
    return out
