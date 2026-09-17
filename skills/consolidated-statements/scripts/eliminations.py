"""Evidence-driven eliminations; source company cells remain unchanged."""
from __future__ import annotations
import ast
import json
import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter
from engine import CATALOG, COMPANIES, KINDS
import workbook as writer

CENT = Decimal(".01")
ZERO = Decimal("0")
RATE = Decimal("1.06")
RULE_VERSION = "jgy-eliminations-20260917.1"

def money(value):
    if isinstance(value, bool) or value is None:
        raise ValueError("金额必须为明确的十进制值")
    n = Decimal(str(value))
    if not n.is_finite() or abs(n) >= Decimal("1e20") or n != n.quantize(CENT):
        raise ValueError("金额不是有效的两位小数")
    return n

def evaluator(book):
    memo, active = {}, set()
    refs = re.compile(r"\$[A-Z]+\$[1-9][0-9]*")
    def calculate(sheet, address):
        key = sheet, address.replace("$", "")
        if key in memo:
            return memo[key]
        if key in active:
            raise ValueError("循环引用")
        active.add(key)
        raw = book[sheet][key[1]].value
        def visit(node):
            if isinstance(node, ast.Expression):
                return visit(node.body)
            if isinstance(node, ast.Constant) and type(node.value) in (int, float):
                return Decimal(ast.get_source_segment(body, node))
            if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
                return (-1 if isinstance(node.op, ast.USub) else 1) * visit(node.operand)
            if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub)):
                left, right = visit(node.left), visit(node.right)
                return left + right if isinstance(node.op, ast.Add) else left - right
            raise ValueError("存在不支持的公式，停止交付")
        if isinstance(raw, str) and raw.startswith("="):
            body = refs.sub(lambda m: "(" + str(calculate(sheet, m[0])) + ")", raw[1:])
            result = visit(ast.parse(body, mode="eval"))
        else:
            result = ZERO if raw is None or raw == "" else Decimal(str(raw))
        active.remove(key)
        memo[key] = result
        return result
    return calculate

def verify_formulas(book, expected):
    calc = evaluator(book)
    for key, expected_amount in expected.items():
        if calc(*key) != expected_amount:
            raise ValueError("公式独立求值不一致")

def validate_evidence(data, period):
    if data.get("period") != period or data.get("ledger") != "HEAD":
        raise ValueError("抵销证据的期间或取数账套不符")
    if not data.get("collected_at") or not isinstance(data.get("companies"), dict):
        raise ValueError("缺取数时间或公司明细")
    if set(data["companies"]) - (set(COMPANIES) - {"HEAD"}):
        raise ValueError("抵销证据含未知公司或本部自身")
    for company, record in data["companies"].items():
        for kind, evidence in record.items():
            if kind not in {"bs", "is", "cf"}:
                raise ValueError("未知抵销类型")
            status = evidence.get("status")
            if status not in {"observed", "blank", "no_data", "failed"}:
                raise ValueError("缺少明确的取数状态")
            if not evidence.get("refs"):
                raise ValueError("缺少可追溯的取数证据引用")
            if status not in {"observed", "blank"}:
                continue
            if kind == "bs":
                if status == "observed":
                    money(evidence["ending_debit"])
            elif kind == "cf":
                if status != "observed":
                    raise ValueError("现金流需逐项确认零值，不能整块空白")
                for metric in ("month", "ytd"):
                    for direction in ("in", "out"):
                        money(evidence[metric + "_" + direction])
            else:
                if status != "observed" or not evidence.get("accounts_complete"):
                    raise ValueError("利润抵销需核查全部供应商科目")
                seen = set()
                for account in evidence["accounts"]:
                    code = account["account"]
                    if not code or code in seen:
                        raise ValueError("同公司科目的本年累计被重复计入")
                    seen.add(code)
                    for flag in ("has_receipt_month", "has_receipt_ytd"):
                        if type(account.get(flag)) is not bool:
                            raise ValueError("收票资格未核实")
                    if account["has_receipt_month"] and not account["has_receipt_ytd"]:
                        raise ValueError("本月收票资格与累计资格冲突")
                    monthly = money(account["receipt_month"])
                    money(account["credit_ytd"])
                    if monthly and not account["has_receipt_month"]:
                        raise ValueError("非收票记录不能计入本月抵销")
                    if not account.get("refs"):
                        raise ValueError("缺科目明细证据")
    return data

def amounts(data, companies, kind):
    values, missing = [], []
    for company in companies:
        row = data["companies"].get(company, {}).get(kind)
        if not row or row["status"] not in {"observed", "blank"}:
            missing.append(company)
        else:
            values.append((company, row))
    if missing:
        return None, missing
    if kind == "bs":
        return sum((money(r["ending_debit"]) if r["status"] == "observed" else ZERO
                    for _, r in values), ZERO), []
    if kind == "is":
        monthly, ytd = ZERO, ZERO
        for _, row in values:
            accounts = row["accounts"]
            monthly += (sum((money(a["receipt_month"]) for a in accounts), ZERO) / RATE).quantize(CENT, rounding=ROUND_HALF_UP)
            # Eligibility is company-level; the year-to-date column is counted once per account.
            if any(a["has_receipt_ytd"] for a in accounts):
                ytd += (sum((money(a["credit_ytd"]) for a in accounts), ZERO) / RATE).quantize(CENT, rounding=ROUND_HALF_UP)
        return (monthly, ytd), []
    totals = {key: sum((money(r[key]) for _, r in values), ZERO)
              for key in ("month_in", "month_out", "ytd_in", "ytd_out")}
    result = [(totals[metric + "_out"], totals[metric + "_out"]) for metric in ("month", "ytd")]
    return result, []

def apply_eliminations(path, reports, period, data, base_result):
    validate_evidence(data, period)
    book = openpyxl.load_workbook(path)
    audit, issues = [], []
    original_values = {(s.title, c.coordinate): c.value for s in book for row in s for c in row}
    order = {c: i for i, c in enumerate(COMPANIES)}
    groups = {}
    for kind in KINDS:
        for scope in ("合并", "母公司"):
            selected = sorted((r for r in reports if r.kind == kind and (scope == "合并" or COMPANIES[r.company][2])),
                              key=lambda r: order[r.company])
            if not selected:
                continue
            sheet = book[scope + KINDS[kind] + period]
            catalog = CATALOG["standard"][kind]
            width, delta = catalog["width"], int(kind == "is")
            by_id = {x["id"]: x for x in catalog["rows"] if x["id"]}
            start = len(selected) * width
            def address(item, metric, start=start, by_id=by_id, kind=kind, delta=delta):
                row = by_id[item]
                col = start + row["col"] + 3 + (1 - metric if kind == "bs" else metric)
                return "$" + get_column_letter(col) + "$" + str(row["row"] + delta)
            groups[(kind, scope)] = (sheet, address, selected, start)
            members = [r.company for r in selected if r.company != "HEAD"]
            if "HEAD" not in [r.company for r in selected]:
                issues.append({"scope": scope, "kind": kind, "reason": "未纳入本部来源，未执行基于本部账套的抵销"})
                continue
            if kind == "is" and scope == "母公司":
                continue
            result, missing = amounts(data, members, kind)
            if missing:
                issues.append({"scope": scope, "kind": kind, "reason": "抵销证据缺失或口径未确认", "missing": missing})
            else:
                deductions = ([(8, 0, result), (44, 0, result)] if kind == "bs" else
                              [(1, m, result[m]) for m in (0, 1)] + [(4, m, result[m]) for m in (0, 1)] if kind == "is" else
                              [(3, m, result[m][0]) for m in (0, 1)] + [(8, m, result[m][1]) for m in (0, 1)])
                for item, metric, amount in deductions:
                    cell = sheet[address(item, metric)]
                    before = cell.value
                    cell.value = before + "-(" + str(amount) + ")"
                    audit.append({"sheet": sheet.title, "cell": cell.coordinate, "item": item, "metric": metric,
                                  "deduction": str(amount), "members": members, "before_formula": before, "after_formula": cell.value})
            if kind == "bs" and scope == "合并":
                investment, capital = sheet[address(18, 0)], sheet[address(61, 0)]
                # Use original company references, never the investment cell after zeroing.
                original = investment.value
                investment.value = original + "-(" + original[1:] + ")"
                capital.value += "-(" + original[1:] + ")"
                audit.append({"sheet": sheet.title, "rule": "investment_against_capital", "investment_before_formula": original,
                              "investment_after_formula": investment.value, "capital_after_formula": capital.value})
    # Every summary subtotal must depend on its own eliminated detail rows.
    for (kind, scope), (sheet, address, selected, start) in groups.items():
        for item, components in writer.TOTAL_RULES[kind].items():
            for metric in (0, 1):
                cell = sheet[address(item, metric)]
                if cell.value is not None:
                    cell.value = "=" + "".join(("+" if i > 0 else "-") + address(abs(i), metric)
                                               for i in components).lstrip("+")
        for row in sheet:
            for cell in row:
                if cell.column <= start and cell.value != original_values[sheet.title, cell.coordinate]:
                    raise ValueError("公司来源区域被改变")
    calc = evaluator(book)
    cache = {(s.title, c.coordinate): calc(s.title, c.coordinate)
             for s in book for row in s for c in row if c.data_type == "f"}
    checks = []
    for (kind, scope), (sheet, address, selected, _) in groups.items():
        equations = [(34, [(72, 1)])] if kind == "bs" else [(33, [(35, 1), (34, -1)])] if kind == "cf" else []
        for metric in (0, 1):
            for lhs, rhs in equations:
                diff = calc(sheet.title, address(lhs, metric)) - sum((sign * calc(sheet.title, address(item, metric)) for item, sign in rhs), ZERO)
                checks.append({"sheet": sheet.title, "metric": metric, "difference": str(diff), "passed": abs(diff) < CENT})
    writer.apply_workpaper_format(book)
    writer.save_cached_workbook(book, path, cache)
    readback = openpyxl.load_workbook(path, data_only=True)
    try:
        for key, expected in cache.items():
            raw = readback[key[0]][key[1]].value
            if raw is None or abs(Decimal(str(raw)) - expected) >= Decimal(".005"):
                raise ValueError("抵销后公式缓存回读失败")
    finally:
        readback.close()
        book.close()
    output = dict(base_result)
    output["eliminations"] = audit
    output["elimination_issues"] = issues
    output["elimination_checks"] = checks
    output["elimination_complete"] = not issues and all(c["passed"] for c in checks)
    output["complete"] = bool(base_result["complete"] and output["elimination_complete"])
    output["formula_count"] = len(cache)
    output["rules_version"] = RULE_VERSION
    output["elimination_evidence"] = data
    output["compilation_notes"] = [r for r in output["compilation_notes"] if r[0] != "范围"]
    output["compilation_notes"].append(["范围", "总合并=本部+子公司+分公司；母公司合并=本部+分公司。母公司利润表不抵销。"])
    for item in output["values"]:
        item["value"] = str(calc(item["sheet"], item["cell"]))
    return output
