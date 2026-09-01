#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Evaluate written Excel formula text. Not a second copy of business checks."""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

_CELL = re.compile(r"^(?:'?(?P<sheet>[^'!]+)'?!)?\$?(?P<col>[A-Za-z]+)\$?(?P<row>\d+)$")


class FormulaError(ValueError):
    pass


def col_to_idx(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


class WorkbookEval:
    def __init__(self, wb):
        self.wb = wb
        self.cache: dict[tuple[str, int, int], object] = {}
        self.stack: set[tuple[str, int, int]] = set()

    def cell_value(self, sheet: str, row: int, col: int):
        key = (sheet, row, col)
        if key in self.cache:
            return self.cache[key]
        if key in self.stack:
            return None
        ws = self.wb[sheet]
        raw = ws.cell(row, col).value
        if raw is None or raw == "":
            self.cache[key] = None
            return None
        if isinstance(raw, (int, float, Decimal)):
            self.cache[key] = Decimal(str(raw))
            return self.cache[key]
        if isinstance(raw, str) and raw.startswith("="):
            self.stack.add(key)
            try:
                val = evaluate(raw[1:], self, sheet)
            finally:
                self.stack.discard(key)
            self.cache[key] = val
            return val
        self.cache[key] = raw
        return raw


def evaluate(expr: str, book: WorkbookEval, sheet: str):
    parser = Parser(expr, book, sheet)
    val = parser.parse_expr()
    parser.skip()
    if parser.i < len(parser.s):
        raise FormulaError(f"trailing {parser.s[parser.i:]}")
    return val


class Parser:
    def __init__(self, expr: str, book: WorkbookEval, sheet: str):
        self.s = expr.strip()
        self.i = 0
        self.book = book
        self.sheet = sheet

    def skip(self):
        while self.i < len(self.s) and self.s[self.i].isspace():
            self.i += 1

    def parse_expr(self):
        return self.parse_compare()

    def parse_compare(self):
        left = self.parse_add()
        self.skip()
        for op in ("<>", "<=", ">=", "=", "<", ">"):
            if self.s.startswith(op, self.i):
                self.i += len(op)
                right = self.parse_add()
                lv = _num(left)
                rv = _num(right)
                if op == "=":
                    return lv == rv
                if op == "<>":
                    return lv != rv
                if op == "<":
                    return lv < rv
                if op == ">":
                    return lv > rv
                if op == "<=":
                    return lv <= rv
                if op == ">=":
                    return lv >= rv
        return left

    def parse_add(self):
        val = self.parse_mul()
        while True:
            self.skip()
            if self.i < len(self.s) and self.s[self.i] in "+-":
                op = self.s[self.i]
                self.i += 1
                rhs = self.parse_mul()
                lv = _num(val)
                rv = _num(rhs)
                val = lv + rv if op == "+" else lv - rv
            else:
                return val

    def parse_mul(self):
        val = self.parse_unary()
        while True:
            self.skip()
            if self.i < len(self.s) and self.s[self.i] in "*/":
                op = self.s[self.i]
                self.i += 1
                rhs = self.parse_unary()
                lv = _num(val)
                rv = _num(rhs)
                if op == "*":
                    val = lv * rv
                else:
                    if rv == 0:
                        return None
                    val = lv / rv
            else:
                return val

    def parse_unary(self):
        self.skip()
        if self.i < len(self.s) and self.s[self.i] == "-":
            self.i += 1
            return -_num(self.parse_unary())
        if self.i < len(self.s) and self.s[self.i] == "+":
            self.i += 1
            return self.parse_unary()
        return self.parse_primary()

    def parse_primary(self):
        self.skip()
        if self.i >= len(self.s):
            raise FormulaError("empty")
        ch = self.s[self.i]
        if ch == "(":
            self.i += 1
            val = self.parse_expr()
            self.skip()
            if self.i >= len(self.s) or self.s[self.i] != ")":
                raise FormulaError("missing )")
            self.i += 1
            return val
        if ch in "\"'":
            return self.parse_string()
        if ch.isdigit() or (ch == "." and self.i + 1 < len(self.s) and self.s[self.i + 1].isdigit()):
            return self.parse_number()
        ident = self.read_ident_or_ref()
        self.skip()
        if ident and self.i < len(self.s) and self.s[self.i] == "(":
            self.i += 1
            args = self.parse_args()
            return self.call_fn(ident.upper(), args)
        if ident:
            return self.resolve_ref(ident)
        raise FormulaError(f"bad token at {self.s[self.i:self.i+20]}")

    def parse_args(self):
        args = []
        self.skip()
        if self.i < len(self.s) and self.s[self.i] == ")":
            self.i += 1
            return args
        while True:
            start = self.i
            depth = 0
            while self.i < len(self.s):
                c = self.s[self.i]
                if c == "(":
                    depth += 1
                elif c == ")":
                    if depth == 0:
                        break
                    depth -= 1
                elif c == "," and depth == 0:
                    break
                self.i += 1
            piece = self.s[start:self.i]
            if ":" in piece and "(" not in piece:
                args.append(("range", piece.strip()))
            else:
                args.append(Parser(piece, self.book, self.sheet).parse_expr())
            self.skip()
            if self.i < len(self.s) and self.s[self.i] == ",":
                self.i += 1
                continue
            if self.i < len(self.s) and self.s[self.i] == ")":
                self.i += 1
                return args
            raise FormulaError("bad args")

    def parse_string(self):
        quote = self.s[self.i]
        self.i += 1
        out = []
        while self.i < len(self.s):
            ch = self.s[self.i]
            if ch == quote:
                self.i += 1
                if self.i < len(self.s) and self.s[self.i] == quote:
                    out.append(quote)
                    self.i += 1
                    continue
                return "".join(out)
            out.append(ch)
            self.i += 1
        raise FormulaError("unterminated string")

    def parse_number(self):
        m = re.match(r"\d*\.?\d+(?:[eE][+-]?\d+)?", self.s[self.i:])
        if not m:
            raise FormulaError("bad number")
        self.i += m.end()
        return Decimal(m.group(0))

    def read_ident_or_ref(self):
        m = re.match(
            r"(?:'[^']+'|[^\s'!()+\-*/=,<>:]+)!\$?[A-Za-z]+\$?\d+|"
            r"\$?[A-Za-z]+\$?\d+|"
            r"[A-Za-z_\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]*",
            self.s[self.i:],
        )
        if not m:
            return ""
        self.i += m.end()
        return m.group(0)

    def resolve_ref(self, token: str):
        m = _CELL.match(token.replace("$", ""))
        if not m:
            if token.upper() in {"TRUE", "FALSE"}:
                return token.upper() == "TRUE"
            raise FormulaError(f"bad ref {token}")
        sheet = m.group("sheet") or self.sheet
        col = col_to_idx(m.group("col"))
        row = int(m.group("row"))
        return self.book.cell_value(sheet, row, col)

    def call_fn(self, name: str, args: list):
        if name == "SUM":
            total = Decimal("0")
            for arg in args:
                if isinstance(arg, tuple) and arg[0] == "range":
                    total += sum_range(arg[1], self.book, self.sheet)
                else:
                    total += _num(arg)
            return total
        if name == "IF":
            cond = args[0] if args else None
            a = args[1] if len(args) > 1 else None
            b = args[2] if len(args) > 2 else None
            truth = False
            if isinstance(cond, bool):
                truth = cond
            elif cond not in (None, "", 0, Decimal("0")):
                truth = True
            return a if truth else b
        if name == "N":
            return _num(args[0] if args else None)
        if name == "ABS":
            return abs(_num(args[0] if args else None))
        if name == "OR":
            return any(_truth(a) for a in args)
        if name == "AND":
            return all(_truth(a) for a in args)
        raise FormulaError(f"fn {name}")


def _truth(val) -> bool:
    if isinstance(val, bool):
        return val
    if val in (None, "", 0, Decimal("0")):
        return False
    return True


def _num(val) -> Decimal:
    if val is None or val == "":
        return Decimal("0")
    if isinstance(val, bool):
        return Decimal(1 if val else 0)
    if isinstance(val, Decimal):
        return val
    if isinstance(val, (int, float)):
        return Decimal(str(val))
    try:
        return Decimal(str(val))
    except (InvalidOperation, ValueError):
        return Decimal("0")


def sum_range(spec: str, book: WorkbookEval, sheet: str) -> Decimal:
    left, right = spec.split(":", 1)
    lm = _CELL.match(left.replace("$", "").strip())
    rm = _CELL.match(right.replace("$", "").strip())
    if not lm or not rm:
        raise FormulaError(f"range {spec}")
    sh = lm.group("sheet") or sheet
    r1, r2 = int(lm.group("row")), int(rm.group("row"))
    c1, c2 = col_to_idx(lm.group("col")), col_to_idx(rm.group("col"))
    if r1 > r2:
        r1, r2 = r2, r1
    if c1 > c2:
        c1, c2 = c2, c1
    total = Decimal("0")
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            total += _num(book.cell_value(sh, r, c))
    return total


def eval_workbook(wb):
    book = WorkbookEval(wb)
    values = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                raw = cell.value
                if isinstance(raw, str) and raw.startswith("="):
                    values[(sheet, cell.row, cell.column)] = book.cell_value(sheet, cell.row, cell.column)
    return book, values
