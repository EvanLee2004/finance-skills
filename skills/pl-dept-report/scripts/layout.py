#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from common import col_idx, col_letter, load_json


def load_layout() -> dict:
    return load_json("版式.json")


def load_books() -> dict:
    return load_json("账套清单.json")


def load_dept_map() -> dict:
    return load_json("部门名映射.json")


def load_export_aliases() -> dict:
    return load_json("引出列名.json")


def account_row_map(layout: dict) -> dict[str, int]:
    start = 3
    return {row["code"]: start + i for i, row in enumerate(layout["accounts"]) if row.get("code")}


def parent_code(code: str, code_set: set[str]) -> str | None:
    if len(code) >= 8:
        p6 = code[:6]
        if p6 in code_set:
            return p6
    if len(code) >= 6:
        p4 = code[:4]
        if p4 in code_set:
            return p4
    return None


def direct_children(layout: dict) -> dict[str, list[str]]:
    codes = [row["code"] for row in layout["accounts"] if row.get("code")]
    code_set = set(codes)
    kids: dict[str, list[str]] = {c: [] for c in codes}
    for code in codes:
        p = parent_code(code, code_set)
        if p:
            kids[p].append(code)
    return kids


def dept_columns(layout: dict) -> list[tuple[str, str]]:
    start = col_idx(layout["dept_start_col"])
    names = layout["departments"]
    return [(names[i], col_letter(start + i)) for i in range(len(names))]


def first_dept_col(layout: dict, excel_name: str) -> str | None:
    for name, letter in dept_columns(layout):
        if name == excel_name:
            return letter
    return None


def map_dept_name(archive_name: str, mapping: dict, layout: dict) -> str | None:
    raw = str(archive_name or "").strip()
    if not raw:
        return None
    table = mapping.get("金蝶档案名到Excel列") or {}
    if raw in table:
        return str(table[raw]).strip()
    if raw in set(layout["departments"]):
        return raw
    return None
