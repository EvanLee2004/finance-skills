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


def load_name_synonyms() -> dict[str, str]:
    try:
        data = load_json("科目名同义.json")
    except FileNotFoundError:
        return {}
    table = data.get("source_to_layout") or {}
    return {_norm_acct_name(k): _norm_acct_name(v) for k, v in table.items() if k and v}


def _norm_acct_name(value: str) -> str:
    return str(value or "").replace("\xa0", "").replace(" ", "").replace("　", "").strip()


def canonical_account_name(name: str, synonyms: dict[str, str] | None = None) -> str:
    raw = _norm_acct_name(name)
    if not raw:
        return ""
    return (synonyms or {}).get(raw, raw)


def names_match(src: str, layout_name: str, synonyms: dict[str, str] | None = None) -> bool:
    if not _norm_acct_name(src):
        return True
    a = canonical_account_name(src, synonyms)
    b = canonical_account_name(layout_name, synonyms)
    if a == b:
        return True
    for part in str(src).replace("\xa0", "").split("_"):
        if canonical_account_name(part, synonyms) == b:
            return True
    if a and b and (a + "费" == b or b + "费" == a):
        return True
    return False


def account_nature(code: str) -> str:
    c = str(code or "")
    if c.startswith(("51", "52", "53")):
        return "income"
    if c.startswith(("54", "55", "56")):
        return "expense"
    return "other"


def layout_code_names(layout: dict) -> dict[str, str]:
    return {row["code"]: str(row.get("name") or "") for row in layout["accounts"] if row.get("code")}


def source_ancestor_in(code: str, source_codes: set[str]) -> str | None:
    if len(code) >= 8 and code[:6] in source_codes:
        return code[:6]
    if len(code) >= 6 and code[:4] in source_codes:
        return code[:4]
    return None


def find_layout_by_name(
    src_name: str,
    src_code: str,
    layout: dict,
    synonyms: dict[str, str] | None = None,
    prefer_prefix: str = "",
) -> str | None:
    if not _norm_acct_name(src_name):
        return None
    by_code = layout_code_names(layout)
    nature = account_nature(src_code)
    hits = [
        code
        for code, name in by_code.items()
        if account_nature(code) == nature and names_match(src_name, name, synonyms)
    ]
    if not hits:
        return None
    prefix = prefer_prefix or (src_code[:4] if len(src_code) >= 4 else src_code)
    ranked = [c for c in hits if prefix and c.startswith(prefix)] or list(hits)
    src_canon = canonical_account_name(src_name, synonyms)
    exact = [c for c in ranked if canonical_account_name(by_code[c], synonyms) == src_canon]
    if not exact:
        for part in src_name.split("_"):
            part_canon = canonical_account_name(part, synonyms)
            exact.extend(
                c for c in ranked if canonical_account_name(by_code[c], synonyms) == part_canon
            )
    pool = exact or ranked
    if len(src_code) == 4:
        four = [c for c in pool if len(c) == 4]
        if four:
            pool = four
        pool = sorted(set(pool), key=lambda c: (len(c), c))
    else:
        pool = sorted(set(pool), key=lambda c: (-len(c), c))
    if not pool:
        return None
    top = pool[0]
    same = [c for c in pool if len(c) == len(top)]
    if prefix:
        pref = [c for c in same if c.startswith(prefix)]
        if len(pref) == 1:
            return pref[0]
        if len(pref) > 1:
            return sorted(pref)[0]
    if len(same) == 1:
        return same[0]
    return None


def build_prefix_remap(
    source_accounts: dict,
    layout: dict,
    synonyms: dict[str, str] | None = None,
) -> dict[str, str]:
    remap: dict[str, str] = {}
    for code, pair in (source_accounts or {}).items():
        if len(str(code)) != 4:
            continue
        name = str((pair or {}).get("name") or "") if isinstance(pair, dict) else ""
        found = find_layout_by_name(name, str(code), layout, synonyms, prefer_prefix="")
        if found and found[:4] != str(code):
            remap[str(code)] = found[:4]
    return remap


def resolve_account_target(
    code: str,
    name: str,
    source_codes: set[str],
    layout: dict,
    synonyms: dict[str, str] | None = None,
    prefix_remap: dict[str, str] | None = None,
) -> tuple[str | None, str | None]:
    code = str(code or "").strip()
    name = str(name or "").strip()
    if not code:
        return None, None
    by_code = layout_code_names(layout)
    layout_codes = set(by_code)
    remap = prefix_remap or {}
    p4 = code[:4] if len(code) >= 4 else code
    prefer = remap.get(p4, p4)

    if code in layout_codes and names_match(name, by_code[code], synonyms):
        return code, None

    found = find_layout_by_name(name, code, layout, synonyms, prefer_prefix=prefer)
    if found:
        if found != code and found in source_codes and code.startswith(found):
            return None, None
        note = f"同码不同名={code}->{found}" if code in layout_codes and found != code else None
        return found, note

    if code in layout_codes:
        return None, f"名称不符={code}"

    if source_ancestor_in(code, source_codes):
        return None, None

    parent = prefer if prefer in layout_codes else parent_code(code, layout_codes)
    if parent:
        return parent, None
    if code.startswith("5"):
        return None, f"表外科目={code}"
    return None, None


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
