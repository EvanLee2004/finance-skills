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
    return dept_col_letter(layout, excel_name, 1)


def dept_col_letter(layout: dict, excel_name: str, occurrence: int = 1) -> str | None:
    n = 0
    for name, letter in dept_columns(layout):
        if name == excel_name:
            n += 1
            if n == occurrence:
                return letter
    return None


def _ben_gongsi_rule_matches(book: dict, code: str, name: str) -> bool:
    for prefix in book.get("never_map_prefixes") or []:
        if str(code).startswith(str(prefix)):
            return False
    for rule in book.get("rules") or []:
        prefixes = [str(p) for p in (rule.get("prefixes") or [])]
        needles = [str(n) for n in (rule.get("name_contains") or [])]
        if prefixes and not any(str(code).startswith(p) for p in prefixes):
            continue
        if needles and not any(n in name for n in needles):
            continue
        return True
    return False


def load_ben_gongsi_rules() -> dict:
    return load_json("本公司规则.json")


def apply_ben_gongsi(
    entity: str,
    archive_dept: str,
    account_code: str,
    account_name: str,
    rules: dict,
    layout: dict,
) -> tuple[bool, str | None]:
    """本公司按账套规则。返回 (已处理, 列字母)。已处理且列为 None = 进未映射，禁止猜。"""
    book = (rules.get("per_entity") or {}).get(entity) or {}
    if not book:
        return False, None
    source = str(book.get("source_dept") or "本公司")
    if str(archive_dept or "").strip() != source:
        return False, None
    code = str(account_code or "")
    name = str(account_name or "")
    for prefix in book.get("never_map_prefixes") or []:
        if code.startswith(str(prefix)):
            return True, None
    for rule in book.get("rules") or []:
        prefixes = [str(p) for p in (rule.get("prefixes") or [])]
        needles = [str(n) for n in (rule.get("name_contains") or [])]
        if prefixes and not any(code.startswith(p) for p in prefixes):
            continue
        if needles and not any(n in name for n in needles):
            continue
        if rule.get("action") == "unmapped":
            return True, None
        excel_name = str(rule.get("excel_dept") or "").strip()
        occ = int(rule.get("occurrence") or 1)
        return True, dept_col_letter(layout, excel_name, occ)
    return True, None


def accounts_as_bengongsi(
    entity_amts: dict,
    existing_dept_rows: list[dict],
    special_rules: dict,
    layout: dict,
) -> list[dict]:
    """核算项目没有这笔费用时，把科目余额当成该公司的本公司。已有本公司辅助的科目不重复。"""
    kids = direct_children(layout)
    names = layout_code_names(layout)
    have: set[tuple[str, str]] = set()
    any_assist: set[tuple[str, str]] = set()
    for row in existing_dept_rows or []:
        ent = str(row.get("entity") or "")
        book = (special_rules.get("per_entity") or {}).get(ent) or {}
        source = str(book.get("source_dept") or "本公司")
        code = str(row.get("code") or "").strip()
        if code:
            any_assist.add((ent, code))
        if str(row.get("dept") or "").strip() != source:
            continue
        if code:
            have.add((ent, code))
    extra: list[dict] = []
    for ent, book in (special_rules.get("per_entity") or {}).items():
        if not book.get("fill_from_accounts_when_assist_lacks_code"):
            continue
        source = str(book.get("source_dept") or "本公司")
        blocked = tuple(str(p) for p in (book.get("never_map_prefixes") or []))
        for code, pair in (entity_amts.get(ent) or {}).items():
            code = str(code or "").strip()
            if not code or not str(code).startswith("5"):
                continue
            if any(code.startswith(p) for p in blocked):
                continue
            if kids.get(code):
                continue
            if (ent, code) in have:
                continue
            if book.get("fill_skip_if_any_assist") and (ent, code) in any_assist:
                continue
            name = names.get(code) or ""
            if book.get("fill_only_matching_rules") or book.get("fill_skip_if_any_assist"):
                if not _ben_gongsi_rule_matches(book, code, name):
                    continue
            extra.append(
                {
                    "entity": ent,
                    "dept": source,
                    "code": code,
                    "name": names.get(code) or "",
                    "debit": (pair or {}).get("debit") if isinstance(pair, dict) else None,
                    "credit": (pair or {}).get("credit") if isinstance(pair, dict) else None,
                }
            )
    return extra


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
