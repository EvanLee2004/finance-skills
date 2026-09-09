#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""职工薪酬台账 → 损益表工资/社保/公积金叶子。金额只脚本加总，对话不回显。"""
from __future__ import annotations

import re
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from common import detect_period_text, load_json, money
from layout import dept_col_letter, direct_children, load_layout

CONFIG_NAME = "薪酬台账.json"
PAYROLL_NEEDLES = ("基本工资", "单位部分养老", "单位月缴存额", "组织架构1", "组织架构-1")


def load_payroll_rules() -> dict:
    return load_json(CONFIG_NAME)


def _text(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").replace("\n", "").strip()


def _norm(value) -> str:
    return _text(value).replace(" ", "").replace("　", "")


def looks_like_payroll(path: Path) -> bool:
    rules = load_payroll_rules()
    name = path.name
    if any(h in name for h in rules.get("file_name_hints") or []):
        return True
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    try:
        wb = load_workbook(path, read_only=True, data_only=False)
    except Exception:
        return False
    try:
        titles = " ".join(wb.sheetnames)
        if any(s in titles for s in ("组织架构", "甲骨易工资", "文化工资")):
            return True
    finally:
        wb.close()
    return False


def discover_payroll_files(homes: list[Path], period: str) -> list[Path]:
    found: list[Path] = []
    seen: set[str] = set()
    for folder in homes:
        if not folder:
            continue
        home = Path(folder)
        paths = [home] if home.is_file() else []
        if home.is_dir():
            paths = list(home.rglob("*.xlsx")) + list(home.rglob("*.xlsm"))
        for path in paths:
            path = Path(path)
            if not path.is_file() or path.name.startswith("~$"):
                continue
            if "损益类部门科目余额表" in path.name or path.name.startswith("月度损益表_"):
                continue
            key = str(path.resolve())
            if key in seen:
                continue
            if not looks_like_payroll(path):
                continue
            blob = path.name
            per = detect_period_text(blob)
            if per and per != period:
                continue
            seen.add(key)
            found.append(path)
    return found


def _skip_sheet(title: str, rules: dict) -> bool:
    raw = _text(title)
    for needle in rules.get("skip_sheet_contains") or []:
        if needle and needle in raw:
            return True
    return False


def _match_sheet_spec(title: str, rules: dict) -> dict | None:
    raw = _text(title)
    if any(n and n in raw for n in rules.get("skip_title_contains") or []):
        return None
    for spec in rules.get("sheet_entity") or []:
        hints = spec.get("contains") or []
        if any(h and h in raw for h in hints):
            return spec
    return None


def _header_map(ws, max_row: int = 5, max_col: int = 40) -> dict[str, tuple[int, int]]:
    found: dict[str, tuple[int, int]] = {}
    last_r = min(ws.max_row or 1, max_row)
    last_c = min(ws.max_column or 1, max_col)
    for r in range(1, last_r + 1):
        for c in range(1, last_c + 1):
            t = _norm(ws.cell(r, c).value)
            if t and t not in found:
                found[t] = (r, c)
    return found


def _find_header(headers: dict[str, tuple[int, int]], aliases: list[str], prefer_unit: bool = False) -> tuple[int, int] | None:
    ranked: list[tuple[int, tuple[int, int]]] = []
    for alias in aliases:
        key = _norm(alias)
        if key in headers:
            ranked.append((len(key), headers[key]))
            continue
        for raw, pos in headers.items():
            if key and key in raw:
                if prefer_unit and "个人" in raw and "单位" not in raw:
                    continue
                ranked.append((len(raw), pos))
    if not ranked:
        return None
    ranked.sort(reverse=True)
    return ranked[0][1]


def _cell_money(ws, r: int, c: int | None) -> Decimal | None:
    if not c:
        return None
    return money(ws.cell(r, c).value)


def _load_org_map(wb, rules: dict) -> dict[str, tuple[str, str]]:
    out: dict[str, tuple[str, str]] = {}
    for title in wb.sheetnames:
        if not any(h in title for h in rules.get("org_sheet_hints") or ["组织架构"]):
            continue
        ws = wb[title]
        headers = _header_map(ws, 3, 6)
        aliases = rules.get("header_aliases") or {}
        name_pos = _find_header(headers, aliases.get("name") or ["姓名"])
        org1_pos = _find_header(headers, aliases.get("org1") or ["组织架构1"])
        org2_pos = _find_header(headers, aliases.get("org2") or ["组织架构2"])
        if not name_pos:
            name_pos = (2, 1)
            org1_pos = org1_pos or (2, 2)
            org2_pos = org2_pos or (2, 3)
        start = max(p[0] for p in (name_pos, org1_pos or (2, 2), org2_pos or (2, 3))) + 1
        for r in range(start, (ws.max_row or start) + 1):
            name = _text(ws.cell(r, name_pos[1]).value)
            if not name or name in {"姓名", "合计", "总计"}:
                continue
            o1 = _text(ws.cell(r, (org1_pos or (0, 2))[1]).value)
            o2 = _text(ws.cell(r, (org2_pos or (0, 3))[1]).value)
            if o1.startswith("=") or o2.startswith("="):
                continue
            if name not in out and o1 and o1 != "组织架构-1":
                out[name] = (o1, o2)
    return out


def _override_dept(entity: str, *keys: str, rules: dict | None = None) -> tuple[str | None, int, str | None] | None:
    rules = rules or {}
    table = (rules.get("entity_dept_override") or {}).get(entity or "") or {}
    if not table:
        return None
    for key in keys:
        spec = table.get(_text(key)) if key else None
        if spec:
            return spec.get("excel_dept"), int(spec.get("occurrence") or 1), spec.get("prefix")
    spec = table.get("默认")
    if spec:
        return spec.get("excel_dept"), int(spec.get("occurrence") or 1), spec.get("prefix")
    return None


def resolve_dept(
    center: str,
    org2: str,
    rules: dict,
    entity: str = "",
    extra_key: str = "",
) -> tuple[str | None, int, str | None]:
    """返回 (excel_dept, occurrence, prefix覆盖)。对不上三元组都空。"""
    forced = _override_dept(entity, extra_key, org2, center, rules=rules)
    if forced:
        return forced
    fallbacks = set(rules.get("org2_fallback_to_center") or [])
    leaf = _text(org2)
    if leaf in fallbacks or (leaf and leaf.startswith("=")):
        leaf = ""
    table = rules.get("org2_to_dept") or {}
    spec = table.get(leaf) if leaf else None
    if spec:
        dept = spec.get("excel_dept")
        occ_map = spec.get("occurrence_by_center") or {}
        occ = int(occ_map.get(center) or spec.get("occurrence") or 1)
        prefix = spec.get("prefix")
        return dept, occ, prefix
    center_dept = (rules.get("center_to_dept") or {}).get(center)
    if center_dept:
        return center_dept, 1, None
    return None, 1, None


def item_code(item: str, prefix: str, rules: dict) -> str | None:
    table = rules.get("item_to_code") or {}
    row = table.get(item) or {}
    return row.get(prefix)


def _is_pay_name(name: str, rules: dict) -> bool:
    return any(n and n in (name or "") for n in rules.get("pay_needles") or [])


def _add_row(
    rows: list[dict],
    entity: str,
    code: str,
    name: str,
    dept: str,
    occurrence: int,
    amount: Decimal | None,
    fill_left: bool,
) -> None:
    if amount is None or amount == 0 or not code or not dept:
        return
    rows.append(
        {
            "entity": entity,
            "code": code,
            "name": name,
            "dept": dept,
            "occurrence": occurrence,
            "debit": amount,
            "credit": None,
            "fill_left": fill_left,
        }
    )


def _iter_people(ws, start: int, name_c: int | None):
    for r in range(start, (ws.max_row or start) + 1):
        label = _text(ws.cell(r, 1).value)
        if label in {"合计", "总计", "行标签"}:
            continue
        name = _text(ws.cell(r, name_c).value) if name_c else ""
        if name in {"姓名", "合计", "总计"}:
            continue
        yield r, name


def parse_wage_sheet(ws, spec: dict, org_map: dict, rules: dict, unmapped: list[str]) -> list[dict]:
    headers = _header_map(ws)
    aliases = rules.get("header_aliases") or {}
    wage_pos = _find_header(headers, aliases.get("wage") or ["基本工资"])
    if not wage_pos:
        return []
    name_pos = _find_header(headers, aliases.get("name") or ["姓名"])
    org1_pos = _find_header(headers, aliases.get("org1") or ["组织架构1"])
    org2_pos = _find_header(headers, aliases.get("org2") or ["组织架构2"])
    dept_pos = _find_header(headers, aliases.get("dept") or ["部门"])
    region_pos = _find_header(headers, aliases.get("region") or ["地区"])
    start = wage_pos[0] + 1
    entity = spec["entity"]
    fill_left = entity in (rules.get("fill_left_entities") or [])
    forced_dept = spec.get("dept")
    forced_prefix = spec.get("prefix")
    rows: list[dict] = []
    for r, name in _iter_people(ws, start, name_pos[1] if name_pos else None):
        amt = _cell_money(ws, r, wage_pos[1])
        if amt is None:
            continue
        o1 = _text(ws.cell(r, org1_pos[1]).value) if org1_pos else ""
        o2 = _text(ws.cell(r, org2_pos[1]).value) if org2_pos else ""
        if name and (not o1 or not o2) and name in org_map:
            o1 = o1 or org_map[name][0]
            o2 = o2 or org_map[name][1]
        dept_val = _text(ws.cell(r, dept_pos[1]).value) if dept_pos else ""
        if not o1:
            o1 = dept_val
        extra = _text(ws.cell(r, region_pos[1]).value) if region_pos else dept_val
        if forced_dept:
            dept, occ, prefix = forced_dept, 1, forced_prefix
        else:
            dept, occ, prefix = resolve_dept(o1, o2, rules, entity=entity, extra_key=extra)
            prefix = prefix or (rules.get("center_to_prefix") or {}).get(o1)
        prefix = forced_prefix or prefix
        if not dept or not prefix:
            tag = o2 or o1 or extra or name or f"R{r}"
            if tag not in unmapped:
                unmapped.append(tag)
            continue
        if dept in (rules.get("never_fill_depts") or []):
            continue
        code = item_code("工资", prefix, rules)
        _add_row(rows, entity, code or "", "工资", dept, occ, amt, fill_left)
    return rows


def _si_columns(headers: dict, aliases: dict) -> dict[str, tuple[int, int]]:
    out = {}
    pairs = [
        ("养老保险", aliases.get("si_pension") or ["单位部分养老"], True),
        ("医疗保险", aliases.get("si_medical") or ["单位部分医疗"], True),
        ("失业保险", aliases.get("si_unemp") or ["单位部分失业"], True),
        ("工伤保险", aliases.get("si_injury") or ["单位部分工伤"], True),
        ("生育保险", aliases.get("si_maternity") or ["单位部分生育"], True),
        ("住房公积金", aliases.get("hf_unit") or ["单位月缴存额"], False),
    ]
    for item, al, unit in pairs:
        pos = _find_header(headers, al, prefer_unit=unit)
        if pos:
            out[item] = pos
    return out


def parse_si_sheet(ws, spec: dict, org_map: dict, rules: dict, unmapped: list[str]) -> list[dict]:
    headers = _header_map(ws, 6, 30)
    aliases = rules.get("header_aliases") or {}
    cols = _si_columns(headers, aliases)
    if not cols:
        return []
    name_pos = _find_header(headers, aliases.get("name") or ["姓名"])
    org1_pos = _find_header(headers, aliases.get("org1") or ["组织架构1"])
    org2_pos = _find_header(headers, aliases.get("org2") or ["组织架构2"])
    dept_pos = _find_header(headers, aliases.get("dept") or ["部门"])
    start = max(p[0] for p in cols.values()) + 1
    entity = spec["entity"]
    fill_left = entity in (rules.get("fill_left_entities") or [])
    forced_dept = spec.get("dept")
    forced_prefix = spec.get("prefix")
    rows: list[dict] = []
    for r, name in _iter_people(ws, start, name_pos[1] if name_pos else None):
        o1 = _text(ws.cell(r, org1_pos[1]).value) if org1_pos else ""
        o2 = _text(ws.cell(r, org2_pos[1]).value) if org2_pos else ""
        if name and name in org_map:
            o1 = o1 or org_map[name][0]
            o2 = o2 or org_map[name][1]
        if not o1 and dept_pos:
            o1 = _text(ws.cell(r, dept_pos[1]).value)
        extra = _text(ws.cell(r, dept_pos[1]).value) if dept_pos and not o1 else o1
        if forced_dept:
            dept, occ, prefix = forced_dept, 1, forced_prefix
        else:
            dept, occ, prefix = resolve_dept(o1, o2, rules, entity=entity, extra_key=extra)
            prefix = prefix or (rules.get("center_to_prefix") or {}).get(o1)
        prefix = forced_prefix or prefix
        if not dept or not prefix:
            has_amt = any(_cell_money(ws, r, pos[1]) for pos in cols.values())
            if has_amt:
                tag = o2 or o1 or extra or name or f"R{r}"
                if tag not in unmapped:
                    unmapped.append(tag)
            continue
        if dept in (rules.get("never_fill_depts") or []):
            continue
        for item, pos in cols.items():
            amt = _cell_money(ws, r, pos[1])
            code = item_code(item, prefix, rules)
            _add_row(rows, entity, code or "", item, dept, occ, amt, fill_left)
    return rows


def parse_agency_si(ws, spec: dict, rules: dict) -> list[dict]:
    """济南分/济南子对账单：雇主栏按险种进成本树。"""
    entity = spec["entity"]
    dept = spec.get("dept") or ""
    prefix = spec.get("prefix") or "5401"
    headers = _header_map(ws, 4, 24)
    # 两行表头：H养老雇主、J失业雇主、L工伤、M医疗雇主、O公积金雇主
    mapping = {
        8: "养老保险",
        10: "失业保险",
        12: "工伤保险",
        13: "医疗保险",
        15: "住房公积金",
    }
    # 若表头能认，优先表头
    aliases = rules.get("header_aliases") or {}
    named = _si_columns(headers, aliases)
    rows: list[dict] = []
    start = 5
    for r in range(start, (ws.max_row or start) + 1):
        seq = _text(ws.cell(r, 1).value)
        if not seq or seq in {"合计", "总计"}:
            if seq in {"合计", "总计"}:
                break
            continue
        if not seq.isdigit() and seq not in {"1", "2", "3"}:
            if not _text(ws.cell(r, 2).value):
                continue
        if named:
            for item, pos in named.items():
                amt = _cell_money(ws, r, pos[1])
                code = item_code(item, prefix, rules)
                _add_row(rows, entity, code or "", item, dept, 1, amt, True)
            continue
        for col, item in mapping.items():
            amt = _cell_money(ws, r, col)
            code = item_code(item, prefix, rules)
            _add_row(rows, entity, code or "", item, dept, 1, amt, True)
    return rows


def _hunan_unit_cols(ws) -> dict[str, int]:
    """湖南五险一金：只取各险种下的「单位」列，跳过大病/个人/基数。"""
    pairs = [
        ("工伤保险", ("工伤",)),
        ("失业保险", ("失业",)),
        ("养老保险", ("养老",)),
        ("医疗保险", ("医疗",)),
        ("住房公积金", ("公积金",)),
    ]
    found: dict[str, int] = {}
    last_r = min(ws.max_row or 1, 4)
    last_c = min(ws.max_column or 1, 24)
    for item, aliases in pairs:
        pos = None
        for r in range(1, last_r + 1):
            for c in range(1, last_c + 1):
                t = _norm(ws.cell(r, c).value)
                if not t or "大病" in t:
                    continue
                if item == "医疗保险" and "生育" in t and "医疗" not in t:
                    continue
                if not any(a in t for a in aliases):
                    continue
                if item == "医疗保险" and "大病" in t:
                    continue
                unit_c = None
                for rr in range(r, min(r + 2, last_r) + 1):
                    for cc in range(c, min(c + 4, last_c) + 1):
                        tt = _norm(ws.cell(rr, cc).value)
                        if "单位" in tt and "小计" not in tt and "五险" not in tt:
                            unit_c = cc
                            break
                    if unit_c:
                        break
                pos = unit_c or c
                break
            if pos:
                break
        if pos:
            found[item] = pos
    return found


def _hunan_name_dept_cols(ws) -> tuple[int, int, int]:
    headers = _header_map(ws, 4, 20)
    name_pos = _find_header(headers, ["姓名"]) or (2, 3)
    dept_pos = _find_header(headers, ["部门"]) or (2, 2)
    start = max(name_pos[0], dept_pos[0]) + 2
    return start, name_pos[1], dept_pos[1]


def parse_hunan_si(
    ws,
    spec: dict,
    rules: dict,
    unmapped: list[str],
    name_prefix: dict[str, tuple[str, str]] | None = None,
) -> list[dict]:
    cols = _hunan_unit_cols(ws)
    if not cols:
        return []
    start, name_c, dept_c = _hunan_name_dept_cols(ws)
    entity = spec["entity"]
    fill_left = entity in (rules.get("fill_left_entities") or [])
    rows: list[dict] = []
    for r in range(start, (ws.max_row or start) + 1):
        name = _text(ws.cell(r, name_c).value)
        dept_val = _text(ws.cell(r, dept_c).value)
        label = _text(ws.cell(r, 1).value)
        if label in {"合计", "总计"} or name in {"合计", "总计"}:
            break
        if not name or name in {"姓名"}:
            continue
        if name_prefix and name in name_prefix:
            dept, prefix = name_prefix[name]
            occ = 1
        else:
            dept, occ, prefix = resolve_dept(dept_val, "", rules, entity=entity, extra_key=dept_val)
            prefix = prefix or (rules.get("center_to_prefix") or {}).get(dept_val)
        if not dept or not prefix:
            tag = dept_val or name or f"R{r}"
            if tag not in unmapped:
                unmapped.append(tag)
            continue
        if dept in (rules.get("never_fill_depts") or []):
            continue
        for item, col in cols.items():
            amt = _cell_money(ws, r, col)
            code = item_code(item, prefix, rules)
            _add_row(rows, entity, code or "", item, dept, occ, amt, fill_left)
    return rows


def parse_hunan_zgs(ws, spec: dict, rules: dict, unmapped: list[str]) -> list[dict]:
    """湖南子：上半五险一金 + 下半工资表。销售/成本靠地区。"""
    wage_header = None
    last_r = min(ws.max_row or 1, 30)
    last_c = min(ws.max_column or 1, 16)
    for r in range(1, last_r + 1):
        for c in range(1, last_c + 1):
            if _norm(ws.cell(r, c).value) == "基本工资":
                wage_header = r
                break
        if wage_header:
            break
    name_prefix: dict[str, tuple[str, str]] = {}
    wage_rows: list[dict] = []
    entity = spec["entity"]
    fill_left = entity in (rules.get("fill_left_entities") or [])
    if wage_header:
        col_of = {}
        for c in range(1, last_c + 1):
            t = _norm(ws.cell(wage_header, c).value)
            if t and t not in col_of:
                col_of[t] = c
        name_c = col_of.get("姓名")
        region_c = col_of.get("地区")
        wage_c = col_of.get("基本工资")
        for r in range(wage_header + 1, (ws.max_row or wage_header) + 1):
            name = _text(ws.cell(r, name_c).value) if name_c else ""
            if not name or name in {"姓名", "合计", "总计"}:
                continue
            region = _text(ws.cell(r, region_c).value) if region_c else ""
            dept, occ, prefix = resolve_dept("", "", rules, entity=entity, extra_key=region)
            if not dept or not prefix:
                if region not in unmapped:
                    unmapped.append(region or name)
                continue
            name_prefix[name] = (dept, prefix)
            amt = _cell_money(ws, r, wage_c)
            code = item_code("工资", prefix, rules)
            _add_row(wage_rows, entity, code or "", "工资", dept, occ, amt, fill_left)
    si_rows = parse_hunan_si(ws, spec, rules, unmapped, name_prefix=name_prefix)
    return si_rows + wage_rows


def parse_payroll_file(path: Path, period: str) -> dict:
    rules = load_payroll_rules()
    wb = load_workbook(path, data_only=True)
    try:
        org_map = _load_org_map(wb, rules)
        rows: list[dict] = []
        unmapped: list[str] = []
        notes: list[str] = []
        used: list[str] = []
        for title in wb.sheetnames:
            if _skip_sheet(title, rules):
                continue
            spec = _match_sheet_spec(title, rules)
            if not spec:
                continue
            per = detect_period_text(title + "\n" + path.name)
            if per and per != period:
                notes.append(f"薪酬sheet期间不符={title}:{per}")
                continue
            ws = wb[title]
            kind = spec.get("kind")
            if kind == "wage":
                chunk = parse_wage_sheet(ws, spec, org_map, rules, unmapped)
            elif kind == "agency_si":
                chunk = parse_agency_si(ws, spec, rules)
            elif kind == "hunan_si":
                chunk = parse_hunan_si(ws, spec, rules, unmapped)
            elif kind == "hunan_zgs":
                chunk = parse_hunan_zgs(ws, spec, rules, unmapped)
            else:
                chunk = parse_si_sheet(ws, spec, org_map, rules, unmapped)
            if chunk:
                used.append(title)
                rows.extend(chunk)
        return {
            "path": str(path),
            "rows": rows,
            "unmapped": unmapped,
            "notes": notes,
            "sheets": used,
            "entities": sorted({r["entity"] for r in rows}),
        }
    finally:
        wb.close()


def apply_payroll(
    entity_amts: dict,
    dept_amts: dict,
    parsed: dict,
    layout: dict,
    notes: list[str],
    report: dict,
) -> tuple[set[str], dict[str, set[str]]]:
    """写入薪酬叶子。左列已有金蝶/代账的不覆盖。返回 (主体, 各主体写过的科目)。"""
    rules = load_payroll_rules()
    kids = direct_children(layout)
    blocked = set(rules.get("never_fill_depts") or [])
    covered: set[str] = set()
    codes_by_ent: dict[str, set[str]] = {}
    pre_left: set[tuple[str, str]] = set()
    for ent, bucket in (entity_amts or {}).items():
        for code, pair in (bucket or {}).items():
            if (pair or {}).get("debit") is not None or (pair or {}).get("credit") is not None:
                pre_left.add((ent, str(code)))
    unmapped = list(parsed.get("unmapped") or [])
    for row in parsed.get("rows") or []:
        code = str(row.get("code") or "")
        dept = str(row.get("dept") or "")
        if not code or kids.get(code) or dept in blocked:
            continue
        letter = dept_col_letter(layout, dept, int(row.get("occurrence") or 1))
        if not letter:
            tag = dept
            if tag and tag not in unmapped:
                unmapped.append(tag)
            continue
        amt = row.get("debit")
        if amt is None:
            continue
        bucket = dept_amts.setdefault(code, {})
        bucket[letter] = (bucket.get(letter) or Decimal("0")) + amt
        ent = row.get("entity") or ""
        covered.add(ent)
        codes_by_ent.setdefault(ent, set()).add(code)
        if row.get("fill_left") and ent and (ent, code) not in pre_left:
            cell = entity_amts.setdefault(ent, {}).setdefault(code, {"debit": None, "credit": None})
            cell["debit"] = (cell.get("debit") or Decimal("0")) + amt
    if parsed.get("sheets"):
        notes.append("薪酬台账=" + ",".join(parsed["sheets"]))
    notes.extend(parsed.get("notes") or [])
    existing = list(report.get("unmapped_depts") or [])
    for name in unmapped:
        if name and name not in existing:
            existing.append(name)
    report["unmapped_depts"] = existing
    report["payroll_codes"] = {k: sorted(v) for k, v in codes_by_ent.items()}
    return covered, codes_by_ent


def is_payroll_account_name(name: str) -> bool:
    return _is_pay_name(name, load_payroll_rules())
