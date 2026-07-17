#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
部门费用归集分摊主脚本。

阶段能力（随样例材料到位迭代）：
  v0  读表结构模板 + 读各主体余额（本期借贷）→ 拼左半边+合计；部门列先空/仅整挂规则
  v1  + 人员归属 + 收入三源透视（等样例）
  v2  + 按人费用（用友/工资社保）填管理/研发等
  v3  + 与定稿逐格 diff 验收

当前：v0 骨架 + 合成/真实主体余额填充 + 整挂规则 + 核对公式列。
金额只由本脚本计算；核对不平不改数。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
CONFIG = SKILL / "config"

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except ImportError:
    print("需要 openpyxl: pip install openpyxl", file=sys.stderr)
    raise SystemExit(1)


def log(msg: str):
    print(msg, file=sys.stderr)


def load_structure():
    p = CONFIG / "表结构模板.json"
    return json.loads(p.read_text(encoding="utf-8"))


def load_dept_map():
    p = CONFIG / "部门名映射.json"
    if not p.is_file():
        return {}
    d = json.loads(p.read_text(encoding="utf-8"))
    return {k: v for k, v in d.items() if not k.startswith("_")}


def clean_code(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, float):
        if abs(v - int(v)) < 1e-9:
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if re.match(r"^\d+\.0+$", s):
        return s.split(".")[0]
    return s


def num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    if not s or s == "-":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def read_balance_xls(path: Path) -> dict[str, dict]:
    """读 U8 风格发生额及余额表 → {科目编码: {name, debit, credit}} 本期发生。"""
    import xlrd

    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)
    # find header row with 科目编码
    start = 0
    for r in range(min(10, sh.nrows)):
        row = [str(sh.cell_value(r, c)) for c in range(min(6, sh.ncols))]
        if any("科目编码" in x for x in row):
            start = r + 1
            # maybe next row is 借方贷方
            if start < sh.nrows and any("借" in str(sh.cell_value(start, c)) for c in range(min(8, sh.ncols))):
                start += 1
            break
    # columns: 0 code, 1 name, 4 debit period?, 5 credit — from sample:
    # R3: 科目编码 科目名称 期初余额 本期发生 期末余额
    # R4: 借方 贷方 借方 贷方 借方 贷方
    # so period debit=col4 credit=col5
    out = {}
    for r in range(start, sh.nrows):
        code = clean_code(sh.cell_value(r, 0))
        if not code or not re.match(r"^\d+", code):
            continue
        name = str(sh.cell_value(r, 1)).strip()
        # try period columns: often 4,5
        debit = credit = 0.0
        if sh.ncols >= 6:
            debit = num(sh.cell_value(r, 4))
            credit = num(sh.cell_value(r, 5))
        out[code] = {"name": name, "debit": debit, "credit": credit}
    return out


def read_balance_any(path: Path) -> dict[str, dict]:
    if path.suffix.lower() == ".xls":
        return read_balance_xls(path)
    # xlsx fallback similar layout
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    sh = wb[wb.sheetnames[0]]
    rows = list(sh.iter_rows(values_only=True))
    wb.close()
    start = 0
    for i, row in enumerate(rows[:10]):
        if row and any(x and "科目编码" in str(x) for x in row):
            start = i + 1
            if start < len(rows) and any(x and "借" in str(x) for x in (rows[start] or [])):
                start += 1
            break
    out = {}
    for row in rows[start:]:
        if not row:
            continue
        code = clean_code(row[0] if len(row) > 0 else "")
        if not code or not re.match(r"^\d+", code):
            continue
        name = str(row[1] or "").strip() if len(row) > 1 else ""
        debit = num(row[4]) if len(row) > 4 else 0.0
        credit = num(row[5]) if len(row) > 5 else 0.0
        out[code] = {"name": name, "debit": debit, "credit": credit}
    return out


def discover_entity_files(input_dir: Path) -> dict[str, Path]:
    """粗映射：文件名含 甲骨易/文化/上海 等。"""
    mapping = {}
    keys = [
        ("甲骨易", "甲骨易"),
        ("北京", "甲骨易"),
        ("文化", "文化"),
        ("上海", "上海"),
        ("山东", "山东分公司"),
        ("湖南", "湖南分公司"),
    ]
    for f in input_dir.rglob("*"):
        if not f.is_file() or f.name.startswith("~$"):
            continue
        if f.suffix.lower() not in {".xls", ".xlsx"}:
            continue
        for kw, ent in keys:
            if kw in f.name or kw in str(f.parent):
                mapping.setdefault(ent, f)
    return mapping


# 整挂规则（config 文字版的可执行子集；冲突项不在此硬写死 5401 整挂）
# 整挂：科目前缀匹配（长前缀优先）。5401 仅 540101/540103 整挂项目总监及助理，其余 5401* 按人。
DIRECT_HANG = {
    "540101": "项目总监及助理",  # 斯佳标黄 20260717
    "540103": "项目总监及助理",  # 翻译语言服务等
    "5402": "财务中心",
    "5504": "财务中心",
}


def prefix_rule(code: str) -> str | None:
    if not code:
        return None
    # 5401 大类本身不整挂；仅明确子目
    for pref, dept in sorted(DIRECT_HANG.items(), key=lambda x: -len(x[0])):
        if code == pref or code.startswith(pref):
            # 避免 5402 误匹配到别的；用 startswith 对 540101 安全
            if pref.startswith("5401") and not (code == pref or code.startswith(pref)):
                continue
            return dept
    return None


def build_workbook(structure: dict, entity_data: dict[str, dict[str, dict]], report: list[str]):
    """entity_data: 实体名 -> {code: {debit, credit, name}}"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "部门科目余额表"

    entities = [e["name"] for e in structure["entities"]]
    depts = [d["name"] for d in structure["departments"]]
    accounts = structure["account_rows"]
    tol = float(structure.get("tolerance", 0.02))

    # Row 0 group headers, Row 1 col headers
    # cols: 0 code, 1 name, then each entity debit/credit, total d/c, depts..., fee_total, check
    col = 0
    ws.cell(1, 1, entities[0] if entities else "")
    ws.cell(2, 1, "科目编码")
    ws.cell(2, 2, "科目名称")
    col = 3  # 1-based openpyxl: code=1 name=2, entity start 3
    ent_cols = {}  # name -> (debit_col, credit_col) 1-based
    for i, en in enumerate(entities):
        dcol, ccol = col, col + 1
        ws.cell(1, dcol, en)
        ws.cell(2, dcol, "本期发生借方")
        ws.cell(2, ccol, "本期发生贷方")
        ent_cols[en] = (dcol, ccol)
        col += 2
    total_d, total_c = col, col + 1
    ws.cell(1, total_d, "合计")
    ws.cell(2, total_d, "本期发生借方")
    ws.cell(2, total_c, "本期发生贷方")
    col += 2

    # parent groups for marketing/project
    dept_cols = {}
    for dname in depts:
        ws.cell(2, col, dname)
        # parent from structure
        parent = next((d.get("parent") for d in structure["departments"] if d["name"] == dname), None)
        if parent:
            ws.cell(1, col, parent)
        dept_cols[dname] = col
        col += 1
    fee_col, check_col = col, col + 1
    ws.cell(2, fee_col, "费用合计")
    ws.cell(2, check_col, "核对")

    # data rows
    for ri, acc in enumerate(accounts):
        r = ri + 3
        code = acc["code"]
        ws.cell(r, 1, code)
        ws.cell(r, 2, acc["name"])

        sum_d = sum_c = 0.0
        for en, (dc, cc) in ent_cols.items():
            ed = entity_data.get(en, {}).get(code, {})
            d = float(ed.get("debit", 0) or 0)
            c = float(ed.get("credit", 0) or 0)
            if d:
                ws.cell(r, dc, d)
            if c:
                ws.cell(r, cc, c)
            sum_d += d
            sum_c += c
        if sum_d:
            ws.cell(r, total_d, sum_d)
        if sum_c:
            ws.cell(r, total_c, sum_c)

        # department allocation (v0: direct hang only on leaf-ish codes)
        dept_amounts = defaultdict(float)
        hang = prefix_rule(code) if code else None
        # use net period amount: expenses usually debit, income credit
        net_for_hang = sum_d if sum_d else sum_c
        if hang and net_for_hang and hang in dept_cols:
            # income hangs on credit side amount
            if sum_c and not sum_d:
                dept_amounts[hang] += sum_c
            else:
                dept_amounts[hang] += sum_d

        for dname, amount in dept_amounts.items():
            ws.cell(r, dept_cols[dname], amount)

        # fee total & check as values (formula-friendly later)
        fee = sum(dept_amounts.values())
        # For rows with no dept split yet, fee=0 and check = total side
        # 核对: 收入类用贷方合计 - 费用合计；成本费用用借方合计 - 费用合计
        if sum_c and not sum_d:
            total_side = sum_c
        else:
            total_side = sum_d
        check = round(total_side - fee, 6)
        if fee:
            ws.cell(r, fee_col, fee)
        ws.cell(r, check_col, check)
        if abs(check) > tol and fee:  # only report when we attempted split
            report.append(f"核对不平 {code} {acc['name']}: 差额={check}")

    # 利润表简表
    ws2 = wb.create_sheet("利润表")
    ws2.cell(1, 1, "项目")
    for i, en in enumerate(entities):
        ws2.cell(1, i + 2, en)
    ws2.cell(1, len(entities) + 2, "合计")
    # map from major codes
    pl_map = [
        ("收入", "5101", "credit"),
        ("成本", "5401", "debit"),
        ("税金及附加", "5402", "debit"),
        ("销售费用", "5501", "debit"),
        ("管理费用", "5502", "debit"),
        ("研发费用", "5503", "debit"),
        ("财务费用", "5504", "debit"),
    ]
    for ri, (lab, code, side) in enumerate(pl_map):
        r = ri + 2
        ws2.cell(r, 1, lab)
        total = 0.0
        for i, en in enumerate(entities):
            ed = entity_data.get(en, {}).get(code, {})
            v = float(ed.get(side, 0) or 0)
            if v:
                ws2.cell(r, i + 2, v)
            total += v
        if total:
            ws2.cell(r, len(entities) + 2, total)

    return wb


def main():
    ap = argparse.ArgumentParser(description="部门费用归集分摊")
    ap.add_argument("--input-dir", default=str(SKILL / "工作区" / "input"))
    ap.add_argument("--out", default="", help="输出 xlsx 路径")
    ap.add_argument("--dry-run", action="store_true", help="只盘点规则与输入，不写表")
    args = ap.parse_args()

    input_dir = Path(args.input_dir).expanduser().resolve()
    structure = load_structure()
    report = []
    report.append("=== 部门费用归集分摊 allocate ===")
    report.append(f"输入目录: {input_dir}")
    report.append(f"模板科目行: {len(structure['account_rows'])} 部门列: {len(structure['departments'])}")
    report.append("5401：仅 540101/540103 整挂项目总监及助理，其余按人（斯佳标黄）。")
    report.append("本版：主体列 + 540101/540103/5402/5504 整挂。其余按人待样例。")

    entity_files = discover_entity_files(input_dir)
    report.append(f"识别到主体文件: { {k: str(v) for k,v in entity_files.items()} }")

    entity_data = {}
    for en, fpath in entity_files.items():
        try:
            entity_data[en] = read_balance_any(fpath)
            report.append(f"  读入 {en}: {len(entity_data[en])} 个科目 ← {fpath.name}")
        except Exception as e:
            report.append(f"  ⚠ 读失败 {en} {fpath}: {e}")

    if args.dry_run:
        out_rep = SKILL / "工作区" / "output" / "运行报告_allocate_dryrun.txt"
        out_rep.parent.mkdir(parents=True, exist_ok=True)
        out_rep.write_text("\n".join(report) + "\n", encoding="utf-8")
        print("\n".join(report))
        print(f"dry-run 报告: {out_rep}")
        return 0

    wb = build_workbook(structure, entity_data, report)
    out = Path(args.out) if args.out else SKILL / "工作区" / "output" / "部门科目余额表_产出.xlsx"
    out = out.expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    report.append(f"已写出: {out}")
    rep_path = out.parent / "运行报告_allocate.txt"
    rep_path.write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))
    print(f"报告: {rep_path}")
    # 不打印任何金额明细到成功摘要以外
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
