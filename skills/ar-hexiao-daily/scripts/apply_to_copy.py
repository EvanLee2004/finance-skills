#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把校验通过的计划写进盈亏**工作副本**（plan → validate → **execute**）。

2026-07-23 明妹口径：
  - **只改「明细」sheet**，其它 sheet 一个字节不许动
  - 可以长期用同一份副本；输出仍写**新文件**（她对照原版其它 sheet 总数验收）
  - 回填内容来自第 6 步智云判定结果，不是流转表

安全设计：
  输入 = 她给的盈亏副本（只读打开）+ 校验后的计划
  输出 = 04_产出/盈亏核算表_已回填_日期.xlsx（新文件）+ 变更清单
  用 OOXML 补丁只改明细格，避免 openpyxl 毁图/透视

写完立刻回读逐格比对；对不上非 0 退出。
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
import sys
from pathlib import Path
from typing import Dict, List

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402
from validate_plan import FIVE, _norm, read_ledger_rows  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def locate_columns(ws, aliases) -> Dict[str, int]:
    """找到『明细』的表头行与五列（+SO/SOD 用于写后核对）。返回 1-based 列号。"""
    all_rows = list(ws.iter_rows(values_only=True))
    hrow, headers = common.find_header_row(
        all_rows, "盈亏明细", ["SO", "SOD", "计提", "回款明细", "是否结账"], aliases
    )
    cols = common.resolve_columns(
        headers,
        "盈亏明细",
        ["SO", "SOD", "计提", "回款明细", "是否结账", "收款时间", "收款方式"],
        aliases,
    )
    return {k: v + 1 for k, v in cols.items()}  # openpyxl 是 1-based


def write_plan(src: Path, out: Path, items: List[dict]) -> List[dict]:
    """
    把 items 的五列写进 out（out 是 src 的无损副本）。返回变更明细。

    **不用 openpyxl 保存**：实测她的真表用 openpyxl 载入再保存会丢 5 个 drawing、
    1 张内嵌图片和若干 rels（74 个部件变 59）。改用 `xlsx_patch` 只补丁目标格，
    其余部件逐字节原样写回。
    """
    import openpyxl
    import xlsx_patch

    # 先用只读模式拿列位置与改前值（不保存，纯读）
    wb = openpyxl.load_workbook(str(src), read_only=True, data_only=True)
    if "明细" not in wb.sheetnames:
        names = list(wb.sheetnames)
        wb.close()
        raise ValueError(
            f"盈亏表无『明细』sheet（明妹规定只许改明细）：现有={names}"
        )
    # 铁律：后续 patch 目标名写死「明细」，禁止调用方改成别的 sheet
    target_sheet = "明细"
    cols = locate_columns(wb[target_sheet], common.load_aliases())
    wb.close()

    before_rows = read_ledger_rows(src)
    changes: List[dict] = []
    edits: List = []
    for it in items:
        r = int(it["ledger_row_ref"])
        five = it.get("five_cols") or {}
        before = before_rows.get(r, {})
        for k in FIVE:
            v = five.get(k)
            if v is None:
                continue  # 部分核销时计提留空——留空就是留空，不写 0
            if k == "收款时间":
                v = common.norm_date(v) or v
            elif k in ("计提", "回款明细"):
                v = float(v)
            edits.append((r, cols[k], v))
        sod = five.get("实收SOD") or it.get("sod")
        if sod:
            edits.append((r, cols["SOD"], sod))
        changes.append(
            {
                "案例ID": it.get("case_id"),
                "行号": r,
                "SO": it.get("so"),
                "SOD": sod,
                "改前": {k: _norm(before.get(k)) for k in FIVE},
                "改后": {k: _norm(five.get(k)) for k in FIVE},
            }
        )
    xlsx_patch.patch_cells(src, out, target_sheet, edits)

    # 写完立刻自证没搞坏她的表：少一个部件都算失败
    lost = xlsx_patch.parts_diff(src, out)
    if lost:
        raise ValueError(f"写入后工作簿部件缺失（不该发生）：{lost[:5]}")
    return changes


def verify_written(out: Path, items: List[dict]) -> List[str]:
    """回读逐格比对——写完必须证明真写对了，而不是"保存没报错就算成"。"""
    rows = read_ledger_rows(out)
    problems: List[str] = []
    for it in items:
        r = int(it["ledger_row_ref"])
        five = it.get("five_cols") or {}
        row = rows.get(r)
        if row is None:
            problems.append(f"第 {r} 行写完却读不到")
            continue
        for k in FIVE:
            if five.get(k) is None:
                continue
            if _norm(row.get(k)) != _norm(five.get(k)):
                problems.append(
                    f"第 {r} 行 {k}：期望 {_norm(five.get(k))!r} 实际 {_norm(row.get(k))!r}"
                )
    return problems


def write_change_report(changes: List[dict], path: Path) -> None:
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "变更清单"
    headers = ["案例ID", "行号", "SO", "SOD"] + [f"改前_{k}" for k in FIVE] + [f"改后_{k}" for k in FIVE]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
    for ch in changes:
        ws.append(
            [ch["案例ID"], ch["行号"], ch["SO"], ch["SOD"]]
            + [ch["改前"][k] for k in FIVE]
            + [ch["改后"][k] for k in FIVE]
        )
    ws.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="把校验通过的计划写进副本（不动原件）")
    ap.add_argument("--checked", required=True, help="validate_plan.py 产出的校验后计划")
    ap.add_argument("--ledger", required=True, help="她给的盈亏副本（只读，作基线）")
    ap.add_argument("--out", default="", help="回填后的新文件路径")
    ap.add_argument("--report", default="", help="变更清单 xlsx")
    ap.add_argument("--force", action="store_true", help="即使计划里有 conflict 也照写可写的那部分")
    args = ap.parse_args(argv)

    checked_p, src = Path(args.checked), Path(args.ledger)
    for p, name in ((checked_p, "校验后计划"), (src, "盈亏副本")):
        if not p.is_file():
            print(f"ERROR: 找不到{name} {p}", file=sys.stderr)
            return 2

    plan = json.loads(checked_p.read_text(encoding="utf-8"))
    writable = plan.get("write") or []
    conflicts = plan.get("conflict") or []
    if conflicts and not args.force:
        print(
            f"ERROR: 计划里还有 {len(conflicts)} 笔冲突没处理，先看清楚再写。\n"
            f"  （确认要跳过冲突、只写可写的那部分，就加 --force）",
            file=sys.stderr,
        )
        return 2
    if not writable:
        print("没有可写的笔（可能都已经填过了）。什么都没改。")
        return 0

    today = dt.date.today().strftime("%Y%m%d")
    out = Path(args.out) if args.out else src.parent.parent / "04_产出" / f"盈亏核算表_已回填_{today}.xlsx"
    report = Path(args.report) if args.report else out.with_name(f"变更清单_{today}.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)

    try:
        changes = write_plan(src, out, writable)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    problems = verify_written(out, writable)
    write_change_report(changes, report)

    print(f"已写入 {len(changes)} 笔 → {out}")
    print(f"变更清单 → {report}")
    print(f"基线未动（她给的副本）→ {src}")
    if problems:
        print("⚠ 写后回读比对不符：", file=sys.stderr)
        for x in problems[:10]:
            print(f"  - {x}", file=sys.stderr)
        return 1
    print("写后回读逐格比对：全部一致 ✓")
    return 0


if __name__ == "__main__":
    sys.exit(main())
