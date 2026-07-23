#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第7步：判定 + 写入前校验 → **一份**《核销日清》xlsx。

2026-07-23 明妹当面要求（原话）：
  「这问题都都给我在这一个里边去显示吧」「第一列应该是单号」
  「用 sheet 名字就在这都写成一列就行了」
所以这一版把旧的《今日工作清单》(4 页签) + 《回填审核单》(5 页签) **合成一份**：
主页签《今日清单》一行一个 SOD，**第一列单号、第二列状态**，问题类型变成一列。

还修了她当场问倒的那个 bug —— 「我明明填过了他为什么没跳过」：
旧版「今天能填」直接倒 classify 的 auto，**没过写入校验**，所以把昨天已经填过的
23 笔又列了一遍（真跑下来那 23 笔全是"已填过"，实际要填 0 笔）。
本版清单的状态**以 validate_plan 的结论为准**，跳过的就明明白白写"已填过·跳过"。

只写 04_产出/，不写用户任何 Excel。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402

# 状态 → (排序权重, 颜色, 这行该干嘛)
STATUS = {
    "今天要填": (1, "C6EFCE", "按「应填」那几列填进盈亏『明细』对应行"),
    "已填过·跳过": (2, "D9D9D9", "不用管，你之前已经填过而且跟这次算的一样"),
    "冲突·需你定": (3, "FFC7CE", "这行已经填过但跟这次算的不一样，看「差异」那列，你定按哪个来"),
    "挂账待办": (4, "FFEB9C", "今天填不了，按「怎么办」处理；处理完说「重扫挂账」"),
    "异常": (5, "F8CBAD", "数据不对劲，别填，按「怎么办」先查清楚"),
}

HEADERS = [
    "单号", "状态", "怎么办",
    "SO", "到账号(AR)", "客户(打码)", "本次金额",
    "应填_计提", "应填_回款明细", "应填_是否结账", "应填_收款时间", "应填_收款方式", "应填_实收SOD",
    "当前_计提", "当前_回款明细", "当前_是否结账", "当前_收款时间", "当前_收款方式", "当前_实收SOD",
    "差异(一眼扫)", "怎么找到这行", "这笔到账在流转表哪一行", "流转表单号列建议填",
    "判定依据", "码",
]

FIVE = ["计提", "回款明细", "是否结账", "收款时间", "收款方式"]


def _norm(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    try:
        return f"{float(s):.2f}"
    except (TypeError, ValueError):
        return s


def _diff_text(five: dict, cur: dict) -> str:
    parts = []
    for k in FIVE + ["实收SOD"]:
        if five.get(k) is None:
            continue
        a, b = _norm(cur.get(k)), _norm(five.get(k))
        if a != b:
            parts.append(f"{k}：表里={a or '(空)'} → 这次算={b}")
    return "；".join(parts)


# 这些码的 reason 是**逐笔算出来的、自带操作指引**（插哪一行、候选 SOD 是哪几个…），
# 比 config 里的通用建议有用得多 → 「怎么办」直接用 reason。
SPECIFIC_CODES = {"E4", "E5", "E7", "E8"}


def _row(item: dict, status: str, codes: dict, action_override: str = "") -> List[Any]:
    five = item.get("five_cols") or {}
    cur = item.get("current_values") or {}
    code = item.get("code") or ""
    info = codes.get(code) or {}
    sod = item.get("sod") or five.get("实收SOD") or ""
    warn = "⚠" in (item.get("reason") or "")
    if warn:
        # 判定里带⚠的（比如智云交付额和她表里应收对不上）必须顶到「怎么办」，别只埋在依据列
        action = str(item["reason"]).split("；", 1)[-1]
    elif action_override:
        action = action_override
    elif code in SPECIFIC_CODES and item.get("reason"):
        action = item["reason"]
    elif code:
        action = info.get("建议动作") or info.get("大白话") or STATUS[status][2]
    else:
        action = STATUS[status][2]
    return [
        # 整笔到账级的挂起（没 SO 也没 SOD）不能留空——退化显示到账号，否则她定位不到
        sod or item.get("so") or item.get("ar") or "",
        status,
        action,
        item.get("so") or "",
        item.get("ar") or "",
        item.get("customer_masked") or "",
        five.get("回款明细"),
        five.get("计提"), five.get("回款明细"), five.get("是否结账"),
        five.get("收款时间"), five.get("收款方式"), five.get("实收SOD"),
        cur.get("计提"), cur.get("回款明细"), cur.get("是否结账"),
        cur.get("收款时间"), cur.get("收款方式"), cur.get("实收SOD"),
        _diff_text(five, cur) if status == "冲突·需你定" else "",
        item.get("locate_hint") or "",
        item.get("flow_locate") or "",
        item.get("flow_order_suggest") or "",
        item.get("reason") or info.get("原因") or "",
        code,
    ]


def collect_rows(result: dict, checked: Optional[dict]) -> List[List[Any]]:
    codes = common.load_codes()
    rows: List[List[Any]] = []

    if checked:
        by_status = [
            (checked.get("write") or [], "今天要填"),
            (checked.get("skip") or [], "已填过·跳过"),
            (checked.get("conflict") or [], "冲突·需你定"),
        ]
        for items, st in by_status:
            for it in items:
                extra = ((it.get("_check") or {}).get("reason") or "") if st != "今天要填" else ""
                rows.append(_row(it, st, codes, action_override=extra or STATUS[st][2]))
    else:
        for it in result.get("auto") or []:
            rows.append(_row(it, "今天要填", codes))

    for it in result.get("hold") or []:
        rows.append(_row(it, "挂账待办", codes))
    for it in result.get("exception") or []:
        rows.append(_row(it, "异常", codes))

    rows.sort(key=lambda r: (STATUS.get(r[1], (9,))[0], str(r[4]), str(r[3]), str(r[0])))
    return rows


def build_workbook(result: dict, checked: Optional[dict], out_path: Path) -> Path:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = openpyxl.Workbook()

    # ── 使用说明 ────────────────────────────────────────────
    ws0 = wb.active
    ws0.title = "先看这里"
    counts = result.get("counts") or {}
    n = {k: 0 for k in STATUS}
    for r in collect_rows(result, checked):
        n[r[1]] = n.get(r[1], 0) + 1
    lines = [
        ["《核销日清》—— 一份就够，别的表都不用开"],
        [""],
        [f"这次一共 {result.get('payment_count', '?')} 笔到账，拆成 {counts.get('total', 0)} 个订单行。"],
        [""],
        [f"① 今天要填     {n.get('今天要填', 0):>4} 行  ← 只有这些要动手"],
        [f"② 已填过·跳过  {n.get('已填过·跳过', 0):>4} 行  ← 不用管"],
        [f"③ 冲突·需你定  {n.get('冲突·需你定', 0):>4} 行  ← 你看一眼定个方向"],
        [f"④ 挂账待办     {n.get('挂账待办', 0):>4} 行  ← 今天填不了，看「怎么办」"],
        [f"⑤ 异常         {n.get('异常', 0):>4} 行  ← 数据不对劲，先别填"],
        [""],
        ["怎么用："],
        ["  1. 翻到《今日清单》，按「状态」列筛「今天要填」。"],
        ["  2. 一行一个单号。拿「怎么找到这行」去盈亏『明细』筛出那一行。"],
        ["  3. 照「应填_xxx」几列填进去；「当前_xxx」是你表里现在的值，可对照。"],
        ["  4. 「冲突·需你定」那几行看「差异(一眼扫)」，决定按哪个来。"],
        ["  5. 都看完了跟我说「确认」，我再把「今天要填」的写进你的副本。"],
        [""],
        ["※ 这份只是清单，程序还没动你的任何表。"],
        ["※ 智云永远不写。"],
    ]
    for ln in lines:
        ws0.append(ln)
    ws0["A1"].font = Font(bold=True, size=14)
    ws0.column_dimensions["A"].width = 78

    # ── 今日清单（主表）────────────────────────────────────
    ws = wb.create_sheet("今日清单")
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(vertical="center", wrap_text=True)
    rows = collect_rows(result, checked)
    for r in rows:
        ws.append(r)
    for i, r in enumerate(rows, start=2):
        color = STATUS.get(r[1], (9, "FFFFFF"))[1]
        fill = PatternFill("solid", fgColor=color)
        ws.cell(row=i, column=1).fill = fill
        ws.cell(row=i, column=2).fill = fill
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{chr(64 + len(HEADERS))}{len(rows) + 1}" if len(HEADERS) <= 26 else None
    for col, w in (("A", 16), ("B", 14), ("C", 46), ("D", 14), ("E", 14), ("F", 16), ("T", 52), ("U", 40)):
        ws.column_dimensions[col].width = w

    # ── 按到账汇总（流转表「是否更新应收款」建议）──────────
    summary = result.get("ar_summary") or []
    if summary:
        try:
            from flow_ledger import flow_status_policy
        except Exception:  # pragma: no cover
            def flow_status_policy(_s):
                return {}

        wsum = wb.create_sheet("流转表怎么填")
        wsum.append([
            "到账号(AR)", "本笔关联SO数", "订单行数", "『是否更新应收款』建议填",
            "填法", "公式策略", "颜色标注", "还没处理完的SO", "流转表定位", "说明",
        ])
        for c in wsum[1]:
            c.font = Font(bold=True)
        for s in summary:
            raw = s.get("流转表_是否更新应收款_建议") or ""
            pol = flow_status_policy(raw)
            wsum.append([
                s.get("ar") or "", s.get("so_count") or 0, s.get("行数") or 0,
                raw if raw else "（空白）",
                pol.get("填法") or "", pol.get("公式策略") or "", pol.get("颜色标注") or "",
                " ".join(s.get("待处理SO") or []),
                s.get("flow_locate") or "",
                pol.get("人话") or ("已完成" if raw == "是" else "还要做（空或部分都要重更）"),
            ])
        wsum.freeze_panes = "A2"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def _latest(out_dir: Path, pattern: str) -> Optional[Path]:
    c = [p for p in sorted(out_dir.glob(pattern)) if not p.name.startswith("~$")]
    return c[-1] if c else None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="生成《核销日清》（一份合一，只写产出目录）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--result", default="", help="判定结果 json；默认取 04_产出 最新")
    ap.add_argument(
        "--checked", default="",
        help="validate_plan 产出的写入计划_校验后.json；默认取 04_产出 最新。"
             "给了才分得出「今天要填 / 已填过跳过 / 冲突」",
    )
    ap.add_argument("--out", default="", help="清单 xlsx 路径")
    args = ap.parse_args(argv)

    ws = Path(args.workspace)
    common.ensure_out_dirs(ws)
    out_dir = ws / "04_产出"

    result_path = Path(args.result) if args.result else _latest(out_dir, "判定结果_*.json")
    if not result_path or not result_path.is_file():
        print("ERROR: 找不到判定结果 json，请先跑 classify_hexiao.py", file=sys.stderr)
        return 2
    result = json.loads(result_path.read_text(encoding="utf-8"))

    checked_path = Path(args.checked) if args.checked else _latest(out_dir, "写入计划_校验后*.json")
    checked = None
    if checked_path and checked_path.is_file():
        checked = json.loads(checked_path.read_text(encoding="utf-8"))
    else:
        print(
            "WARN: 没有写入计划_校验后.json —— 清单里所有可填的都会标「今天要填」，"
            "**分不出你昨天已经填过的**。请先跑 validate_plan.py 再跑本脚本。",
            file=sys.stderr,
        )

    today = dt.date.today().strftime("%Y%m%d")
    out_path = Path(args.out) if args.out else out_dir / f"核销日清_{today}.xlsx"
    build_workbook(result, checked, out_path)

    rows = collect_rows(result, checked)
    tally: Dict[str, int] = {}
    for r in rows:
        tally[r[1]] = tally.get(r[1], 0) + 1
    print("《核销日清》已生成：" + " / ".join(f"{k} {v}" for k, v in tally.items()))
    print(f"结果: {out_path}")

    (out_dir / f"运行报告_{today}.txt").write_text(
        f"清单 {out_path.name}\n"
        f"判定 {result_path.name}\n"
        f"校验 {checked_path.name if checked else '(未跑 validate_plan)'}\n"
        f"到账笔数 {result.get('payment_count')}\n"
        f"计数 {result.get('counts')}\nE码 {result.get('e_code_dist')}\n"
        f"清单分布 {tally}\n",
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
