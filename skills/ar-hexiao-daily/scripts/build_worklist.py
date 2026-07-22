#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第7步：判定结果 → 今日工作清单 xlsx（三页签）。
只写 工作区/04_产出/，不写用户任何 Excel。
今天能填：按 SO 定位说明 + 五列该填值 + 当前值；禁止行号定位。
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


def load_result(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _codes() -> dict:
    return common.load_codes()


def row_auto(item: dict) -> List[Any]:
    five = item.get("five_cols") or {}
    cur = item.get("current_values") or {}
    return [
        item.get("ar") or "",
        item.get("so") or "",
        item.get("sod") or five.get("实收SOD") or "",
        item.get("customer_masked") or "",
        item.get("locate_hint") or (
            f"在盈亏『明细』按「新智云单号」筛选：{item.get('so') or ''}（禁止用行号）"
        ),
        five.get("计提"),
        five.get("回款明细"),
        five.get("是否结账"),
        five.get("收款时间"),
        five.get("收款方式"),
        cur.get("计提"),
        cur.get("回款明细"),
        cur.get("是否结账"),
        cur.get("收款时间"),
        cur.get("收款方式"),
        cur.get("实收SOD"),
        item.get("reason") or "",
        item.get("channel") or "",
    ]


def row_hold(item: dict) -> List[Any]:
    codes = _codes()
    code = item.get("code") or ""
    info = codes.get(code) or {}
    return [
        item.get("ar") or "",
        item.get("so") or "",
        item.get("sod") or "",
        item.get("customer_masked") or "",
        code,
        info.get("名称") or code,
        item.get("reason") or info.get("原因") or "",
        info.get("建议动作") or info.get("大白话") or "",
        item.get("locate_hint") or "",
        ",".join(str(x) for x in (item.get("candidates") or [])),
    ]


def row_exc(item: dict) -> List[Any]:
    return row_hold(item)


def build_workbook(result: dict, out_path: Path) -> Path:
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    # 今天能填
    ws = wb.active
    ws.title = "今天能填"
    headers_auto = [
        "AR",
        "SO",
        "SOD",
        "客户(打码)",
        "怎么找到这行(按SO筛选)",
        "应填_计提",
        "应填_回款明细",
        "应填_是否结账",
        "应填_收款时间",
        "应填_收款方式",
        "当前_计提",
        "当前_回款明细",
        "当前_是否结账",
        "当前_收款时间",
        "当前_收款方式",
        "当前_实收SOD",
        "判定依据",
        "通道",
    ]
    ws.append(headers_auto)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for item in result.get("auto") or []:
        ws.append(row_auto(item))
    ws.freeze_panes = "A2"

    # 挂账待办
    wh = wb.create_sheet("挂账待办")
    headers_h = [
        "AR",
        "SO",
        "SOD",
        "客户(打码)",
        "码",
        "原因名",
        "说明",
        "你该做什么",
        "定位提示",
        "候选行(仅参考)",
    ]
    wh.append(headers_h)
    for cell in wh[1]:
        cell.font = Font(bold=True)
    for item in result.get("hold") or []:
        wh.append(row_hold(item))
    wh.freeze_panes = "A2"

    # 异常
    we = wb.create_sheet("异常")
    we.append(headers_h)
    for cell in we[1]:
        cell.font = Font(bold=True)
    for item in result.get("exception") or []:
        we.append(row_exc(item))
    we.freeze_panes = "A2"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="生成今日工作清单（只写产出目录）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--result", default="", help="判定结果 json；默认取 04_产出 最新")
    ap.add_argument("--out", default="", help="清单 xlsx 路径")
    args = ap.parse_args(argv)

    common.ensure_out_dirs()
    ws = Path(args.workspace)
    out_dir = ws / "04_产出" if (ws / "04_产出").is_dir() else common.OUT_DIR

    if args.result:
        result_path = Path(args.result)
    else:
        cands = sorted(out_dir.glob("判定结果_*.json"))
        if not cands:
            print("ERROR: 找不到判定结果 json，请先跑 classify_hexiao", file=sys.stderr)
            return 2
        result_path = cands[-1]

    result = load_result(result_path)
    today = dt.date.today().strftime("%Y%m%d")
    out_path = Path(args.out) if args.out else out_dir / f"今日工作清单_{today}.xlsx"
    build_workbook(result, out_path)

    c = result.get("counts") or {}
    print(
        f"清单已生成 auto={c.get('auto', len(result.get('auto') or []))} "
        f"hold={c.get('hold', len(result.get('hold') or []))} "
        f"exception={c.get('exception', len(result.get('exception') or []))}"
    )
    print(f"结果: {out_path}")
    # 运行报告
    report = out_dir / f"运行报告_{today}.txt"
    report.write_text(
        f"工作清单 {out_path.name}\n"
        f"来源判定 {result_path.name}\n"
        f"计数 {c}\nE码 {result.get('e_code_dist')}\n",
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
