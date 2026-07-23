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
    partial = five.get("是否结账") == "否"
    return [
        item.get("case_id") or "",
        item.get("ar") or "",
        item.get("so") or "",
        item.get("sod") or five.get("实收SOD") or "",
        item.get("customer_masked") or "",
        item.get("flow_locate") or "",
        item.get("flow_order_suggest") or "",
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
        ("这单没回满：填完这几列后，记得在这一单上方插一行、复制原信息、"
         "把应收金额改成剩余待核销、是否结账填「否」（我们只提醒，不替你插）")
        if partial else "",
        item.get("reason") or "",
        item.get("channel") or "",
    ]


def row_hold(item: dict) -> List[Any]:
    codes = _codes()
    code = item.get("code") or ""
    info = codes.get(code) or {}
    return [
        item.get("case_id") or "",
        item.get("ar") or "",
        item.get("so") or "",
        item.get("sod") or "",
        item.get("customer_masked") or "",
        item.get("flow_locate") or "",
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
        "案例ID",
        "AR",
        "SO",
        "SOD",
        "客户(打码)",
        "这笔到账在流转表哪一行",
        "流转表单号列建议填",
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
        "部分核销提醒",
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
        "案例ID",
        "AR",
        "SO",
        "SOD",
        "客户(打码)",
        "这笔到账在流转表哪一行",
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

    # 按到账汇总：流转表「是否更新应收款」三态 + 公式/颜色策略（明妹 7-23 录音）
    summary = result.get("ar_summary") or []
    if summary:
        from flow_ledger import flow_status_policy

        wsum = wb.create_sheet("按到账汇总")
        wsum.append(
            [
                "AR",
                "本笔关联SO数",
                "流转表『是否更新应收款』建议填",
                "填法",
                "公式策略",
                "颜色标注",
                "还没处理完的SO",
                "流转表定位",
                "说明",
            ]
        )
        for cell in wsum[1]:
            cell.font = Font(bold=True)
        for s in summary:
            sug_raw = s.get("流转表_是否更新应收款_建议") or ""
            sug = sug_raw if sug_raw else "（空白）"
            pol = flow_status_policy(sug_raw)
            # 2026-07-23：空 + 部分 都算还要做；「是」才算完
            note = pol.get("人话") or (
                "已完成" if sug == "是" else "还要做（空或部分都要重更；红字=未更新）"
            )
            wsum.append(
                [
                    s.get("ar") or "",
                    s.get("so_count") or 0,
                    sug,
                    pol.get("填法") or "",
                    pol.get("公式策略") or "",
                    pol.get("颜色标注") or "",
                    " ".join(s.get("待处理SO") or []),
                    s.get("flow_locate") or "",
                    note,
                ]
            )
        wsum.freeze_panes = "A2"

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
