#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
跑批台账：**哪个核销日跑过了、跑到哪一步**，以及**哪几天从来没跑过**。

为什么必须有它（2026-07-25 立）：
  旧版取数写死 `--date yesterday`，"昨天"是相对**运行那天**算的。
  → 核销可能发生在周末；如果漏天只检查工作日，周六、周日的数据会永远没人管。
  → 她请假两天、出差一周、系统故障没跑，同理静默漏。
  漏一天 = 那天的到账永远不会回填进盈亏表，而且**没有任何地方看得出来**。

所以：每跑一个核销日就在这里登记；每次开跑前先查空档，
**有空档就交给编排器从最早日期自动补跑**，而不是只闷头跑昨天。

一天一个核销日，**不合并**：合并会让 AR 覆盖率校验、幂等校验和她对着清单
逐行核对全部失真（她核的是"这一天的到账"）。

用法：
    python3 scripts/batch_ledger.py gaps   --workspace 工作区 [--through 2026-07-25]
    python3 scripts/batch_ledger.py record --workspace 工作区 --hexiao-date 2026-07-24 \
            --stage classified --payments 4
    python3 scripts/batch_ledger.py show   --workspace 工作区 [--limit 15]
退出码：gaps 有空档=1（供编排器读取补跑范围，不代表需要人工确认），无空档=0；其余 0/2。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import atomic_json_store  # noqa: E402
import common  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

LEDGER_NAME = "跑批台账.json"
# 跑到哪一步了。写表(applied)才算这一天真正收工。
STAGES = ("fetched", "classified", "listed", "applied")
STAGE_CN = {
    "fetched": "只取了数",
    "classified": "判过了·没出清单",
    "listed": "出过清单·还没写表",
    "applied": "已写表·收工",
}


def _empty_ledger() -> dict:
    return {"start_date": None, "runs": {}}


def _validate_ledger(data: dict) -> None:
    if not isinstance(data, dict):
        raise TypeError("跑批台账必须是对象")
    if data.get("start_date") is not None and not isinstance(data.get("start_date"), str):
        raise TypeError("start_date 必须是日期文本或空值")
    if not isinstance(data.get("runs", {}), dict):
        raise TypeError("runs 必须是对象")


def ledger_path(workspace: Path) -> Path:
    d = Path(workspace) / "03_台账"
    d.mkdir(parents=True, exist_ok=True)
    return d / LEDGER_NAME


def load(workspace: Path) -> dict:
    p = ledger_path(workspace)
    data = atomic_json_store.load_json(
        p,
        default_factory=_empty_ledger,
        validate=_validate_ledger,
    )
    data.setdefault("runs", {})
    data.setdefault("start_date", None)
    return data


def save(workspace: Path, data: dict) -> Path:
    p = ledger_path(workspace)
    atomic_json_store.update_json(
        p,
        default_factory=_empty_ledger,
        validate=_validate_ledger,
        mutate=lambda current: current.clear() or current.update(data),
    )
    return p


def record(
    workspace: Path,
    hexiao_date: dt.date,
    stage: str,
    *,
    payments: Optional[int] = None,
    counts: Optional[dict] = None,
    written: Optional[dict] = None,
    note: str = "",
) -> dict:
    """登记一次跑批。同一天可多次登记，stage 只前进不后退。"""
    if stage not in STAGES:
        raise ValueError(f"stage 只能是 {STAGES}，收到 {stage!r}")
    key = hexiao_date.isoformat()
    now = dt.datetime.now().isoformat(timespec="seconds")

    def mutate(data: dict) -> dict:
        data.setdefault("runs", {})
        rec = data["runs"].get(key) or {"hexiao_date": key, "first_run_at": now}
        old = rec.get("stage")
        # 只前进不后退：重跑判定不该把「已写表」降级成「判过了」
        if old is None or STAGES.index(stage) >= STAGES.index(old):
            rec["stage"] = stage
        rec["last_run_at"] = now
        if payments is not None:
            rec["payment_count"] = int(payments)
            rec["empty_batch"] = int(payments) == 0
        if counts is not None:
            rec["counts"] = counts
        if written is not None:
            rec["written"] = written
        if note:
            rec["note"] = note
        if stage == "applied":
            rec["applied_at"] = now
        data["runs"][key] = rec
        if not data.get("start_date"):
            # 第一次用 = 起点。之前的历史不算漏（否则会从年初开始报一堆空档）
            data["start_date"] = key
        return dict(rec)

    return atomic_json_store.update_json(
        ledger_path(workspace),
        default_factory=_empty_ledger,
        validate=_validate_ledger,
        mutate=mutate,
    )


def done_dates(data: dict) -> set:
    """算「这一天已经处理过」的集合。空批（那天真的没人核销）也算处理过。"""
    out = set()
    for k, rec in (data.get("runs") or {}).items():
        d = common.norm_date(k)
        if d is None:
            continue
        # 只取过数没判过 → 不算处理过，她多半是跑了一半被打断
        if rec.get("stage") in ("classified", "listed", "applied"):
            out.add(d)
    return out


def find_gaps(
    workspace: Path,
    through: Optional[dt.date] = None,
    include_weekend: bool = True,
) -> Dict[str, object]:
    """
    从起点到 through 之间，**没被处理过**的核销日。
    默认包含周末；只有明确传入 include_weekend=False 时才只检查工作日。
    """
    data = load(workspace)
    through = through or (dt.date.today() - dt.timedelta(days=1))
    start = common.norm_date(data.get("start_date"))
    done = done_dates(data)
    if start is None:
        # 台账还是空的：第一次用，没有"漏"的概念
        return {
            "first_use": True,
            "gaps": [],
            "skipped_weekend": [],
            "through": through,
            "last_done": None,
        }
    gaps: List[dt.date] = []
    skipped: List[dt.date] = []
    cur = start
    while cur <= through:
        if cur not in done:
            if cur.weekday() >= 5 and not include_weekend:
                skipped.append(cur)
            else:
                gaps.append(cur)
        cur += dt.timedelta(days=1)
    return {
        "first_use": False,
        "gaps": gaps,
        "skipped_weekend": skipped,
        "through": through,
        "last_done": max(done) if done else None,
    }


def suggest_date(workspace: Path, today: Optional[dt.date] = None) -> Dict[str, object]:
    """
    建议这次该跑哪个核销日 = **最早那个没跑过的**（有空档就先补最早的），
    否则 = 昨天。附上理由，供编排器直接执行。
    """
    today = today or dt.date.today()
    default = today - dt.timedelta(days=1)
    info = find_gaps(workspace, through=default)
    gaps = info["gaps"]
    if gaps:
        return {
            "date": gaps[0],
            "reason": (
                f"有 {len(gaps)} 天没跑过（{'、'.join(d.isoformat() for d in gaps[:5])}"
                f"{' 等' if len(gaps) > 5 else ''}），建议**从最早的一天开始补**，一天一批"
            ),
            "gaps": gaps,
            "info": info,
        }
    return {
        "date": default,
        "reason": "没有漏掉的天，按常规跑昨天",
        "gaps": [],
        "info": info,
    }


def _cmd_gaps(args) -> int:
    ws = common.resolve_workspace(args.workspace, quiet=True)
    through = common.resolve_batch_date(args.through) if args.through else None
    info = find_gaps(ws, through=through, include_weekend=args.all_days)
    if info["first_use"]:
        print("跑批台账还是空的（第一次用）——本次跑完会自动记下起点，之后才能查漏天。")
        return 0
    gaps = info["gaps"]
    if not gaps:
        print(
            f"没有漏掉的核销日：起点至 {common.date_cn(info['through'])} 之间"
            f"{'（周末未计入）' if not args.all_days else ''}全都处理过了。"
        )
        return 0
    print(f"⚠ 有 {len(gaps)} 个核销日从来没跑过：")
    for d in gaps:
        print(f"  · {common.date_cn(d)}")
    if info["skipped_weekend"] and not args.all_days:
        print(f"（另有 {len(info['skipped_weekend'])} 个周末未计入；要连周末一起补加 --all-days）")
    print("→ 一天一批，从最早的那天开始补；别把几天合成一批跑。")
    # 给编排器一段短提示：读取补跑范围后直接继续，不把流程变成人工确认闸。
    short = "、".join(f"{d.month}-{d.day}" for d in gaps[:6]) + ("…" if len(gaps) > 6 else "")
    print(
        f'\n【自动处理计划】有 {len(gaps)} 天待处理：{short}。'
        "编排器将从最早日期开始逐日取数、判定、校验并写入；不合并日期，也不等待人工确认。"
    )
    return 1


def _cmd_record(args) -> int:
    ws = common.resolve_workspace(args.workspace, quiet=True)
    d = common.resolve_batch_date(args.hexiao_date)
    if d is None:
        print(f"ERROR: 认不出核销日期 {args.hexiao_date!r}", file=sys.stderr)
        return 2
    counts = json.loads(args.counts) if args.counts else None
    written = json.loads(args.written) if args.written else None
    rec = record(
        ws, d, args.stage,
        payments=args.payments, counts=counts, written=written, note=args.note,
    )
    print(f"跑批台账已登记：{common.date_cn(d)} → {STAGE_CN.get(rec['stage'], rec['stage'])}")
    return 0


def _cmd_show(args) -> int:
    ws = common.resolve_workspace(args.workspace, quiet=True)
    data = load(ws)
    runs = data.get("runs") or {}
    if not runs:
        print("跑批台账为空（还没跑过任何一天）。")
        return 0
    keys = sorted(runs, reverse=True)[: args.limit]
    print(f"跑批台账（起点 {data.get('start_date')}，共 {len(runs)} 天）：")
    for k in keys:
        r = runs[k]
        d = common.norm_date(k)
        n = r.get("payment_count")
        tail = "（那天没人核销·空批）" if r.get("empty_batch") else (f"到账 {n} 笔" if n is not None else "")
        print(f"  · {common.date_cn(d)}  {STAGE_CN.get(r.get('stage'), r.get('stage'))}  {tail}")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="跑批台账（哪天跑过 / 哪天漏了）")
    ap.add_argument("action", choices=["gaps", "record", "show"])
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--through", default="", help="查到哪天为止（默认昨天）")
    date_scope = ap.add_mutually_exclusive_group()
    date_scope.add_argument(
        "--all-days",
        dest="all_days",
        action="store_true",
        default=True,
        help="周末也算该跑的日子（当前默认行为，保留参数兼容旧调用）",
    )
    date_scope.add_argument(
        "--workdays-only",
        dest="all_days",
        action="store_false",
        help="只检查周一至周五",
    )
    ap.add_argument("--hexiao-date", default="", help="record：这批的核销日期")
    ap.add_argument("--stage", default="classified", choices=list(STAGES))
    ap.add_argument("--payments", type=int, default=None, help="record：这批到账笔数")
    ap.add_argument("--counts", default="", help="record：判定计数 json 串")
    ap.add_argument("--written", default="", help="record：写入笔数 json 串")
    ap.add_argument("--note", default="")
    ap.add_argument("--limit", type=int, default=15)
    args = ap.parse_args(argv)

    if args.action == "gaps":
        return _cmd_gaps(args)
    if args.action == "record":
        if not args.hexiao_date:
            print("ERROR: record 必须给 --hexiao-date", file=sys.stderr)
            return 2
        return _cmd_record(args)
    return _cmd_show(args)


if __name__ == "__main__":
    sys.exit(main())
