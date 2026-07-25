"""
覆盖台账：**哪些「下单日期」已经被统计过**，以及**哪几天从来没统计过**。

为什么要有它（2026-07-25 立）：
  日期窗口只认周末、**不认法定节假日**。长假后第一个上班日跑，规则只给"昨天"一天
  → 假期里的下单全漏，而且**事后没有任何地方看得出来**。
  （实测：10-1~10-7 放假、10-8 周四上班 → 那天只抓 10-07，10-1~10-6 全漏。）

所以每跑一次就把**覆盖到的下单日**登记下来；下次跑时对比，
**断档直接印在产物的《统计区间》页上**，亮晶打开就看得见，不用记也不用算。

台账落在输出目录（跟产物放一起），WorkBuddy 每天跑都读得到同一份。

⚠ 补断档**别用"一天一天 --today 补跑"**：窗口按运行日倒推、逐天跑会互相重叠，
   同一天被算好几遍。正确做法见 SKILL / 手册（离线模式喂一张覆盖整段的导出表）。
"""

from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable

LEDGER_NAME = "覆盖台账.json"
WEEKDAY_CN = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")


def cn(d: date) -> str:
    """给她看的日期写法：2026-07-24（周五）。"""
    return f"{d.isoformat()}（{WEEKDAY_CN[d.weekday()]}）"


def ledger_path(out_dir: Path | str) -> Path:
    d = Path(out_dir)
    d.mkdir(parents=True, exist_ok=True)
    return d / LEDGER_NAME


def load(out_dir: Path | str) -> dict:
    p = ledger_path(out_dir)
    if not p.is_file():
        return {"covered": {}}
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
    except (ValueError, OSError):
        return {"covered": {}}
    data.setdefault("covered", {})
    return data


def _parse(s: str) -> date | None:
    try:
        return date.fromisoformat(str(s)[:10])
    except (TypeError, ValueError):
        return None


def covered_dates(out_dir: Path | str) -> set[date]:
    out: set[date] = set()
    for k in (load(out_dir).get("covered") or {}):
        d = _parse(k)
        if d:
            out.add(d)
    return out


def daterange(start: date, end: date) -> list[date]:
    n = (end - start).days
    return [start + timedelta(days=i) for i in range(n + 1)] if n >= 0 else []


def record(
    out_dir: Path | str,
    days: Iterable[date],
    *,
    run_day: date | None = None,
    rows_by_date: dict[str, int] | None = None,
) -> Path:
    """登记这次覆盖到的下单日。同一天重复跑 = 覆盖更新，不重复累计。"""
    data = load(out_dir)
    now = datetime.now().isoformat(timespec="seconds")
    rows_by_date = rows_by_date or {}
    for d in days:
        key = d.isoformat()
        data["covered"][key] = {
            "last_run_at": now,
            "run_day": (run_day or date.today()).isoformat(),
            "rows": rows_by_date.get(key),
        }
    p = ledger_path(out_dir)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return p


def find_gaps(
    out_dir: Path | str,
    through: date,
    *,
    include_weekend: bool = False,
) -> dict:
    """
    从台账里**最早统计过的那天**到 `through` 之间，没被统计过的下单日。

    默认不报周末（销售周末基本不下单，报了全是噪音）；节假日**照报**——
    那正是我们要抓的（程序不认节假日，只能靠这里兜住）。
    """
    covered = covered_dates(out_dir)
    if not covered:
        return {"first_use": True, "gaps": [], "skipped_weekend": [], "through": through}
    start = min(covered)
    gaps: list[date] = []
    skipped: list[date] = []
    for d in daterange(start, through):
        if d in covered:
            continue
        if d.weekday() >= 5 and not include_weekend:
            skipped.append(d)
        else:
            gaps.append(d)
    return {
        "first_use": False,
        "gaps": gaps,
        "skipped_weekend": skipped,
        "through": through,
    }


def gap_note(gaps: list[date]) -> str:
    """给产物页用的一句人话。"""
    if not gaps:
        return ""
    short = "、".join(d.isoformat() for d in gaps[:8]) + ("…" if len(gaps) > 8 else "")
    return (
        f"⚠ 有 {len(gaps)} 个工作日从来没统计过：{short}。"
        "多半是法定假期/机器没跑——本程序只认周末、不认节假日。"
        "补法见《处理日志》「怎么补」一行，⛔ 别一天一天补跑（会重复算）。"
    )
