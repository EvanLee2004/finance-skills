"""日期窗口：按运行日 weekday 决定抓取「下单日期」闭区间 [start, end]。

规则（锁死）：
- 周一 (0)：上周五～上周日
- 周二～周五 (1–4)：昨天～昨天
- 周六/日 (5–6)：上周五～昨天（默认周末不跑，CLI 可 --force）
"""

from __future__ import annotations

from datetime import date, timedelta


def date_window(today: date) -> tuple[date, date]:
    """Return inclusive (start, end) order-date window for a run day."""
    if not isinstance(today, date):
        raise TypeError(f"today 须为 date，收到 {type(today)!r}")
    wd = today.weekday()  # Mon=0 … Sun=6
    if wd == 0:  # Monday
        start = today - timedelta(days=3)  # Fri
        end = today - timedelta(days=1)  # Sun
        return start, end
    if 1 <= wd <= 4:  # Tue–Fri
        y = today - timedelta(days=1)
        return y, y
    # Sat/Sun: last Friday through yesterday
    # Friday is today - (wd - 4): Sat→2d, Sun→3d back to Friday
    start = today - timedelta(days=wd - 4)
    end = today - timedelta(days=1)
    return start, end


def is_weekend(today: date) -> bool:
    return today.weekday() >= 5
