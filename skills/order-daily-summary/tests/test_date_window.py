"""date_window 表驱动单测：覆盖周一 / 周三 / 周日。"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

_SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(_SCRIPTS))

from date_window import date_window, is_weekend  # noqa: E402


@pytest.mark.parametrize(
    "today, expect_start, expect_end",
    [
        # Monday → last Fri–Sun
        (date(2026, 7, 20), date(2026, 7, 17), date(2026, 7, 19)),
        # Wednesday → yesterday only
        (date(2026, 7, 22), date(2026, 7, 21), date(2026, 7, 21)),
        # Thursday (gold run day 7/23 style: run Thu → Wed)
        (date(2026, 7, 23), date(2026, 7, 22), date(2026, 7, 22)),
        # Tuesday
        (date(2026, 7, 21), date(2026, 7, 20), date(2026, 7, 20)),
        # Friday
        (date(2026, 7, 24), date(2026, 7, 23), date(2026, 7, 23)),
        # Sunday → last Fri through yesterday (Sat)
        (date(2026, 7, 19), date(2026, 7, 17), date(2026, 7, 18)),
        # Saturday → last Fri through yesterday (Fri)
        (date(2026, 7, 18), date(2026, 7, 17), date(2026, 7, 17)),
    ],
)
def test_date_window_table(today, expect_start, expect_end):
    start, end = date_window(today)
    assert start == expect_start
    assert end == expect_end
    assert start <= end


def test_weekend_flags():
    assert is_weekend(date(2026, 7, 18)) is True  # Sat
    assert is_weekend(date(2026, 7, 19)) is True  # Sun
    assert is_weekend(date(2026, 7, 20)) is False  # Mon
