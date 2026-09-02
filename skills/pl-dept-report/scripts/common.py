#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import re
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
CONFIG = SKILL / "config"

MONEY_Q = Decimal("0.01")
CHECK_TOL = Decimal("0.05")


def load_json(name: str) -> dict:
    path = CONFIG / name
    return json.loads(path.read_text(encoding="utf-8"))


def default_desktop_dir(prefix: str, today: date | None = None, home: Path | None = None) -> Path:
    day = (today or date.today()).strftime("%Y%m%d")
    root = Path(home) if home else Path.home()
    desktop = root / "Desktop"
    base = desktop if desktop.is_dir() else Path.cwd()
    path = base / f"{prefix}_{day}"
    path.mkdir(parents=True, exist_ok=True)
    return path


def default_period(today: date | None = None) -> str:
    d = today or date.today()
    prev = d.replace(day=1) - timedelta(days=1)
    return prev.strftime("%Y%m")


def prev_period(period: str) -> str:
    y, m = int(period[:4]), int(period[4:6])
    if m == 1:
        return f"{y - 1}12"
    return f"{y}{m - 1:02d}"


def detect_period_text(text: str) -> str | None:
    blob = str(text or "")
    m = re.search(r"期间[:：]\s*(\d{6})", blob)
    if m:
        return m.group(1)
    m2 = re.search(r"(20\d{2})年\s*0?(\d{1,2})\s*期", blob)
    if m2:
        month = int(m2.group(2))
        if 1 <= month <= 12:
            return f"{m2.group(1)}{month:02d}"
    m3 = re.search(r"月度损益表_(\d{6})", blob)
    if m3:
        return m3.group(1)
    m4 = re.search(r"(20\d{2})年\s*0?(\d{1,2})\s*月", blob)
    if m4:
        month = int(m4.group(2))
        if 1 <= month <= 12:
            return f"{m4.group(1)}{month:02d}"
    m5 = re.search(r"(20\d{2})[-/](0?\d{1,2})(?!\d)", blob)
    if m5:
        month = int(m5.group(2))
        if 1 <= month <= 12:
            return f"{m5.group(1)}{month:02d}"
    return None


def discover_input_dir(explicit: str = "") -> Path:
    if explicit:
        return Path(explicit).expanduser()
    root = SKILL
    homes: list[Path] = []
    for _ in range(6):
        cand = root / "技能" / "金蝶" / "损益表利润表" / "工作区" / "引出"
        if cand.is_dir():
            homes.append(cand)
            break
        root = root.parent
    desktop = Path.home() / "Desktop"
    if desktop.is_dir():
        homes.extend(sorted(desktop.glob("月度损益表_*"), reverse=True))
        homes.append(desktop)
    homes.append(Path.cwd())

    def _has_source(folder: Path) -> bool:
        if not folder.is_dir():
            return False
        for path in folder.glob("*.xlsx"):
            name = path.name
            if name.startswith("~$") or name.startswith("月度损益表_"):
                continue
            return True
        return False

    for folder in homes:
        if _has_source(folder):
            return folder
    return Path.cwd()


def parse_period(raw: str | None) -> str:
    s = str(raw or "").strip().replace("-", "").replace("/", "")
    if len(s) >= 6 and s[:6].isdigit():
        return s[:6]
    return default_period()


def col_idx(letter: str) -> int:
    n = 0
    for ch in letter.strip().upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(idx: int) -> str:
    out = []
    n = idx
    while n:
        n, rem = divmod(n - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def money(value) -> Decimal | None:
    if value is None or str(value).strip() in {"", "-", "—", "None"}:
        return None
    try:
        return Decimal(str(value).replace(",", "")).quantize(MONEY_Q, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def add_money(left: Decimal | None, right: Decimal | None) -> Decimal | None:
    if left is None and right is None:
        return None
    return (left or Decimal("0.00")) + (right or Decimal("0.00"))


def cell_num(value: Decimal | None):
    if value is None:
        return None
    return float(value)


def local_json_path(env_name: str, default: Path) -> Path:
    raw = os.environ.get(env_name, "").strip()
    return Path(raw) if raw else default


def load_optional_json(path: Path) -> dict | None:
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


_AMOUNT_RE = re.compile(r"(?<!\d)(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}(?!\d)")


def stdout_safe(text: str, secrets: list[str] | None = None) -> str:
    out = _AMOUNT_RE.sub("[金额已省略]", text)
    for secret in secrets or []:
        if secret and secret in out:
            out = out.replace(secret, "[secret]")
    return out
