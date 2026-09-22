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


def find_desktop(root: Path | None = None) -> Path | None:
    """Windows 桌面可能在 OneDrive 下；按常见位置找，找不到返回 None。"""
    root = Path(root) if root else Path.home()
    for rel in ("Desktop", "桌面", "OneDrive/Desktop", "OneDrive/桌面"):
        cand = root / rel
        if cand.is_dir():
            return cand
    return None


def default_desktop_dir(prefix: str, today: date | None = None, home: Path | None = None) -> Path:
    day = (today or date.today()).strftime("%Y%m%d")
    root = Path(home) if home else Path.home()
    desktop = find_desktop(root)
    base = desktop if desktop else Path.cwd()
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


_RANGE = r"[-~～到至]+"


def _yyyy_mm(year: str, month: int) -> str | None:
    if 1 <= month <= 12:
        return f"{year}{month:02d}"
    return None


def detect_period_text(text: str) -> str | None:
    """只认一个会计期间。起止跨月、打印日期、导出时间戳都不算。多个月并存则当没写。"""
    blob = str(text or "")
    if not blob.strip():
        return None
    found: list[str] = []

    for m in re.finditer(rf"(\d{{6}})\s*{_RANGE}\s*(\d{{6}})", blob):
        if m.group(1) != m.group(2):
            return None
        found.append(m.group(1))

    period_re = re.compile(
        rf"(20\d{{2}})年\s*第?\s*0?(\d{{1,2}})\s*期(?:\s*{_RANGE}\s*(20\d{{2}})年\s*第?\s*0?(\d{{1,2}})\s*期)?"
    )
    for m in period_re.finditer(blob):
        left = _yyyy_mm(m.group(1), int(m.group(2)))
        if m.group(3):
            right = _yyyy_mm(m.group(3), int(m.group(4)))
            if not left or not right or left != right:
                return None
            found.append(left)
        elif left:
            found.append(left)

    span_re = re.compile(
        rf"(20\d{{2}})[./-](0?\d{{1,2}})\s*{_RANGE}\s*(20\d{{2}})[./-](0?\d{{1,2}})(?!\d)"
    )
    for m in span_re.finditer(blob):
        left = _yyyy_mm(m.group(1), int(m.group(2)))
        right = _yyyy_mm(m.group(3), int(m.group(4)))
        if not left or not right or left != right:
            return None
        found.append(left)

    for m in re.finditer(r"期间[:：]\s*(\d{6})(?!\d)", blob):
        found.append(m.group(1))
    for m in re.finditer(r"月度损益表_(\d{6})(?!\d)", blob):
        found.append(m.group(1))
    for m in re.finditer(r"(20\d{2})年\s*0?(\d{1,2})\s*月(?!\s*\d)", blob):
        value = _yyyy_mm(m.group(1), int(m.group(2)))
        if value:
            found.append(value)
    for m in re.finditer(r"(20\d{2})[./-](0?\d{1,2})(?!\d)(?![./-]\d)", blob):
        value = _yyyy_mm(m.group(1), int(m.group(2)))
        if value:
            found.append(value)
    for m in re.finditer(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)", blob):
        found.append(f"{m.group(1)}{m.group(2)}")

    uniq = list(dict.fromkeys(found))
    if len(uniq) == 1:
        return uniq[0]
    return None


def clean_header(text: str) -> str:
    """表头统一：去空白、全角括号，去掉金额单位。不改业务含义。"""
    raw = str(text or "")
    t = (
        raw.replace("\xa0", "")
        .replace("\u3000", "")
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .replace("（", "(")
        .replace("）", ")")
    )
    for suffix in ("(人民币元)", "(人民币)", "(元)"):
        if t.endswith(suffix):
            t = t[: -len(suffix)]
    return t.strip()


def _folder_has_source(folder: Path) -> bool:
    if not folder.is_dir():
        return False
    for path in folder.iterdir():
        if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            continue
        if path.name.startswith("~$") or path.name.startswith("月度损益表_"):
            continue
        return True
    return False


def discover_input_dir(explicit: str = "") -> Path:
    if explicit:
        return Path(explicit).expanduser()
    cwd = Path.cwd()
    parts = {p.lower() for p in cwd.parts}
    in_skill_tree = "finance-skills" in parts and "skills" in parts
    desktop = find_desktop()
    cwd_is_desktop = bool(desktop and cwd.resolve() == desktop.resolve())
    if (
        not in_skill_tree
        and not cwd_is_desktop
        and not cwd.name.startswith(("pytest-", "tmp"))
        and _folder_has_source(cwd)
    ):
        return cwd
    return default_desktop_dir("月度损益表")


def parse_period(raw: str | None) -> str:
    """空字符串用上一个已过完的月。写了但认不出，返回空，调用方问人，不要改成别的月。"""
    s = str(raw or "").strip()
    if not s:
        return default_period()
    found = detect_period_text(s)
    if found:
        return found
    compact = s.replace("-", "").replace("/", "").replace(".", "")
    if len(compact) == 6 and compact.isdigit() and compact[4:6] != "00":
        month = int(compact[4:6])
        if 1 <= month <= 12:
            return compact
    return ""


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
    if hasattr(value, "year") and hasattr(value, "month") and not isinstance(value, Decimal):
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
