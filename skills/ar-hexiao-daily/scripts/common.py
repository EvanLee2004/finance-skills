#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""共享工具：列名模糊匹配、数值/日期、配置加载。金额计算只在此/脚本内进行。"""
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
CONFIG = SKILL / "config"
WORK = SKILL / "工作区"
OUT_DIR = WORK / "04_产出"
LEDGER_DIR = WORK / "03_台账"

CNY_ALIASES = {"人民币CNY", "人民币", "CNY", "RMB", ""}


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def load_aliases() -> dict:
    p = CONFIG / "列名别名.json"
    return load_json(p) if p.is_file() else {}


def load_codes() -> dict:
    p = CONFIG / "判定码.json"
    return load_json(p) if p.is_file() else {}


def tail_threshold() -> float:
    """从业务规则.md 解析尾差阈值；缺省 0。"""
    p = CONFIG / "业务规则.md"
    if not p.is_file():
        return 0.0
    text = p.read_text(encoding="utf-8")
    m = re.search(r"尾差阈值[^`]*`([0-9.]+)`", text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return 0.0
    return 0.0


def current_year() -> int:
    p = CONFIG / "业务规则.md"
    if p.is_file():
        m = re.search(r"当前年[^\d]*(\d{4})", p.read_text(encoding="utf-8"))
        if m:
            return int(m.group(1))
    return dt.date.today().year


def _norm(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def to_number(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("，", "").replace("¥", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def norm_date(v: Any) -> Optional[dt.date]:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v.strip():
        s = v.strip()[:10].replace("/", "-")
        try:
            return dt.date.fromisoformat(s)
        except ValueError:
            return None
    return None


def fuzzy_find_col(
    headers: Sequence[str],
    aliases: Sequence[str],
) -> Optional[int]:
    """在表头列表中按别名模糊匹配，返回 0-based 列索引。"""
    norms = [_norm(h) for h in headers]
    for alias in aliases:
        a = alias.strip()
        if not a:
            continue
        for i, h in enumerate(norms):
            if not h:
                continue
            if h == a or h.startswith(a) or a in h:
                return i
    return None


def find_header_row(rows, role: str, required, aliases=None, scan: int = 8):
    """
    在前 scan 行里找真正的表头行（她的表常有标题行/空行，盲取第一行会全崩）。
    返回 (行号0基, 表头列表)；找不到则回退第一行，让 resolve_columns 去报错列出实际表头。
    """
    aliases = aliases or load_aliases()
    role_map = aliases.get(role, {})
    best = (-1, None, -1)
    for i, row in enumerate(rows[:scan]):
        headers = [str(x).strip() if x is not None else "" for x in row]
        if not any(headers):
            continue
        hit = 0
        for key in required:
            if fuzzy_find_col(headers, role_map.get(key, [key])) is not None:
                hit += 1
        if hit > best[2]:
            best = (i, headers, hit)
        if hit == len(required):
            return i, headers
    if best[1] is not None:
        return best[0], best[1]
    first = rows[0] if rows else []
    return 0, [str(x).strip() if x is not None else "" for x in first]


def resolve_columns(
    headers: Sequence[str],
    role: str,
    required: Sequence[str],
    aliases: Optional[dict] = None,
) -> Dict[str, int]:
    """
    按 config 列名别名解析列。
    找不到任一 required → raise ValueError，文案含实际表头。
    """
    aliases = aliases or load_aliases()
    role_map = aliases.get(role, {})
    found: Dict[str, int] = {}
    missing = []
    for key in required:
        cands = role_map.get(key, [key])
        idx = fuzzy_find_col(headers, cands)
        if idx is None:
            missing.append(key)
        else:
            found[key] = idx
    if missing:
        actual = [h for h in headers if _norm(h)]
        raise ValueError(
            f"表头缺列：{missing}（角色={role}）。实际表头：{actual}"
        )
    return found


class ColumnError(Exception):
    """列名匹配失败（脚本应非 0 退出）。"""


def is_cny(currency: str) -> bool:
    c = _norm(currency)
    return c in CNY_ALIASES or "人民币" in c or c.upper() in {"CNY", "RMB"}


def year_from_so(so: str) -> Optional[int]:
    """SO26030412 → 2026；SOD2512xx → 2025。"""
    s = _norm(so)
    m = re.search(r"(?:SO|SOD)(\d{2})", s, re.I)
    if not m:
        return None
    yy = int(m.group(1))
    return 2000 + yy


def mask_customer(name: str) -> str:
    s = _norm(name)
    if not s:
        return ""
    if len(s) <= 2:
        return s[0] + "*"
    return s[0] + "*" * (len(s) - 2) + s[-1]


def ensure_out_dirs(workspace=None) -> None:
    """建齐工作区四个子目录。传 workspace 时跟随它，否则用技能自带工作区。"""
    from pathlib import Path as _P

    base = _P(workspace) if workspace else WORK
    for d in ("01_智云导出", "02_我的表副本", "03_台账", "04_产出"):
        (base / d).mkdir(parents=True, exist_ok=True)


def parse_rate_args(rate_items: Optional[Sequence[str]]) -> Dict[str, float]:
    """解析 --rate 美元USD=7.0 / 美元=7.0。"""
    out: Dict[str, float] = {}
    if not rate_items:
        return out
    for item in rate_items:
        if "=" not in item:
            continue
        k, v = item.split("=", 1)
        try:
            out[k.strip()] = float(v.strip())
        except ValueError:
            continue
    return out


def pick_rate(currency: str, rates: Dict[str, float]) -> Optional[float]:
    if is_cny(currency):
        return 1.0
    c = _norm(currency)
    if c in rates:
        return rates[c]
    for k, v in rates.items():
        if k in c or c in k:
            return v
        if "美元" in c and "美元" in k:
            return v
        if "USD" in c.upper() and "USD" in k.upper():
            return v
    return None


def compute_local_amount(
    amount_orig: Optional[float],
    currency: str,
    rates: Dict[str, float],
) -> Tuple[Optional[float], Optional[str]]:
    """返回 (本币金额, 错误码)。外币无汇率 → (None, 'E6')。"""
    if amount_orig is None:
        return None, "E7"
    if is_cny(currency):
        return round(float(amount_orig), 2), None
    r = pick_rate(currency, rates)
    if r is None:
        return None, "E6"
    return round(float(amount_orig) * r, 2), None


def receipt_time(
    shoukuan: Optional[dt.date],
    hexiao: Optional[dt.date],
) -> Optional[dt.date]:
    """R5：同月填到账日，跨月填核销日。"""
    if shoukuan is None and hexiao is None:
        return None
    if shoukuan is None:
        return hexiao
    if hexiao is None:
        return shoukuan
    if shoukuan.year == hexiao.year and shoukuan.month == hexiao.month:
        return shoukuan
    return hexiao


def pay_way(
    status: str,
    huikuan_type: str = "",
    shoukuan=None,
    hexiao=None,
) -> str:
    """
    收款方式。她表里实际只用两个值：「汇」和「冲预收」。

    「冲预收」有两种业务外观（v2 §5 / 规则 R6），**都要判成冲预收**：
      甲 真预收消耗：预存类核销 —— 看核销状态/回款类型；
      乙 晚核销标签：钱早到账、销售下个月才关联订单 —— **到账月 ≠ 核销月**。
        乙这条此前漏了，实测对不上她的表：AR26070033 六月到账七月核销，
        她填「冲预收」我们填「汇」。这标签是给月度提成筛选用的，填错会漏提成。
    """
    prepaid = {"预存已核销", "预存部分核销", "预存待核销"}
    if status in prepaid or "预存" in _norm(huikuan_type):
        return "冲预收"
    if shoukuan is not None and hexiao is not None:
        if (shoukuan.year, shoukuan.month) != (hexiao.year, hexiao.month):
            return "冲预收"
    return "汇"


SETTLED = {"手动核销", "自动核销", "核销成功", "预存已核销"}
PARTIAL = {"预存部分核销", "核销确认中", "销售待核销", "预存待核销", "待提交"}
PREPAID_STATUS = {"预存已核销", "预存部分核销", "预存待核销"}


def plain_cell(v: Any, options: Optional[Dict[str, str]] = None) -> str:
    """明道云单元格 / Excel 值 → 纯文本。"""
    if v is None:
        return ""
    if isinstance(v, (list, tuple)):
        parts = []
        for item in v:
            if isinstance(item, dict):
                parts.append(item.get("name") or item.get("fullname") or "")
            elif options and str(item) in options:
                parts.append(options[str(item)])
            else:
                parts.append(str(item))
        return "、".join(x for x in parts if x)
    s = str(v).strip()
    if s.startswith("["):
        try:
            arr = json.loads(s)
        except (ValueError, json.JSONDecodeError):
            return s
        return plain_cell(arr, options)
    if options and s in options:
        return options[s]
    # GUID list-like single option key
    if options:
        for k, name in options.items():
            if k == s or (len(s) > 8 and k.startswith(s[:8])):
                return name
    return s


def build_field_index(fields: List[dict]) -> Dict[str, dict]:
    """中文名 → {id, options}；同时 id → 自身。"""
    by_name: Dict[str, dict] = {}
    for f in fields:
        fid = f.get("id") or f.get("controlId") or ""
        name = f.get("name") or f.get("controlName") or ""
        opts = {}
        for o in f.get("options") or []:
            opts[str(o.get("key", ""))] = str(o.get("value", ""))
        info = {"id": fid, "name": name, "options": opts}
        if name:
            by_name[name] = info
        if fid:
            by_name[fid] = info
    return by_name


def get_by_names(row: dict, field_index: dict, names: Sequence[str]) -> Any:
    """按候选中文名从原始行取值（不写死 controlId）。"""
    for n in names:
        info = field_index.get(n)
        if not info:
            # 模糊：名字包含
            for k, inf in field_index.items():
                if n in k or k.startswith(n):
                    info = inf
                    break
        if not info:
            continue
        fid = info["id"]
        if fid in row:
            return plain_cell(row.get(fid), info.get("options") or {})
        if n in row:
            return plain_cell(row.get(n), info.get("options") or {})
    return ""
