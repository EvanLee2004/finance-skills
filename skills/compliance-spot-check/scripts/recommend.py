# -*- coding: utf-8 -*-
"""
合规文件抽查 · recommend.py（财务部 Agent 技能）
================================================================
输入：应收all（+ 可选抽查历史 CSV）→ 本周抽查建议清单（文字版 / 可选 xlsx）。

规则（阈值在 config/抽查规则.md）：
  单位 = 销售+客户+交付月份；金额=销售内相对 + ≥1万兜底；
  账龄≥6（特殊客户相对化）；覆盖在职有资格者每人≥1；未反馈优先；离职跳过。

用法：
  python3 recommend.py --input <应收all.xlsx> [--history <csv>] [--out <清单.txt|xlsx>]
"""
from __future__ import annotations

import os
import re
import sys
import csv
import json
import argparse
import datetime
from collections import defaultdict
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
CONFIG_DIR = os.path.join(SKILL_DIR, "config")
WORK_INPUT = os.path.join(SKILL_DIR, "工作区", "input")
WORK_OUTPUT = os.path.join(SKILL_DIR, "工作区", "output")

try:
    import openpyxl
except ImportError:
    print("✗ 缺 openpyxl → pip install openpyxl")
    sys.exit(1)

# 认 sheet 用（与 split-by-sales 对齐）；真正缺列只检查 REQUIRED_KEYS
HEADER_KEYS = [
    "年度", "销售人员", "客户名称", "新智云单号", "文件名", "应收金额",
    "交付月份", "账龄", "结算阶段", "回款日期", "销售解释", "有无合同",
    "合同分类", "PO单", "客户正式确认", "客户结算周期", "是否按月",
]
REQUIRED_KEYS = ["销售人员", "客户名称", "应收金额", "交付月份", "账龄"]
GM_SUFFIX = "-高美杰"

_DEFAULTS: Dict[str, Any] = {
    "aging_threshold": 6,
    "amount_floor": 10000.0,
    "relative_amount_pct": 0.5,  # 销售内金额分位门槛
    "relative_min_aging": 3,     # 走相对金额时账龄至少这么多月
    "weekly_cap": 20,
    "min_per_sales": 1,
    "unfed_boost": 1.5,
    "include_reason": True,
    "history_mode": "技能自建新格式",  # 仅报告，不改变算法
    "acceptor": "",                   # 仅报告
    "already_checked_policy": "排除",
    "ignore_sales": {"高美杰1"},
    "resigned_sales": {"已离职测"},
    "active_sales": [],
    "special_customers": ["方圆"],
}

Unit = Dict[str, Any]
HistoryRec = Dict[str, str]


def log(m: str) -> None:
    print(m, flush=True)


def norm(s: Any) -> str:
    return re.sub(r"\s+", "", str(s or ""))


def split_names(val: str) -> List[str]:
    return [x.strip() for x in re.split(r"[、,，/]", val or "") if x.strip()]


def normalize_sales_name(name: str) -> str:
    """「X-高美杰」→ X；其余原样。"""
    s = str(name or "").strip()
    if s.endswith(GM_SUFFIX) and len(s) > len(GM_SUFFIX):
        return s[: -len(GM_SUFFIX)]
    return s


# ===== config =====
def _parse_md_table(lines: Sequence[str]) -> List[List[str]]:
    rows, seen = [], False
    for ln in lines:
        s = ln.strip()
        if not s.startswith("|"):
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if cells and all(c and set(c) <= set("-: ") for c in cells):
            seen = True
            continue
        if not seen:
            continue
        rows.append(cells)
    return rows


def _num_from(val: str, default: float) -> float:
    m = re.search(r"-?\d+(?:\.\d+)?", val or "")
    return float(m.group()) if m else default


def _copy_defaults() -> Dict[str, Any]:
    return {
        "aging_threshold": _DEFAULTS["aging_threshold"],
        "amount_floor": _DEFAULTS["amount_floor"],
        "relative_amount_pct": _DEFAULTS["relative_amount_pct"],
        "relative_min_aging": _DEFAULTS["relative_min_aging"],
        "weekly_cap": _DEFAULTS["weekly_cap"],
        "min_per_sales": _DEFAULTS["min_per_sales"],
        "unfed_boost": _DEFAULTS["unfed_boost"],
        "include_reason": _DEFAULTS["include_reason"],
        "history_mode": _DEFAULTS["history_mode"],
        "acceptor": _DEFAULTS["acceptor"],
        "already_checked_policy": _DEFAULTS["already_checked_policy"],
        "ignore_sales": set(_DEFAULTS["ignore_sales"]),
        "resigned_sales": set(_DEFAULTS["resigned_sales"]),
        "active_sales": list(_DEFAULTS["active_sales"]),
        "special_customers": list(_DEFAULTS["special_customers"]),
    }


def load_rules() -> Dict[str, Any]:
    """读 config/抽查规则.md。每个旋钮只从一个 section 读。"""
    cfg = _copy_defaults()
    p = os.path.join(CONFIG_DIR, "抽查规则.md")
    if not os.path.isfile(p):
        return cfg
    try:
        text = open(p, encoding="utf-8").read()
    except Exception as e:
        log(f"⚠ 读 抽查规则.md 失败({e})，用内置默认。")
        return cfg

    for part in re.split(r"(?m)^##\s+", text):
        ls = part.splitlines()
        head = ls[0] if ls else ""
        rows = _parse_md_table(ls)

        # 〇：五问（每周条数 / 特殊客户 / 理由 / 历史模式 / 验收人）
        if "待确认" in head or head.startswith("〇") or "5 个" in head:
            for r in rows:
                if len(r) < 2:
                    continue
                k, v = r[0], r[1]
                if "历史表衔接" in k:
                    cfg["history_mode"] = v
                elif "特殊客户" in k:
                    names = split_names(v)
                    if names:
                        cfg["special_customers"] = names
                elif "每周抽查条数" in k or "每周建议条数" in k:
                    cfg["weekly_cap"] = int(_num_from(v, cfg["weekly_cap"]))
                elif "理由" in k:
                    cfg["include_reason"] = v.strip() not in ("否", "不", "不带", "无", "false", "0")
                elif "验收人" in k:
                    cfg["acceptor"] = v.strip()

        # 一：算法阈值（不含每周条数 / 特殊客户）
        elif "可调阈值" in head or (head.startswith("一") and "阈值" in head):
            for r in rows:
                if len(r) < 2:
                    continue
                k, v = r[0], r[1]
                if "账龄门槛" in k:
                    cfg["aging_threshold"] = int(_num_from(v, cfg["aging_threshold"]))
                elif "金额绝对兜底" in k or "绝对兜底" in k:
                    cfg["amount_floor"] = float(_num_from(v, cfg["amount_floor"]))
                elif "相对金额最低账龄" in k:
                    cfg["relative_min_aging"] = int(_num_from(v, cfg["relative_min_aging"]))
                elif "相对金额" in k:
                    cfg["relative_amount_pct"] = float(_num_from(v, cfg["relative_amount_pct"]))
                elif "每人最少" in k:
                    cfg["min_per_sales"] = int(_num_from(v, cfg["min_per_sales"]))
                elif "未反馈加成" in k:
                    cfg["unfed_boost"] = float(_num_from(v, cfg["unfed_boost"]))
                elif "已抽过" in k:
                    cfg["already_checked_policy"] = v.strip() or cfg["already_checked_policy"]
                elif "忽略销售" in k:
                    names = split_names(v)
                    if names:
                        cfg["ignore_sales"] = set(names)

        elif "离职" in head:
            names = {norm(r[0]) for r in rows if r and r[0] and r[0] not in ("销售名", "销售")}
            if names:
                cfg["resigned_sales"] = names

        elif "在职销售" in head or "覆盖口径" in head:
            names = [r[0].strip() for r in rows if r and r[0].strip() and r[0] not in ("销售名", "销售")]
            cfg["active_sales"] = names

    return cfg


def load_aliases() -> Dict[str, List[str]]:
    default = {
        "销售人员": ["销售人员", "销售", "营销人员"],
        "客户名称": ["客户名称", "客户"],
        "应收金额": ["应收金额", "订单折合本币", "金额"],
        "交付月份": ["交付月份", "项目交付", "销售确认"],
        "账龄": ["账龄(月份）", "账龄(月份)", "账龄", "账龄月份"],
    }
    p = os.path.join(CONFIG_DIR, "列名别名.json")
    if not os.path.isfile(p):
        return default
    try:
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
        aliases = d.get("COLUMN_ALIASES") or d
        out = dict(default)
        for k, v in aliases.items():
            if not k.startswith("_") and isinstance(v, list) and v:
                out[k] = v
        return out
    except Exception as e:
        log(f"⚠ 读 列名别名.json 失败({e})，用内置默认。")
        return default


# ===== 读应收 all =====
def find_data_sheet(wb) -> Optional[str]:
    dated, cands = [], []
    for name in wb.sheetnames:
        if "销售反馈" in name or "建议" in name:
            continue
        ws = wb[name]
        hdr = [norm(ws.cell(1, i).value) for i in range(1, min(ws.max_column, 30) + 1)]
        if sum(1 for k in HEADER_KEYS if any(k in h for h in hdr)) >= 5:
            cands.append((ws.max_row, name))
            if re.fullmatch(r"\d{4}[.年]\d{1,2}[.月]\d{1,2}日?", name.strip()):
                dated.append((ws.max_row, name))
    pool = dated or cands
    if not pool:
        return None
    pool.sort(reverse=True)
    if len(pool) > 1:
        log(f"  注意：多个候选 sheet {[n for _, n in pool]}，取「{pool[0][1]}」")
    return pool[0][1]


def map_required_columns(ws, aliases: Dict[str, List[str]]) -> Dict[str, int]:
    """规范键 → 1-based 列号。别名先命中、再子串（与 split 同级简单度）。"""
    actual = {
        c: norm(ws.cell(1, c).value)
        for c in range(1, ws.max_column + 1)
        if norm(ws.cell(1, c).value)
    }
    mapping: Dict[str, int] = {}
    used: Set[int] = set()
    for key in REQUIRED_KEYS:
        found = None
        for alt in aliases.get(key, [key]):
            a = norm(alt)
            if not a:
                continue
            for c, h in actual.items():
                if c in used:
                    continue
                if a == h or a in h:
                    found = c
                    break
            if found is not None:
                break
        if found is None:
            for c, h in actual.items():
                if c in used:
                    continue
                if key in h:
                    found = c
                    break
        if found is not None:
            mapping[key] = found
            used.add(found)

    missing = [k for k in REQUIRED_KEYS if k not in mapping]
    if missing:
        raise RuntimeError(
            f"应收表缺必要列：{missing}。"
            f"表头需含 销售人员/客户名称/应收金额/交付月份/账龄（可有别名，见 config/列名别名.json）。"
        )
    return mapping


def to_number(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "").replace(" ", "")
    if s in ("", "-", "#N/A", "NA", "nan", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_month(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        n = int(v)
        if 190001 <= n <= 210012:
            return f"{n:06d}"
        s = str(n)
        if len(s) == 6:
            return s
    s = str(v).strip()
    m = re.search(r"(20\d{2})\s*[.\-/年]?\s*(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}{int(m.group(2)):02d}"
    digits = re.sub(r"\D", "", s)
    return digits[:6] if len(digits) >= 6 else s


def read_receivable_rows(ws, aliases: Dict[str, List[str]]) -> List[Dict[str, Any]]:
    mapping = map_required_columns(ws, aliases)
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        def cell(key: str):
            c = mapping[key]
            return raw[c - 1] if c <= len(raw) else None

        sales = str(cell("销售人员") or "").strip()
        cust = str(cell("客户名称") or "").strip()
        amt = to_number(cell("应收金额"))
        month = to_month(cell("交付月份"))
        aging = to_number(cell("账龄"))
        if not any([sales, cust, month, amt is not None]):
            continue
        rows.append({
            "销售": sales,
            "客户": cust,
            "金额": amt if amt is not None else 0.0,
            "交付月份": month,
            "账龄": aging if aging is not None else 0.0,
        })
    return rows


# ===== 历史 =====
def load_history(path: Optional[str]) -> List[HistoryRec]:
    if not path:
        return []
    if not os.path.isfile(path):
        log(f"· ⚠ 抽查历史文件不存在：{path}（按无历史继续跑）")
        return []
    recs: List[HistoryRec] = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
        except Exception:
            dialect = csv.excel
        reader = csv.DictReader(f, dialect=dialect)
        if not reader.fieldnames:
            return []

        def pick(row, *cands):
            for k, v in row.items():
                kn = norm(k)
                for c in cands:
                    if norm(c) == kn or norm(c) in kn:
                        return str(v or "").strip()
            return ""

        for row in reader:
            sales = pick(row, "销售", "营销人员", "销售人员")
            cust = pick(row, "客户", "客户名称")
            month = to_month(pick(row, "交付月份", "月份"))
            status = pick(row, "反馈状态", "反馈")
            date = pick(row, "抽查日期", "日期")
            if not sales and not cust:
                continue
            recs.append({
                "销售": sales, "客户": cust, "交付月份": month,
                "反馈状态": status, "抽查日期": date,
            })
    return recs


def index_history(history: List[HistoryRec]) -> Tuple[Set[str], Set[Tuple[str, str, str]], Set[Tuple[str, str, str]]]:
    """→ (未反馈销售norm集合, 未反馈key集合, 已反馈/已抽key集合)。"""
    unfed_sales: Set[str] = set()
    unfed_keys: Set[Tuple[str, str, str]] = set()
    checked_keys: Set[Tuple[str, str, str]] = set()
    for h in history:
        sn, cn, m = norm(h["销售"]), norm(h["客户"]), h["交付月份"]
        key = (sn, cn, m)
        st = (h.get("反馈状态") or "").strip()
        if "未反馈" in st:
            unfed_sales.add(sn)
            unfed_keys.add(key)
        else:
            checked_keys.add(key)
    return unfed_sales, unfed_keys, checked_keys


# ===== 聚合 / 打分 / 覆盖 =====
def group_units(rows: List[Dict[str, Any]]) -> List[Unit]:
    """按 销售(已归一)+客户+交付月份 聚合。"""
    bag: Dict[Tuple[str, str, str], Unit] = {}
    for r in rows:
        sales = normalize_sales_name(r["销售"])
        key = (sales, r["客户"], r["交付月份"])
        if key not in bag:
            bag[key] = {
                "销售": sales,
                "客户": r["客户"],
                "交付月份": r["交付月份"],
                "金额": 0.0,
                "账龄": 0.0,
                "订单数": 0,
            }
        bag[key]["金额"] += float(r["金额"] or 0)
        bag[key]["账龄"] = max(float(bag[key]["账龄"]), float(r["账龄"] or 0))
        bag[key]["订单数"] += 1
    return list(bag.values())


def _is_special(customer: str, specials: Sequence[str]) -> bool:
    cn = norm(customer)
    for s in specials:
        sn = norm(s)
        if sn and (sn in cn or cn in sn):
            return True
    return False


def _median(vals: List[float]) -> float:
    if not vals:
        return 1.0
    xs = sorted(vals)
    mid = len(xs) // 2
    if len(xs) % 2:
        return max(xs[mid], 1.0)
    return max((xs[mid - 1] + xs[mid]) / 2.0, 1.0)


def _amount_percentile(amts: List[float], amount: float) -> float:
    if not amts:
        return 0.0
    if len(amts) == 1:
        return 1.0
    less = sum(1 for a in amts if a < amount)
    equal = sum(1 for a in amts if a == amount)
    return (less + 0.5 * equal) / len(amts)


def score_units(
    units: List[Unit],
    history: List[HistoryRec],
    cfg: Dict[str, Any],
) -> List[Unit]:
    """过滤离职/忽略/已抽过 → 打分 → 标记资格。资格三条 OR，无隐藏阈值。"""
    resigned = {norm(x) for x in cfg["resigned_sales"]}
    ignore = {norm(x) for x in cfg["ignore_sales"]}
    specials = cfg["special_customers"]
    aging_th = int(cfg["aging_threshold"])
    amount_floor = float(cfg["amount_floor"])
    rel_pct = float(cfg["relative_amount_pct"])
    rel_min_aging = float(cfg["relative_min_aging"])
    unfed_boost = float(cfg["unfed_boost"])
    exclude_checked = "排除" in str(cfg.get("already_checked_policy", "排除"))

    unfed_sales, unfed_keys, checked_keys = index_history(history)

    # 预计算：客户账龄中位、销售金额列表
    ages_by_cust: Dict[str, List[float]] = defaultdict(list)
    amts_by_sales: Dict[str, List[float]] = defaultdict(list)
    for u in units:
        ages_by_cust[norm(u["客户"])].append(float(u["账龄"] or 0))
        amts_by_sales[norm(u["销售"])].append(float(u["金额"] or 0))
    med_by_cust = {k: _median(v) for k, v in ages_by_cust.items()}

    scored: List[Unit] = []
    for u in units:
        sales = u["销售"]
        sn = norm(sales)
        if not sn or not u["客户"] or not u["交付月份"]:
            continue
        if sn in resigned or sn in ignore:
            continue

        key = (sn, norm(u["客户"]), u["交付月份"])
        if exclude_checked and key in checked_keys and key not in unfed_keys:
            continue

        amount = float(u["金额"] or 0)
        aging = float(u["账龄"] or 0)
        is_spec = _is_special(u["客户"], specials)
        reasons: List[str] = []

        # —— 账龄 ——
        if is_spec:
            med = med_by_cust.get(norm(u["客户"]), 1.0)
            aging_ok = aging >= med
            if aging_ok:
                aging_score = 8.0 + min(2.0, aging - med)
                reasons.append(f"特殊客户相对账龄{aging:.0f}月≥中位{med:.0f}")
            else:
                aging_score = max(0.0, aging / med * 3.0)
                reasons.append(f"特殊客户相对账龄{aging:.0f}月(中位{med:.0f})")
        else:
            aging_ok = aging >= aging_th
            if aging_ok:
                aging_score = 8.0 + min(2.0, (aging - aging_th) / 6.0)
                reasons.append(f"账龄{aging:.0f}月≥{aging_th}")
            else:
                aging_score = aging / max(aging_th, 1) * 3.0

        # —— 金额 ——
        pct = _amount_percentile(amts_by_sales.get(sn, []), amount)
        amount_score = pct * 10.0
        if amount >= amount_floor:
            amount_score = max(amount_score, 10.0)
            reasons.append(f"金额{amount:.0f}≥兜底{amount_floor:.0f}")
        else:
            reasons.append(f"销售内金额分位{pct:.0%}")

        # —— 资格：三条 OR，阈值全在 config（无代码内魔法数）——
        relative_ok = pct >= rel_pct and aging >= rel_min_aging
        qualifies = aging_ok or amount >= amount_floor or relative_ok
        if relative_ok and not aging_ok and amount < amount_floor:
            reasons.append(f"销售内相对金额(分位{pct:.0%}·账龄{aging:.0f}≥{rel_min_aging:.0f})")

        base = aging_score + amount_score
        unfed = sn in unfed_sales or key in unfed_keys
        if unfed:
            base *= unfed_boost
            reasons.append("未反馈优先")
        if key in unfed_keys:
            reasons.append("本客户月上次未反馈")

        scored.append({
            "销售": sales,
            "客户": u["客户"],
            "交付月份": u["交付月份"],
            "金额": amount,
            "账龄": aging,
            "订单数": u.get("订单数", 1),
            "分数": base,
            "资格": qualifies,
            "未反馈": unfed,
            "理由": "；".join(reasons),
        })

    scored.sort(key=lambda x: (-x["分数"], x["销售"], x["客户"], x["交付月份"]))
    return scored


def active_sales_list(
    scored: List[Unit],
    units: List[Unit],
    cfg: Dict[str, Any],
) -> List[str]:
    resigned = {norm(x) for x in cfg["resigned_sales"]}
    ignore = {norm(x) for x in cfg["ignore_sales"]}
    if cfg.get("active_sales"):
        return [s for s in cfg["active_sales"] if s and norm(s) not in resigned]
    seen, out = set(), []
    for u in units:
        n = norm(u["销售"])
        if n and n not in seen and n not in resigned and n not in ignore:
            seen.add(n)
            out.append(u["销售"])
    return out


def select_recommendations(
    scored: List[Unit],
    cfg: Dict[str, Any],
    active: Sequence[str],
) -> Tuple[List[Unit], Dict[str, Any]]:
    """只从有「资格」的单位里选：先每人≥min_per，再按分填满 weekly_cap。无资格的销售记入 no_candidate。"""
    weekly_cap = int(cfg["weekly_cap"])
    min_per = int(cfg["min_per_sales"])

    by_sales: Dict[str, List[Unit]] = defaultdict(list)
    for u in scored:
        by_sales[norm(u["销售"])].append(u)

    picked: List[Unit] = []
    picked_keys: Set[Tuple[str, str, str]] = set()
    no_candidate: List[str] = []

    def add(u: Unit) -> bool:
        k = (norm(u["销售"]), norm(u["客户"]), u["交付月份"])
        if k in picked_keys or len(picked) >= weekly_cap:
            return False
        picked.append(u)
        picked_keys.add(k)
        return True

    # 第一轮：每位在职销售，仅从有资格候选里取 top min_per
    for display in active:
        sn = norm(display)
        qual = [c for c in by_sales.get(sn, []) if c.get("资格")]
        if not qual:
            no_candidate.append(display)
            continue
        n = 0
        for c in qual:  # 已按分数排序
            if n >= min_per:
                break
            if add(c):
                n += 1

    # 第二轮：全局有资格、按分填满
    for u in scored:
        if len(picked) >= weekly_cap:
            break
        if u.get("资格"):
            add(u)

    picked.sort(key=lambda x: (-x["分数"], x["销售"], x["客户"], x["交付月份"]))
    return picked, {
        "active_sales": list(active),
        "no_candidate_sales": no_candidate,
        "weekly_cap": weekly_cap,
        "picked": len(picked),
    }


# ===== 输出 =====
def format_text_list(picked: List[Unit], include_reason: bool) -> str:
    headers = ["营销人员", "客户名称", "交付月份"]
    if include_reason:
        headers.append("理由")
    lines = ["\t".join(headers)]
    for u in picked:
        row = [str(u["销售"]), str(u["客户"]), str(u["交付月份"])]
        if include_reason:
            row.append(str(u.get("理由") or ""))
        lines.append("\t".join(row))
    return "\n".join(lines) + ("\n" if lines else "")


def write_xlsx(path: str, picked: List[Unit], include_reason: bool) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "本周抽查建议"
    headers = ["营销人员", "客户名称", "交付月份"]
    if include_reason:
        headers.append("理由")
    ws.append(headers)
    for u in picked:
        row = [u["销售"], u["客户"], u["交付月份"]]
        if include_reason:
            row.append(u.get("理由") or "")
        ws.append(row)
    wb.save(path)


def write_text(path: str, text: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


# ===== 主流程 =====
def recommend(
    input_path: str,
    history_path: Optional[str] = None,
    out_path: Optional[str] = None,
    cfg: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    cfg = cfg or load_rules()
    aliases = load_aliases()

    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    sheet = find_data_sheet(wb)
    if not sheet:
        wb.close()
        raise RuntimeError(
            "没找到数据 sheet（表头需含 销售人员/客户名称/应收金额/交付月份/账龄 等列）。"
        )
    log(f"· 数据 sheet：{sheet}")
    rows = read_receivable_rows(wb[sheet], aliases)
    wb.close()
    log(f"· 读到 {len(rows)} 行订单")

    units = group_units(rows)
    log(f"· 聚合为 {len(units)} 个抽查单位（销售+客户+交付月份）")

    history = load_history(history_path)
    if history_path:
        log(f"· 抽查历史：{len(history)} 条（模式：{cfg.get('history_mode')}）")
    else:
        log("· 未提供抽查历史（可选；按无历史跑）")

    scored = score_units(units, history, cfg)
    active = active_sales_list(scored, units, cfg)
    picked, meta = select_recommendations(scored, cfg, active)
    text = format_text_list(picked, bool(cfg.get("include_reason", True)))

    if not out_path:
        base = os.path.splitext(os.path.basename(input_path))[0]
        out_path = os.path.join(
            os.path.dirname(os.path.abspath(input_path)),
            f"本周抽查建议_{base}.txt",
        )
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    ext = os.path.splitext(out_path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        write_xlsx(out_path, picked, bool(cfg.get("include_reason", True)))
        txt_side = os.path.splitext(out_path)[0] + ".txt"
        write_text(txt_side, text)
        log(f"· 已写 Excel：{out_path}")
        log(f"· 已写文字版：{txt_side}")
    else:
        write_text(out_path, text)
        log(f"· 已写文字版：{out_path}")

    log(f"· 建议 {len(picked)} 条（上限 {meta['weekly_cap']}）；在职口径 {len(meta['active_sales'])} 人")
    if meta["no_candidate_sales"]:
        log(f"· ⚠ 无符合条件候选的销售：{', '.join(meta['no_candidate_sales'])}")
    if cfg.get("acceptor"):
        log(f"· 验收人（配置）：{cfg['acceptor']}")

    return {
        "input_rows": len(rows),
        "units": len(units),
        "scored": len(scored),
        "picked": picked,
        "text": text,
        "out_path": out_path,
        "meta": meta,
        "cfg": cfg,
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="合规文件抽查 · 本周建议清单")
    ap.add_argument("--input", help="应收all.xlsx（绝对路径）")
    ap.add_argument("--history", help="抽查历史 CSV（可选；不存在则按无历史继续）")
    ap.add_argument("--out", help="输出路径 .txt 或 .xlsx（可省：源文件同目录）")
    ap.add_argument("--input-dir", default=WORK_INPUT, help="没给 --input 时去这里找最新 xlsx")
    a = ap.parse_args()

    inp = a.input
    auto_picked = not a.input
    if not inp:
        cands = []
        if os.path.isdir(a.input_dir):
            cands = [
                os.path.join(a.input_dir, f)
                for f in sorted(os.listdir(a.input_dir))
                if f.lower().endswith(".xlsx") and not f.startswith(("~$", "."))
            ]
        if not cands:
            log(f"✗ 没给 --input，也没在 {a.input_dir}/ 找到 xlsx。")
            sys.exit(1)
        inp = max(cands, key=os.path.getmtime)
    if not os.path.isfile(inp):
        log(f"✗ 输入文件不存在：{inp}")
        sys.exit(1)
    if not inp.lower().endswith((".xlsx", ".xlsm")):
        log(f"✗ 输入要是 .xlsx（收到 {os.path.basename(inp)}）。")
        sys.exit(1)

    cfg = load_rules()
    log(f"· 输入：{os.path.basename(inp)}")
    log(
        f"· 规则：账龄≥{cfg['aging_threshold']}月 / 金额兜底{cfg['amount_floor']:.0f} / "
        f"相对分位≥{cfg['relative_amount_pct']:.0%} / 每周上限{cfg['weekly_cap']} / "
        f"理由列{'开' if cfg['include_reason'] else '关'} / "
        f"特殊客户{cfg['special_customers']} / 离职{sorted(cfg['resigned_sales'])}"
    )

    if a.out:
        out = a.out
    elif auto_picked:
        os.makedirs(WORK_OUTPUT, exist_ok=True)
        out = os.path.join(WORK_OUTPUT, f"本周抽查建议_{datetime.date.today().strftime('%Y%m%d')}.txt")
    else:
        out = os.path.join(
            os.path.dirname(os.path.abspath(inp)),
            f"本周抽查建议_{os.path.splitext(os.path.basename(inp))[0]}.txt",
        )

    try:
        rep = recommend(inp, a.history, out, cfg)
    except Exception as e:
        log(f"✗ 推荐出错：{e}")
        sys.exit(1)

    print("\n===== 本周抽查建议清单（可粘贴）=====", flush=True)
    print(rep["text"], end="" if rep["text"].endswith("\n") else "\n", flush=True)
    print("===== 清单结束 =====", flush=True)
    log(f"\n✓ 完成：{len(rep['picked'])} 条建议 → {rep['out_path']}")


if __name__ == "__main__":
    main()
