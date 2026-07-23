#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第6步：核销判定（单入口 · SOD 级 · 三栏分流）v2

2026-07-23 重写。旧版按「回款类型」把每笔到账路由到三条取数通道（对账表 / 预存明细 /
分笔），实测出了两个硬伤，都在这一版根除：

  1. **静默丢单**：预收类既被踢出对账通道、又在明细子表拿不到行，两头落空还不报错
     —— 2026-07-22 真跑丢了 3 笔到账 6.5 万，产出里一个字都没有。
     → 本版取消通道路由；每笔到账都走同一条路，且**强制 AR 覆盖率校验**，
       一笔产不出任何判定就整体报错退出，绝不静默。
  2. **把 SOD 拆行误判成"歧义"**：一个 SO 在她盈亏表占 N 行，其实是 N 个 SOD、
     每行一个金额，合计 = 本次核销额。旧版判 E8「同SO多行歧义」挂起（27 笔里 26 笔是误判）。
     → 本版引入「订单明细」表拿到逐 SOD 交付额，按 SOD 精确落行。

判定口径（明妹 2026-07-23 当面确认 + 接口实调验证）：
  · 逐 SO 本次核销额 H：有「同币种核销明细」就用它（按核销日期过滤本次）；
    **该表 0 行 = 全额核销** → H[so] = 该 SO 交付额。
  · ΣH ≤ 到账额（她的铁律：销的钱不能比到的钱多）。
  · 每个 SO 展开到 SOD：H[so] == Σ该SO的SOD交付额 → 全部 SOD 入选；
    否则在 SOD 里凑唯一子集；凑不出唯一解才挂起。
  · 盈亏表定位：(新智云单号=SO, 应收金额=该SOD交付额) → 唯一行。

红线：分笔一律 hold（禁止比例分摊）；永不写智云；本脚本不写用户任何表。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import common  # noqa: E402

TOL = 0.005  # 金额比较容差（分以下）
SUBSET_MAX_LINES = 22  # 超过这么多 SOD 就不硬凑子集，直接交人

HUIKUAN_NAMES = {
    "ar": ["回款记录ID", "回款记录编号"],
    "currency": ["原币币种", "币种"],
    "arrival_date": ["到账日期", "回款日期"],
    "amount_orig": ["到账金额/原币", "到账金额原币"],
    "amount_local": ["到账金额/本币", "到账金额本币"],
    "fee": ["手续费/原币", "手续费"],
    "huikuan_type": ["回款类型"],
    "status": ["核销状态"],
    "customer": ["开票客户", "客户名称", "客户"],
    "hexiao_date": ["核销日期"],
}


class InputError(Exception):
    """输入缺件/结构不对 —— 脚本非 0 退出，绝不带病继续。"""


class CoverageError(Exception):
    """AR 覆盖率校验没过 —— 有到账没产出任何判定，说明逻辑漏了单。"""


# ══════════════════════════════════════════════════════════════
# 一、读取 01_智云导出 的四张表 → payments
# ══════════════════════════════════════════════════════════════
def _sheet_rows(path: Path) -> Tuple[List[str], List[list]]:
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    headers = [str(c).strip() if c is not None else "" for c in rows[0]]
    body = [list(r) for r in rows[1:] if any(x is not None and str(x).strip() for x in r)]
    return headers, body


def _col(headers: Sequence[str], role: str, key: str, aliases: dict) -> Optional[int]:
    cands = (aliases.get(role) or {}).get(key, [key])
    return common.fuzzy_find_col(headers, cands)


def _need(headers: Sequence[str], role: str, keys: Sequence[str], aliases: dict) -> Dict[str, int]:
    out = {}
    missing = []
    for k in keys:
        i = _col(headers, role, k, aliases)
        if i is None:
            missing.append(k)
        else:
            out[k] = i
    if missing:
        raise InputError(
            f"「{role}」缺列 {missing}。实际表头：{[h for h in headers if h]}"
        )
    return out


def _get(vals: list, idx: Optional[int]) -> Any:
    if idx is None or idx >= len(vals):
        return None
    return vals[idx]


def load_exports(workspace: Path) -> List[dict]:
    """
    01_智云导出/ 四张表 → payments。缺件直接 InputError（不凑合、不猜）。

    文件按名字认：含「回款记录」/「订单交付」/「核销明细」/「订单明细」。
    """
    d = workspace / "01_智云导出"
    if not d.is_dir():
        raise InputError(f"没有 {d}")
    aliases = common.load_aliases()

    def pick(*keys: str) -> Optional[Path]:
        for p in sorted(d.glob("*.xlsx")):
            if p.name.startswith("~$"):
                continue
            if any(k in p.name for k in keys):
                return p
        return None

    p_hk = pick("回款记录")
    p_xd = pick("订单交付", "下单")
    p_mx = pick("核销明细")
    p_sod = pick("订单明细")

    if not p_hk:
        raise InputError(
            "01_智云导出/ 里没有「回款记录」表。\n"
            "  这一版取数只认单入口：先跑 fetch_zhiyun.py（回款记录按核销日期=T-1 + 关联子表），\n"
            "  它会一次产出 回款记录 / 订单交付 / 核销明细 / 订单明细 四份。"
        )
    if not p_xd:
        raise InputError(
            "01_智云导出/ 里没有「订单交付」表（每笔到账关联了哪几个 SO + 交付额）。\n"
            "  旧版的「回款核销对账」已废弃且不兼容——它拿不到 SOD、还会漏掉预收类。\n"
            "  请重新跑 fetch_zhiyun.py 取一次数。"
        )

    # ① 回款记录
    h, body = _sheet_rows(p_hk)
    c = _need(h, "回款记录", ["AR", "核销日期"], aliases)
    for k in ["到账日期", "到账金额原币", "到账金额本币", "手续费", "原币币种",
              "回款类型", "核销状态", "开票客户"]:
        i = _col(h, "回款记录", k, aliases)
        if i is not None:
            c[k] = i
    payments: List[dict] = []
    for vals in body:
        ar = str(_get(vals, c["AR"]) or "").strip()
        if not ar:
            continue
        payments.append({
            "ar": ar,
            "hexiao_date": common.norm_date(_get(vals, c["核销日期"])),
            "arrival_date": common.norm_date(_get(vals, c.get("到账日期"))),
            "amount_orig": common.to_number(_get(vals, c.get("到账金额原币"))),
            "amount_local": common.to_number(_get(vals, c.get("到账金额本币"))),
            "fee": common.to_number(_get(vals, c.get("手续费"))) or 0.0,
            "currency": str(_get(vals, c.get("原币币种")) or "").strip() or "人民币CNY",
            "huikuan_type": str(_get(vals, c.get("回款类型")) or "").strip(),
            "status": str(_get(vals, c.get("核销状态")) or "").strip(),
            "customer": str(_get(vals, c.get("开票客户")) or "").strip(),
            "orders": [],
            "writeoffs": {},
        })
    if not payments:
        raise InputError(f"「回款记录」里一行数据都没有：{p_hk.name}")
    by_ar = {p["ar"]: p for p in payments}

    # ② 订单交付（下单栏）
    h, body = _sheet_rows(p_xd)
    c = _need(h, "订单交付", ["AR", "SO", "交付额原币"], aliases)
    i_rate = _col(h, "订单交付", "汇率", aliases)
    i_cur = _col(h, "订单交付", "币种", aliases)
    i_name = _col(h, "订单交付", "订单名称", aliases)
    for vals in body:
        ar = str(_get(vals, c["AR"]) or "").strip()
        so = str(_get(vals, c["SO"]) or "").strip()
        if not ar or not so or ar not in by_ar:
            continue
        by_ar[ar]["orders"].append({
            "so": so,
            "deliver": common.to_number(_get(vals, c["交付额原币"])),
            "rate": common.to_number(_get(vals, i_rate)),
            "currency": str(_get(vals, i_cur) or "").strip(),
            "name": str(_get(vals, i_name) or "").strip(),
        })

    # ③ 同币种核销明细（累积表：**必须按本次核销日期过滤**）
    if p_mx:
        h, body = _sheet_rows(p_mx)
        c = _need(h, "核销明细", ["回款记录NUM", "核销日期", "本次核销金额"], aliases)
        i_so = _col(h, "核销明细", "SO", aliases)
        for vals in body:
            ar = str(_get(vals, c["回款记录NUM"]) or "").strip()
            if ar not in by_ar:
                continue
            hd = common.norm_date(_get(vals, c["核销日期"]))
            if hd is not None and by_ar[ar]["hexiao_date"] is not None and hd != by_ar[ar]["hexiao_date"]:
                continue  # 上几次的核销，不能重复回填
            so = str(_get(vals, i_so) or "").strip()
            amt = common.to_number(_get(vals, c["本次核销金额"]))
            if not so or amt is None:
                continue
            w = by_ar[ar]["writeoffs"]
            w[so] = round(w.get(so, 0.0) + float(amt), 2)

    # ④ 订单明细（SO → SOD + 逐 SOD 交付额）
    sod_lines: Dict[str, List[dict]] = {}
    if p_sod:
        h, body = _sheet_rows(p_sod)
        c = _need(h, "订单明细", ["SO", "SOD", "交付额原币"], aliases)
        seen = set()
        for vals in body:
            so = str(_get(vals, c["SO"]) or "").strip()
            sod = str(_get(vals, c["SOD"]) or "").strip()
            amt = common.to_number(_get(vals, c["交付额原币"]))
            if not so or not sod or (so, sod) in seen:
                continue
            seen.add((so, sod))
            sod_lines.setdefault(so, []).append({"sod": sod, "deliver": amt})
    for p in payments:
        p["sod_lines"] = sod_lines
    return payments


# ══════════════════════════════════════════════════════════════
# 二、payments → SOD 级 records（核心展开）
# ══════════════════════════════════════════════════════════════
def subset_sum_unique(
    lines: List[dict], target: float, max_lines: int = SUBSET_MAX_LINES
) -> Optional[List[dict]]:
    """
    在 SOD 行里凑出金额恰好 = target 的**唯一**子集。
    多解 / 无解 / 行太多 → None（交给人，绝不随便挑一个）。
    """
    usable = [x for x in lines if x.get("deliver") is not None]
    if not usable or len(usable) > max_lines:
        return None
    cents = [int(round(float(x["deliver"]) * 100)) for x in usable]
    tgt = int(round(target * 100))
    if tgt <= 0:
        return None
    # ways[sum] = (方案数(封顶2), 一个见证掩码)
    ways: Dict[int, Tuple[int, int]] = {0: (1, 0)}
    for i, v in enumerate(cents):
        if v <= 0:
            continue
        nxt = dict(ways)
        for s, (n, mask) in ways.items():
            s2 = s + v
            if s2 > tgt:
                continue
            n0, m0 = nxt.get(s2, (0, 0))
            nxt[s2] = (min(n0 + n, 2), m0 if n0 else (mask | (1 << i)))
        ways = nxt
    n, mask = ways.get(tgt, (0, 0))
    if n != 1:
        return None
    return [usable[i] for i in range(len(usable)) if mask & (1 << i)]


def _payment_local(p: dict, rates: Dict[str, float]) -> Tuple[Optional[float], Optional[str]]:
    """到账本币：系统给了就用系统的，没给才按汇率算。"""
    if p.get("amount_local") is not None:
        return float(p["amount_local"]), None
    return common.compute_local_amount(p.get("amount_orig"), p.get("currency") or "", rates)


def _hold(p: dict, code: str, reason: str, so: str = "", sod: str = "", **extra) -> dict:
    rec = {
        "ar": p["ar"], "so": so, "sod": sod,
        "customer": p.get("customer") or "",
        "amount_orig": None,
        "currency": p.get("currency") or "人民币CNY",
        "hexiao_date": p.get("hexiao_date"),
        "shoukuan_date": p.get("arrival_date"),
        "status": p.get("status") or "",
        "huikuan_type": p.get("huikuan_type") or "",
        "fee": p.get("fee") or 0.0,
        "arrival_total": p.get("amount_orig"),
        "forced_code": code,
        "forced_reason": reason,
    }
    rec.update(extra)
    return rec


def expand_payment(p: dict, rates: Dict[str, float]) -> List[dict]:
    """一笔到账 → 若干 SOD 级 record。**任何情况下至少产出一条**（不许静默丢）。"""
    sod_lines: Dict[str, List[dict]] = p.get("sod_lines") or {}
    orders = p.get("orders") or []
    arrival, err = _payment_local(p, rates)

    # 分笔回款：系统没有逐单金额，铁律一律 hold（禁止比例分摊）
    if "分笔" in (p.get("huikuan_type") or ""):
        return [_hold(p, "E1", "分笔回款：系统无逐单金额，禁止比例分摊")]

    # 手续费口径未拍板
    if float(p.get("fee") or 0.0) > TOL:
        return [_hold(p, "E_FEE", "有手续费，净额/含费口径未拍板")]

    if err == "E6" or arrival is None:
        return [_hold(p, "E6", f"外币缺汇率（{p.get('currency')}），到账本币算不出")]

    if not orders:
        return [_hold(
            p, "E7",
            "这笔到账在智云没关联任何下单（下单栏为空）——先去智云看看这笔回款建对了没",
        )]

    # 逐 SO 本次核销额：核销明细优先；该表无行 = 全额核销
    writeoffs = dict(p.get("writeoffs") or {})
    if writeoffs:
        H = writeoffs
        basis = "核销明细"
    else:
        H = {}
        for o in orders:
            if o.get("deliver") is None:
                return [_hold(p, "E7", f"下单 {o.get('so')} 没有交付额，判不了")]
            H[o["so"]] = round(H.get(o["so"], 0.0) + float(o["deliver"]), 2)
        basis = "全额核销(无核销明细)"

    total_h = round(sum(H.values()), 2)
    cap = arrival + max(float(p.get("fee") or 0.0), 0.0)

    if total_h > cap + TOL:
        if basis == "核销明细":
            return [_hold(
                p, "E4",
                f"超额核销：本次核销合计 {total_h:.2f} > 到账 {arrival:.2f}，"
                "智云数据可能有问题，先找销售核对",
            )]
        return [_hold(
            p, "E5",
            f"这单没回满：交付合计 {total_h:.2f}，实际只到账 {arrival:.2f}，"
            f"差 {total_h - arrival:.2f}。按你的做法要在这单上方插一行、"
            "把应收金额拆成「已收」和「未收」两行，未收那行结账填「否」。"
            "拆完再说「重扫挂账」。",
        )]

    out: List[dict] = []
    deliver_by_so = {}
    for o in orders:
        if o.get("deliver") is not None:
            deliver_by_so[o["so"]] = round(deliver_by_so.get(o["so"], 0.0) + float(o["deliver"]), 2)

    for so, h in H.items():
        lines = sod_lines.get(so) or []
        chosen: Optional[List[dict]] = None
        how = ""
        if lines and all(x.get("deliver") is not None for x in lines):
            total_lines = round(sum(float(x["deliver"]) for x in lines), 2)
            if abs(total_lines - h) <= TOL:
                chosen, how = lines, "全部SOD"
            else:
                chosen = subset_sum_unique(lines, h)
                how = "SOD子集" if chosen else ""
        if chosen is None:
            if lines:
                cand = "、".join(
                    f"{x['sod']}={x['deliver']}" for x in lines[:12]
                ) + ("…" if len(lines) > 12 else "")
                out.append(_hold(
                    p, "E5",
                    f"{so} 本次核销 {h:.2f}，但它下面的 SOD 金额凑不出唯一组合"
                    f"（SOD 合计 {round(sum(float(x['deliver']) for x in lines), 2)}）。"
                    f"候选：{cand}。你指一下这次核的是哪几个 SOD。",
                    so=so,
                ))
            else:
                # 订单明细查不到 SOD → 退化成按 SO 匹配盈亏表（老路，仍可判）
                out.append({
                    "ar": p["ar"], "so": so, "sod": "",
                    "customer": p.get("customer") or "",
                    "amount_orig": h,
                    "currency": p.get("currency") or "人民币CNY",
                    "hexiao_date": p.get("hexiao_date"),
                    "shoukuan_date": p.get("arrival_date"),
                    "status": p.get("status") or "",
                    "huikuan_type": p.get("huikuan_type") or "",
                    "fee": p.get("fee") or 0.0,
                    "arrival_total": p.get("amount_orig"),
                    "deliver_local": deliver_by_so.get(so),
                    "match_basis": f"{basis}/无SOD(按SO匹配)",
                })
            continue
        for line in chosen:
            out.append({
                "so_all_lines": lines,  # 整段对齐消歧要用（见 LedgerIndex.positional_row）
                "ar": p["ar"], "so": so, "sod": line["sod"],
                "customer": p.get("customer") or "",
                "amount_orig": float(line["deliver"]),
                "currency": p.get("currency") or "人民币CNY",
                "hexiao_date": p.get("hexiao_date"),
                "shoukuan_date": p.get("arrival_date"),
                "status": p.get("status") or "",
                "huikuan_type": p.get("huikuan_type") or "",
                "fee": p.get("fee") or 0.0,
                "arrival_total": p.get("amount_orig"),
                "deliver_local": float(line["deliver"]),
                "match_basis": f"{basis}/{how}",
            })

    if not out:  # 兜底：绝不静默丢单
        out.append(_hold(p, "E7", "这笔到账展开不出任何订单行（请把这条截图发给明昊）"))
    return out


def expand_payments(payments: List[dict], rates: Optional[Dict[str, float]] = None) -> List[dict]:
    """全部到账 → records，并做 **AR 覆盖率硬校验**。"""
    rates = rates or {}
    records: List[dict] = []
    for p in payments:
        records.extend(expand_payment(p, rates))
    want = {p["ar"] for p in payments if p.get("ar")}
    got = {r.get("ar") for r in records if r.get("ar")}
    missing = sorted(want - got)
    if missing:
        raise CoverageError(
            "有到账没产出任何判定（这就是 2026-07-22 静默丢 3 笔的那类 bug，绝不放行）："
            f"{missing}"
        )
    return records


# ══════════════════════════════════════════════════════════════
# 三、盈亏表索引
# ══════════════════════════════════════════════════════════════
class LedgerIndex:
    """盈亏表『明细』只读索引。她表里**一行 = 一个 SOD**。"""

    def __init__(self, path: Optional[Path] = None, synthetic: Optional[dict] = None):
        self.path = path
        self.so_index: Dict[str, List[int]] = {}
        self.sod_index: Dict[str, List[int]] = {}
        self.so_amount_index: Dict[Tuple[str, int], List[int]] = {}
        self.row_snapshot: Dict[int, dict] = {}
        self.cols: dict = {}
        if synthetic is not None:
            self.so_index = {k: list(v) for k, v in synthetic.get("so", {}).items()}
            self.sod_index = {k: list(v) for k, v in synthetic.get("sod", {}).items()}
            self.row_snapshot = synthetic.get("rows", {})
            for r, snap in self.row_snapshot.items():
                y = common.to_number(snap.get("yingshou"))
                so = snap.get("so") or ""
                if so and y is not None:
                    self.so_amount_index.setdefault((so, int(round(y * 100))), []).append(int(r))
            return
        if path is not None:
            self._load(path)

    def _load(self, path: Path):
        import openpyxl

        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        if "明细" not in wb.sheetnames:
            names = list(wb.sheetnames)
            wb.close()
            raise ValueError(f"盈亏表无『明细』sheet：{names}")
        all_rows = list(wb["明细"].iter_rows(values_only=True))
        wb.close()

        aliases = common.load_aliases()
        hrow, headers = common.find_header_row(
            all_rows, "盈亏明细", ["SO", "SOD", "计提", "回款明细", "是否结账"], aliases
        )
        cols = common.resolve_columns(
            headers, "盈亏明细",
            ["SO", "SOD", "计提", "回款明细", "是否结账", "收款时间", "收款方式"],
            aliases,
        )
        role = aliases.get("盈亏明细", {})
        yidx = common.fuzzy_find_col(headers, role.get("应收", ["应收金额", "应收"]))
        if yidx is not None:
            cols["应收"] = yidx
        self.cols = cols

        for r_i, row in enumerate(all_rows, start=1):
            if r_i <= hrow + 1:
                continue
            vals = list(row)

            def cell(key):
                i = cols.get(key)
                return vals[i] if i is not None and i < len(vals) else None

            so_s = str(cell("SO") or "").strip()
            sod_s = str(cell("SOD") or "").strip()
            if not so_s.startswith("SO") and not sod_s.startswith("SOD"):
                continue
            if so_s.startswith("SO"):
                self.so_index.setdefault(so_s, []).append(r_i)
            if sod_s.startswith("SOD"):
                self.sod_index.setdefault(sod_s, []).append(r_i)
            yingshou = common.to_number(cell("应收"))
            if so_s.startswith("SO") and yingshou is not None:
                self.so_amount_index.setdefault(
                    (so_s, int(round(yingshou * 100))), []
                ).append(r_i)
            self.row_snapshot[r_i] = {
                "so": so_s, "sod": sod_s,
                "jiti": cell("计提"), "huikuan": cell("回款明细"),
                "jiezhang": cell("是否结账"), "shoukuan_time": cell("收款时间"),
                "shoukuan_way": cell("收款方式"), "yingshou": yingshou,
            }

    def positional_row(
        self, so: str, sod: str, lines: List[dict]
    ) -> Optional[Tuple[int, str, Optional[float]]]:
        """
        (SO+应收金额) 定位不到唯一行时的严格消歧：验证**整段能不能对齐**。

          她表里这个 SO 的全部行（行号升序） vs 智云该 SO 的全部 SOD（编号降序）

        为什么按这个序：两边都源自月初同一份「交付数据」导出，天然同序。

        返回 (行号, 依据, 比例)；对不齐返回 None（老老实实挂起）。依据两种：

        · `exact`  —— 金额序列逐位完全相等。实测 SO26040297（12 行）、SO26040481（10 行）。
        · `ratio`  —— 有几位不等，但**所有不等的位比值完全一致**（系统性口径差，不是行错位）。
                      实测 SO26040322：智云 477.61/661.31 vs 她表 488.64/676.58，
                      两处比值都是 0.977433，她当天就是按智云金额填的。
                      行错位不可能凑出同一个比值，所以这条判据是可证的，不是猜。

        只要行数对不上（比如她把某个 SOD 拆成了两行）→ 直接 None。
        """
        rows = sorted(self.so_index.get(so, []))
        if not rows or not lines or len(rows) != len(lines):
            return None
        ordered = sorted(lines, key=lambda x: str(x.get("sod") or ""), reverse=True)
        ratios: List[float] = []
        for r, ln in zip(rows, ordered):
            y = common.to_number((self.row_snapshot.get(r) or {}).get("yingshou"))
            d = ln.get("deliver")
            if y is None or d is None or float(y) == 0.0:
                return None
            if abs(float(y) - float(d)) > TOL:
                ratios.append(float(d) / float(y))
        kind, ratio = "exact", None
        if ratios:
            ratio = sum(ratios) / len(ratios)
            if not (0.5 < ratio < 1.5):
                return None
            if any(abs(x - ratio) > 5e-4 for x in ratios):
                return None  # 比值不一致 → 更像行错位，不敢认
            kind = "ratio"
        for r, ln in zip(rows, ordered):
            if (ln.get("sod") or "") == sod:
                return r, kind, ratio
        return None

    def match(
        self, so: str, sod: str, amount: Optional[float] = None
    ) -> Tuple[Optional[int], str, List[int]]:
        """
        返回 (行号, 依据, 候选)。依据优先级：
          ① SO + 应收金额  ← 主键：回填前后都成立（她一行 = 一个 SOD，应收金额=该SOD交付额）
          ② SOD           ← 她已回填过「实收金额」列时可用（幂等重跑走这条）
          ③ SO 唯一行
        """
        if so and amount is not None:
            rows = self.so_amount_index.get((so, int(round(float(amount) * 100))), [])
            if len(rows) == 1:
                return rows[0], "SO+应收金额", rows
            if len(rows) > 1:
                return None, "E8", rows
        if sod:
            rows = self.sod_index.get(sod, [])
            if len(rows) == 1:
                return rows[0], "SOD", rows
            if len(rows) > 1:
                return None, "E8", rows
        if so:
            rows = self.so_index.get(so, [])
            if len(rows) == 1:
                return rows[0], "SO", rows
            if len(rows) > 1:
                return None, "E8", rows
            return None, "E2", []
        return None, "E7", []


# ══════════════════════════════════════════════════════════════
# 四、单条判定
# ══════════════════════════════════════════════════════════════
def classify_one(
    rec: dict,
    ledger: Optional[LedgerIndex],
    rates: Dict[str, float],
    thr: float,
    year_now: int,
) -> dict:
    result = {
        "ar": rec.get("ar") or "",
        "so": rec.get("so") or "",
        "sod": rec.get("sod") or "",
        # case_id = AR × SO × SOD：她表里一行一个 SOD，粒度必须到 SOD 否则台账互相覆盖
        "case_id": "|".join(
            x for x in (rec.get("ar") or "-", rec.get("so") or "-", rec.get("sod") or "") if x
        ),
        "customer_masked": common.mask_customer(rec.get("customer") or ""),
        "flow_hits": rec.get("flow_hits"),
        "flow_locate": rec.get("flow_locate") or "",
        "flow_matched_by": rec.get("flow_matched_by") or "",
        "flow_order_suggest": rec.get("flow_order_suggest") or "",
        "huikuan_type": rec.get("huikuan_type") or "",
        "status": rec.get("status") or "",
        "match_basis": rec.get("match_basis") or "",
        "bucket": "exception",
        "code": "",
        "reason": "",
        "five_cols": {},
        "locate_hint": "",
        "current_values": {},
        "candidates": [],
    }
    if rec.get("so"):
        result["locate_hint"] = (
            f"在盈亏『明细』按「新智云单号」筛选：{rec['so']}"
            + (f"，找应收金额={rec.get('amount_orig')} 那行" if rec.get("amount_orig") is not None else "")
            + "（禁止用行号）"
        )

    # 展开阶段已定性的（分笔/手续费/超额/没回满/无下单…）直接落地
    forced = rec.get("forced_code")
    if forced:
        result["code"] = forced
        result["reason"] = rec.get("forced_reason") or ""
        result["bucket"] = "exception" if forced in ("E4", "E7", "E10", "E12", "E0") else "hold"
        return result

    if (rec.get("status") or "") == "已作废":
        result["code"] = "E7"
        result["reason"] = "核销已作废"
        return result

    amount_orig = rec.get("amount_orig")
    if amount_orig is None:
        result["code"] = "E7"
        result["reason"] = "核销金额为空"
        return result

    local, err = common.compute_local_amount(amount_orig, rec.get("currency") or "", rates)
    if err == "E6" or local is None:
        result["code"] = "E6"
        result["reason"] = f"外币缺汇率（{rec.get('currency')}）"
        return result

    # 流转表三键信号（有做才判）
    flow_hits = rec.get("flow_hits")
    if flow_hits is not None:
        try:
            fh = int(flow_hits)
        except (TypeError, ValueError):
            fh = -1
        if fh == 0:
            result["code"] = "E0"
            result["reason"] = "流转表三键对不到账（0 行）"
            return result
        if fh > 1:
            result["code"] = "E12"
            result["reason"] = "同日同额同名命中多行"
            return result
    if rec.get("customer_archive_failed"):
        result["code"] = "E10"
        result["reason"] = "建档失败/搜不到客户"
        return result

    if ledger is None:
        result["bucket"] = "hold"
        result["code"] = "E2"
        result["reason"] = "未提供盈亏表，无法确认 SO 是否在明细"
        return result

    so, sod = rec.get("so") or "", rec.get("sod") or ""
    row, how, cands = ledger.match(so, sod, amount_orig)

    # 定位不到唯一行 → 用「整段逐位对齐」严格消歧（对不齐就继续挂起）
    align_note = ""
    if how in ("E8", "E2") and sod and rec.get("so_all_lines"):
        aligned = ledger.positional_row(so, sod, rec["so_all_lines"])
        if aligned is not None:
            row, kind, ratio = aligned
            how, cands = "SO整段按SOD序对齐", []
            if kind == "ratio":
                snap0 = ledger.row_snapshot.get(row) or {}
                align_note = (
                    f"⚠ 智云交付额 {amount_orig} 与你表里应收 {snap0.get('yingshou')} 不一致"
                    f"（这个 SO 每一行都差同一个比例 {ratio:.6f}），已按**智云金额**填；"
                    "口径待确认，填之前扫一眼"
                )

    if how == "E8":
        result["bucket"] = "hold"
        result["code"] = "E8"
        result["reason"] = (
            f"盈亏表里有 {len(cands)} 行同时满足 SO={so}"
            + (f" 且应收金额={amount_orig}" if amount_orig is not None else "")
            + "，分不清记哪行，你指一下"
        )
        result["candidates"] = cands
        return result
    if how == "E7":
        result["code"] = "E7"
        result["reason"] = "无单号"
        return result
    if how == "E2" or row is None:
        # 表里真找不到这单，这时才区分「跨年老单」还是「还没交付」
        y = common.year_from_so(sod or so)
        result["bucket"] = "hold"
        if y is not None and y < year_now:
            result["code"] = "E3"
            result["reason"] = f"{y} 年的老单，今年盈亏表里没有这行"
        else:
            result["code"] = "E2"
            result["reason"] = "盈亏表里还没有这张单（多半还没交付进表）"
        return result

    snap = ledger.row_snapshot.get(row, {})

    # 超额：本次本币 > 该行应收（人民币）
    yingshou = common.to_number(snap.get("yingshou"))
    if (
        yingshou is not None
        and common.is_cny(rec.get("currency") or "")
        and float(local) > float(yingshou) + max(thr, TOL)
    ):
        result["code"] = "E4"
        result["reason"] = (
            f"这行应收 {yingshou}，本次要记 {local}，比应收还多，不正常"
        )
        return result

    settled = (rec.get("status") or "") in common.SETTLED
    r_time = common.receipt_time(
        common.norm_date(rec.get("shoukuan_date")), common.norm_date(rec.get("hexiao_date"))
    )
    way = common.pay_way(
        rec.get("status") or "",
        common.norm_date(rec.get("shoukuan_date")),
        common.norm_date(rec.get("hexiao_date")),
    )
    local_f = round(float(local), 2)

    # ── 计提金额的判据（2026-07-23 明妹原话确认）──────────────────────
    # 「一个订单的金额分成多次回款，**只有回款明细金额加起来等于交付金额之后**，
    #   才可以填写计提金额」
    #
    # 注意基准是**智云的交付金额**，不是她表里的应收金额 —— 实测 SO26040322 行2567：
    #   她表应收 488.64 / 回款明细 477.61 / 智云交付 477.61 → 她照样填了计提 477.61。
    # 反例（她留空计提但结账仍填「是」）：SO26020068 回 8729.35 而该 SOD 交付 8946.89；
    #   SO26030424 回 199.78 而该 SOD 交付 402.26。
    # 所以「结账」和「计提」是两个判据：结账看这一行的钱到没到，计提看整单回满没有。
    deliver = common.to_number(rec.get("deliver_local"))
    paid_full = deliver is not None and abs(local_f - float(deliver)) <= max(thr, TOL)

    result["five_cols"] = {
        "计提": local_f if (settled and paid_full) else None,
        "回款明细": local_f,
        "是否结账": "是" if settled else "否",
        "收款时间": r_time.isoformat() if r_time else None,
        "收款方式": way,
        "实收SOD": sod or snap.get("sod") or None,
    }
    result["bucket"] = "auto"
    result["code"] = ""
    result["reason"] = f"{rec.get('match_basis') or '判定'} · 定位={how}" + (
        f"；{align_note}" if align_note else ""
    )
    result["current_values"] = {
        "计提": snap.get("jiti"),
        "回款明细": snap.get("huikuan"),
        "是否结账": str(snap.get("jiezhang") or "").strip() if snap.get("jiezhang") is not None else "",
        "收款时间": str(snap.get("shoukuan_time") or "")[:10],
        "收款方式": snap.get("shoukuan_way"),
        "实收SOD": snap.get("sod"),
    }
    result["ledger_row_ref"] = row
    return result


def classify_records(
    records: List[dict],
    ledger: Optional[LedgerIndex] = None,
    rates: Optional[Dict[str, float]] = None,
) -> dict:
    rates = rates or {}
    thr = common.tail_threshold()
    year_now = common.current_year()
    results = [classify_one(rec, ledger, rates, thr, year_now) for rec in records]

    # 同一行被两条计划命中 → 两条都别自动写（谁对谁错要人定）
    seen: Dict[int, str] = {}
    dup: Dict[int, bool] = {}
    for r in results:
        ref = r.get("ledger_row_ref")
        if r["bucket"] != "auto" or ref is None:
            continue
        if ref in seen:
            dup[ref] = True
        seen[ref] = r.get("case_id") or ""
    for r in results:
        ref = r.get("ledger_row_ref")
        if ref is not None and dup.get(ref):
            r["bucket"] = "hold"
            r["code"] = "E8"
            r["reason"] = f"盈亏表第 {ref} 行被本次两笔同时命中，你指一下各记哪行"
            r["five_cols"] = {}

    auto = [r for r in results if r["bucket"] == "auto"]
    hold = [r for r in results if r["bucket"] == "hold"]
    exc = [r for r in results if r["bucket"] == "exception"]
    return {
        "auto": auto,
        "hold": hold,
        "exception": exc,
        "counts": {
            "auto": len(auto), "hold": len(hold),
            "exception": len(exc), "total": len(results),
        },
        "e_code_dist": _dist(results),
        "ar_summary": build_ar_summary(results),
    }


def build_ar_summary(results: List[dict]) -> List[dict]:
    """按 AR 汇总，派生到账流转表「是否更新应收款」三态（是 / 部分 / 空白）。"""
    try:
        from flow_ledger import derive_flow_status
    except Exception:  # pragma: no cover
        def derive_flow_status(states):
            done = sum(1 for s in states if s == "auto")
            if not states or done == 0:
                return ""
            return "是" if done == len(states) else "部分"

    by_ar: Dict[str, List[dict]] = {}
    for r in results:
        by_ar.setdefault(r.get("ar") or "-", []).append(r)
    out = []
    for ar, items in by_ar.items():
        states = [i["bucket"] for i in items]
        out.append({
            "ar": ar,
            "so_count": len({i.get("so") for i in items if i.get("so")}),
            "行数": len(items),
            "buckets": states,
            "流转表_是否更新应收款_建议": derive_flow_status(states),
            "flow_locate": next((i.get("flow_locate") for i in items if i.get("flow_locate")), ""),
            "待处理SO": sorted({i.get("so") for i in items if i["bucket"] != "auto" and i.get("so")}),
        })
    return out


def _dist(items: List[dict]) -> Dict[str, int]:
    d: Dict[str, int] = {}
    for it in items:
        c = it.get("code") or "OK"
        d[c] = d.get(c, 0) + 1
    return d


# ══════════════════════════════════════════════════════════════
# 五、夹具（单测/离线演练用）
# ══════════════════════════════════════════════════════════════
def payments_from_fixture(fixture: dict) -> List[dict]:
    """
    夹具格式（v2）：
      {"payments": [{ar, hexiao_date, arrival_date, amount_orig, ...,
                     orders:[{so, deliver}], writeoffs:{so: amt}}],
       "sod_lines": {"SO…": [{"sod": "SOD…", "deliver": 1.0}]}}
    """
    if "payments" not in fixture:
        raise InputError("夹具不是 v2 格式（需要 payments 键）")
    sod_lines = fixture.get("sod_lines") or {}
    out = []
    for p in fixture["payments"]:
        q = dict(p)
        q["hexiao_date"] = common.norm_date(p.get("hexiao_date"))
        q["arrival_date"] = common.norm_date(p.get("arrival_date"))
        q["amount_orig"] = common.to_number(p.get("amount_orig"))
        q["amount_local"] = common.to_number(p.get("amount_local"))
        q["fee"] = common.to_number(p.get("fee")) or 0.0
        q.setdefault("currency", "人民币CNY")
        q.setdefault("orders", [])
        q.setdefault("writeoffs", {})
        q["sod_lines"] = sod_lines
        out.append(q)
    return out


def serialize_result(result: dict) -> dict:
    def fix(obj):
        if isinstance(obj, dt.date):
            return obj.isoformat()
        if isinstance(obj, dict):
            return {k: fix(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [fix(x) for x in obj]
        return obj

    return fix(result)


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="核销判定（单入口 · SOD 级 · 三栏）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--fixture", default="", help="离线夹具 JSON（v2: payments/sod_lines）")
    ap.add_argument("--ledger", default="", help="盈亏核算表副本（只读）")
    ap.add_argument("--rate", action="append", default=[], help="外币汇率 美元USD=7.0")
    ap.add_argument("--out", default="", help="判定结果 json 路径")
    ap.add_argument("--flow", default="", help="到账流转表副本（只读）；不给则扫 02_我的表副本/")
    ap.add_argument(
        "--flow-complete", action="store_true",
        help="声明当天所有渠道的流转表都已给全；只有这时才判 E0（对不到账）",
    )
    args = ap.parse_args(argv)

    ws = Path(args.workspace)
    common.ensure_out_dirs(ws)
    rates = common.parse_rate_args(args.rate)

    ledger = None
    if args.ledger:
        ledger = LedgerIndex(Path(args.ledger))
    else:
        cand = sorted((ws / "02_我的表副本").glob("*盈亏*")) if (ws / "02_我的表副本").is_dir() else []
        cand = [p for p in cand if not p.name.startswith("~$")]
        if cand:
            ledger = LedgerIndex(cand[0])
        else:
            print(
                "WARN: 未提供盈亏表（--ledger 或 02_我的表副本/*盈亏*）；"
                "全部单将 hold E2，不会瞎填五列",
                file=sys.stderr,
            )

    try:
        if args.fixture:
            payments = payments_from_fixture(json.loads(Path(args.fixture).read_text(encoding="utf-8")))
        else:
            payments = load_exports(ws)
        records = expand_payments(payments, rates)
    except (InputError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2
    except CoverageError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 3

    import flow_ledger as FL

    flow = (
        FL.FlowLedger.from_paths([Path(args.flow)])
        if args.flow
        else FL.FlowLedger.from_workspace(ws)
    )
    if flow.rows:
        FL.annotate_records(records, flow, complete=args.flow_complete)
        print(f"流转表已接入：{len(flow.rows)} 行，来源 {flow.sources}")
    else:
        print(
            "WARN: 未认出到账流转表（02_我的表副本/）→ 本轮不做三键匹配，"
            "E0/E12 不判、清单不给流转定位",
            file=sys.stderr,
        )

    result = classify_records(records, ledger, rates)
    result["flow_sources"] = flow.sources
    result["payment_count"] = len(payments)
    today = dt.date.today().strftime("%Y%m%d")
    out_path = Path(args.out) if args.out else (ws / "04_产出" / f"判定结果_{today}.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(serialize_result(result), ensure_ascii=False, indent=2), encoding="utf-8"
    )

    c = result["counts"]
    print(
        f"判定完成 到账 {len(payments)} 笔 → 订单行 {c['total']} 条："
        f"可填={c['auto']} 挂账={c['hold']} 异常={c['exception']}"
    )
    print(f"E码分布: {result['e_code_dist']}")
    print(f"结果: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
