#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第6步：核销判定（三通道取数 + 三键/单号匹配 + 三栏分流）。

输入：
  - 智云导出 Excel（回款记录 / 对账 / 核销明细）或 离线夹具 JSON
  - 可选：盈亏核算表副本（用于 SO/SOD 是否在表、多行歧义）
输出：
  - 工作区/04_产出/判定结果_YYYYMMDD.json
  - 运行报告（只计数与 E 码，无客户名金额明细进 stdout 时可关）

红线：分笔一律 hold；预存明细按核销日期过滤；永不写智云；不写用户表。
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

# 逻辑字段 → 夹具/字段定义里可能的中文名（按中文解析，不写死 controlId）
DUIZHANG_NAMES = {
    "so": ["SO", "销售订单编号", "新智云单号"],
    "sod": ["SOD", "结算订单编号", "结算编号"],
    "ar": ["回款记录ID", "回款记录"],
    "customer": ["客户名称", "开票客户"],
    "amount_orig": ["回款额/原币", "回款额原币"],
    "amount_local": ["回款额/本币", "回款额本币"],
    "deliver_local": ["交付额/本币", "交付额本币"],
    "deliver_orig": ["交付额/原币", "交付额原币"],
    "currency": ["下单币种", "币种", "原币币种"],
    "hexiao_date": ["核销日期"],
    "shoukuan_date": ["回款日期", "到账日期"],
    "status": ["核销状态"],
    "order_name": ["订单名称"],
}
HUIKUAN_NAMES = {
    "ar": ["回款记录ID", "回款记录编号"],
    "currency": ["原币币种", "币种"],
    "arrival_date": ["到账日期", "回款日期"],
    "amount_orig": ["到账金额/原币", "到账金额原币"],
    "fee": ["手续费/原币", "手续费"],
    "huikuan_type": ["回款类型"],
    "status": ["核销状态"],
    "customer": ["开票客户", "客户名称", "客户"],
    "hexiao_date": ["核销日期"],
}
DETAIL_NAMES = {
    "ar": ["回款记录NUM", "回款记录ID", "回款记录"],
    "hexiao_date": ["核销日期"],
    "amount": ["本次核销金额", "核销金额", "金额"],
    "currency": ["币种", "原币币种"],
    "rate": ["汇率"],
    "so": ["SO", "销售订单编号", "新智云单号"],
}


def _channel_for_type(huikuan_type: str) -> str:
    t = (huikuan_type or "").strip()
    if "分笔" in t:
        return "fenbi"
    if "预存" in t or "预收" in t:
        return "yucun"
    # 整笔 / 自动 / 空
    return "duizhang"


def load_fixture(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def parse_duizhang_rows(
    rows: List[dict],
    fields: List[dict],
    ar_meta: Optional[Dict[str, dict]] = None,
) -> List[dict]:
    """对账表原始行 → 规范化记录列表。字段按中文名解析。"""
    idx = common.build_field_index(fields)
    ar_meta = ar_meta or {}
    out = []
    for row in rows:
        rec = {}
        for key, names in DUIZHANG_NAMES.items():
            rec[key] = common.get_by_names(row, idx, names)
        # 数值
        rec["amount_orig"] = common.to_number(rec.get("amount_orig"))
        rec["amount_local"] = common.to_number(rec.get("amount_local"))
        rec["deliver_local"] = common.to_number(rec.get("deliver_local"))
        rec["deliver_orig"] = common.to_number(rec.get("deliver_orig"))
        rec["hexiao_date"] = common.norm_date(rec.get("hexiao_date"))
        rec["shoukuan_date"] = common.norm_date(rec.get("shoukuan_date"))
        rec["so"] = (rec.get("so") or "").strip()
        rec["sod"] = (rec.get("sod") or "").strip()
        rec["ar"] = (rec.get("ar") or "").strip()
        rec["customer"] = (rec.get("customer") or "").strip()
        rec["status"] = (rec.get("status") or "").strip()
        rec["currency"] = (rec.get("currency") or "").strip()
        rec["huikuan_type"] = ""
        rec["fee"] = 0.0
        rec["channel"] = "duizhang"
        if rec["ar"] and rec["ar"] in ar_meta:
            meta = ar_meta[rec["ar"]]
            if not rec["currency"]:
                rec["currency"] = meta.get("currency") or ""
            if not rec["shoukuan_date"]:
                rec["shoukuan_date"] = meta.get("arrival_date")
            rec["huikuan_type"] = meta.get("huikuan_type") or ""
            rec["fee"] = meta.get("fee") or 0.0
            # 到账总额：跨行校验 Σ核销 ≤ 到账 要用（单条判不出超额核销）
            rec["arrival_total"] = meta.get("amount_orig")
            if meta.get("huikuan_type"):
                rec["channel"] = _channel_for_type(meta["huikuan_type"])
        # 币种兜底
        if not rec["currency"]:
            o, loc = rec["amount_orig"], rec["amount_local"]
            if o is not None and loc is not None and abs(o - loc) > 0.01:
                rec["currency"] = "未知外币"
            else:
                rec["currency"] = "人民币CNY"
        rec["source"] = "duizhang"
        out.append(rec)
    return out


def parse_huikuan_meta(rows: List[dict], fields: List[dict]) -> Dict[str, dict]:
    idx = common.build_field_index(fields)
    m: Dict[str, dict] = {}
    for row in rows:
        ar = common.get_by_names(row, idx, HUIKUAN_NAMES["ar"]).strip()
        if not ar:
            continue
        fee = common.to_number(common.get_by_names(row, idx, HUIKUAN_NAMES["fee"])) or 0.0
        m[ar] = {
            "currency": common.get_by_names(row, idx, HUIKUAN_NAMES["currency"]),
            "arrival_date": common.norm_date(
                common.get_by_names(row, idx, HUIKUAN_NAMES["arrival_date"])
            ),
            "huikuan_type": common.get_by_names(row, idx, HUIKUAN_NAMES["huikuan_type"]),
            "fee": fee,
            "amount_orig": common.to_number(
                common.get_by_names(row, idx, HUIKUAN_NAMES["amount_orig"])
            ),
            "status": common.get_by_names(row, idx, HUIKUAN_NAMES["status"]),
            "customer": common.get_by_names(row, idx, HUIKUAN_NAMES["customer"]),
            "hexiao_date": common.norm_date(
                common.get_by_names(row, idx, HUIKUAN_NAMES["hexiao_date"])
            ),
        }
    return m


def filter_yucun_details_by_date(
    details: List[dict],
    target_date: Optional[dt.date],
) -> List[dict]:
    """
    预存明细子表是累积的：必须按核销日期过滤本次，否则会重复回填。
    target_date 为 None 时不过滤（调用方应避免）。
    """
    if target_date is None:
        return list(details)
    out = []
    for d in details:
        hd = d.get("hexiao_date")
        if isinstance(hd, str):
            hd = common.norm_date(hd)
        if hd == target_date:
            out.append(d)
    return out


def parse_detail_rows(rows: List[dict], fields: Optional[List[dict]] = None) -> List[dict]:
    """核销明细（预存）→ 规范化。支持已是 plain dict 的合成数据。"""
    if fields:
        idx = common.build_field_index(fields)
        out = []
        for row in rows:
            out.append(
                {
                    "ar": common.get_by_names(row, idx, DETAIL_NAMES["ar"]).strip(),
                    "hexiao_date": common.norm_date(
                        common.get_by_names(row, idx, DETAIL_NAMES["hexiao_date"])
                    ),
                    "amount": common.to_number(
                        common.get_by_names(row, idx, DETAIL_NAMES["amount"])
                    ),
                    "currency": common.get_by_names(row, idx, DETAIL_NAMES["currency"]),
                    "rate": common.to_number(common.get_by_names(row, idx, DETAIL_NAMES["rate"])),
                    "so": common.get_by_names(row, idx, DETAIL_NAMES["so"]).strip(),
                }
            )
        return out
    # 合成 plain
    out = []
    for row in rows:
        out.append(
            {
                "ar": (row.get("ar") or row.get("回款记录NUM") or "").strip(),
                "hexiao_date": common.norm_date(row.get("hexiao_date") or row.get("核销日期")),
                "amount": common.to_number(row.get("amount") or row.get("本次核销金额")),
                "currency": (row.get("currency") or row.get("币种") or "人民币CNY"),
                "rate": common.to_number(row.get("rate") or row.get("汇率")),
                "so": (row.get("so") or row.get("SO") or "").strip(),
            }
        )
    return out


class LedgerIndex:
    """盈亏表 SO/SOD 索引（只读）。"""

    def __init__(self, path: Optional[Path] = None, synthetic: Optional[dict] = None):
        self.path = path
        self.so_index: Dict[str, List[int]] = {}
        self.sod_index: Dict[str, List[int]] = {}
        self.row_snapshot: Dict[int, dict] = {}
        self.cols: dict = {}
        if synthetic is not None:
            self.so_index = {k: list(v) for k, v in synthetic.get("so", {}).items()}
            self.sod_index = {k: list(v) for k, v in synthetic.get("sod", {}).items()}
            self.row_snapshot = synthetic.get("rows", {})
            return
        if path is None:
            return
        self._load(path)

    def _load(self, path: Path):
        import openpyxl

        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        if "明细" not in wb.sheetnames:
            wb.close()
            raise ValueError(f"盈亏表无『明细』sheet：{wb.sheetnames}")
        ws = wb["明细"]
        headers = None
        header_row_idx = None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
            names = [str(c).strip() if c is not None else "" for c in row]
            if any(n.startswith("新智云单号") for n in names) and any(
                "计提" in n for n in names
            ):
                headers = names
                header_row_idx = i
                break
        if headers is None:
            wb.close()
            raise ValueError("找不到盈亏表明细表头（需含新智云单号/计提）")
        aliases = common.load_aliases()
        cols = common.resolve_columns(
            headers,
            "盈亏明细",
            ["SO", "SOD", "计提", "回款明细", "是否结账", "收款时间", "收款方式"],
            aliases,
        )
        # 应收可选
        role = aliases.get("盈亏明细", {})
        yidx = common.fuzzy_find_col(headers, role.get("应收", ["应收金额", "应收"]))
        if yidx is not None:
            cols["应收"] = yidx
        self.cols = cols
        # 重建：read_only 流式
        wb.close()
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        ws = wb["明细"]
        for r_i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if header_row_idx is not None and r_i <= header_row_idx:
                continue
            vals = list(row)
            so = vals[cols["SO"]] if cols["SO"] < len(vals) else None
            sod = vals[cols["SOD"]] if cols["SOD"] < len(vals) else None
            so_s = str(so).strip() if so is not None else ""
            sod_s = str(sod).strip() if sod is not None else ""
            if so_s.startswith("SO"):
                self.so_index.setdefault(so_s, []).append(r_i)
            if sod_s.startswith("SOD"):
                self.sod_index.setdefault(sod_s, []).append(r_i)
            if so_s.startswith("SO") or sod_s.startswith("SOD"):
                snap = {
                    "so": so_s,
                    "sod": sod_s,
                    "jiti": vals[cols["计提"]] if cols["计提"] < len(vals) else None,
                    "huikuan": vals[cols["回款明细"]] if cols["回款明细"] < len(vals) else None,
                    "jiezhang": vals[cols["是否结账"]] if cols["是否结账"] < len(vals) else None,
                    "shoukuan_time": vals[cols["收款时间"]]
                    if cols["收款时间"] < len(vals)
                    else None,
                    "shoukuan_way": vals[cols["收款方式"]]
                    if cols["收款方式"] < len(vals)
                    else None,
                    "yingshou": vals[cols["应收"]]
                    if "应收" in cols and cols["应收"] < len(vals)
                    else None,
                }
                self.row_snapshot[r_i] = snap
        wb.close()

    def match(self, so: str, sod: str) -> Tuple[Optional[int], str, List[int]]:
        """返回 (row, how, candidates)。"""
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


def classify_one(
    rec: dict,
    ledger: Optional[LedgerIndex],
    rates: Dict[str, float],
    thr: float,
    year_now: int,
) -> dict:
    """单条判定 → {bucket, code, five_cols, ...}。"""
    result = {
        "ar": rec.get("ar") or "",
        "so": rec.get("so") or "",
        "sod": rec.get("sod") or "",
        # case_id = AR × SO：一笔到账挂多个 SO 时，每个 SO 是独立案例，
        # 否则台账会把多 SO 塌成一行、状态互相覆盖（实测 45% 的流转行是多 SO）。
        "case_id": f"{rec.get('ar') or '-'}|{rec.get('so') or '-'}",
        "customer_masked": common.mask_customer(rec.get("customer") or ""),
        "flow_hits": rec.get("flow_hits"),
        "flow_locate": rec.get("flow_locate") or "",
        "flow_matched_by": rec.get("flow_matched_by") or "",
        "flow_order_suggest": rec.get("flow_order_suggest") or "",
        "channel": rec.get("channel") or "duizhang",
        "huikuan_type": rec.get("huikuan_type") or "",
        "status": rec.get("status") or "",
        "bucket": "exception",
        "code": "",
        "reason": "",
        "five_cols": {},
        "locate_hint": "",
        "current_values": {},
        "candidates": [],
    }

    # 分笔 → 永远 hold
    if result["channel"] == "fenbi" or "分笔" in (rec.get("huikuan_type") or ""):
        result["bucket"] = "hold"
        result["code"] = "E1"
        result["reason"] = "分笔回款：系统无逐单金额，禁止比例分摊"
        return result

    if (rec.get("status") or "") == "已作废":
        result["code"] = "E7"
        result["reason"] = "核销已作废"
        return result

    if not rec.get("so") and not rec.get("sod"):
        result["code"] = "E7"
        result["reason"] = "无 SO/SOD 单号"
        return result

    # 手续费待拍板
    fee = rec.get("fee") or 0.0
    try:
        fee = float(fee)
    except (TypeError, ValueError):
        fee = 0.0
    if fee > 0.005:
        result["bucket"] = "hold"
        result["code"] = "E_FEE"
        result["reason"] = "有手续费，净额/含费口径未拍板"
        return result

    amount_orig = rec.get("amount_orig")
    if amount_orig is None:
        result["code"] = "E7"
        result["reason"] = "核销金额为空"
        return result

    local, err = common.compute_local_amount(amount_orig, rec.get("currency") or "", rates)
    if err == "E6":
        result["code"] = "E6"
        result["reason"] = f"外币缺汇率（{rec.get('currency')}）"
        return result

    # 跨年
    so = rec.get("so") or ""
    sod = rec.get("sod") or ""
    y = common.year_from_so(sod or so)
    if y is not None and y < year_now:
        result["bucket"] = "hold"
        result["code"] = "E3"
        result["reason"] = f"跨年老单（{y}）"
        return result

    # —— 上游/附加信号（流转三键、建档等）；有则优先成码，不瞎填 ——
    flow_hits = rec.get("flow_hits")  # None=未做三键；0=E0；1=唯一；>1=E12
    if flow_hits is not None:
        try:
            fh = int(flow_hits)
        except (TypeError, ValueError):
            fh = -1
        if fh == 0:
            result["bucket"] = "exception"
            result["code"] = "E0"
            result["reason"] = "流转表三键对不到账（0 行）"
            return result
        if fh > 1:
            result["bucket"] = "exception"
            result["code"] = "E12"
            result["reason"] = "同日同额同名命中多行"
            return result
    if rec.get("customer_archive_failed"):
        result["bucket"] = "exception"
        result["code"] = "E10"
        result["reason"] = "建档失败/搜不到客户"
        return result
    if rec.get("prepaid_system_gt_local"):
        result["bucket"] = "hold"
        result["code"] = "E9"
        result["reason"] = "系统可核金额大于线下预收剩余"
        return result
    if rec.get("deliver_changed"):
        result["bucket"] = "hold"
        result["code"] = "E11"
        result["reason"] = "系统交付额与表旧值不一致"
        return result

    # 匹配盈亏：必须有 ledger；无表或表中无 SO → hold E2，禁止无表 auto 瞎填
    row = None
    how = ""
    cands: List[int] = []
    if ledger is None:
        result["bucket"] = "hold"
        result["code"] = "E2"
        result["reason"] = "未提供盈亏表，无法确认 SO 是否在明细"
        return result

    row, how, cands = ledger.match(so, sod)
    if how == "E8":
        result["bucket"] = "hold"
        result["code"] = "E8"
        result["reason"] = "SO/SOD 在盈亏表多行"
        result["candidates"] = cands
        result["locate_hint"] = f"按 SO={so} 筛选，候选行号仅供参考勿当定位键：{cands}"
        return result
    if how == "E7" or (not so and not sod):
        result["bucket"] = "exception"
        result["code"] = "E7"
        result["reason"] = "无单号"
        return result
    if how == "E2" or row is None:
        result["bucket"] = "hold"
        result["code"] = "E2"
        result["reason"] = "盈亏表无此 SO/SOD"
        if so:
            result["locate_hint"] = f"在盈亏『明细』按「新智云单号」筛选：{so}（禁止用行号）"
        return result

    # 超额：本次本币 > 应收（人民币；容差 0.01）
    snap = ledger.row_snapshot.get(row, {}) if row is not None else {}
    yingshou = common.to_number(rec.get("yingshou"))
    if yingshou is None:
        yingshou = common.to_number(snap.get("yingshou"))
    if (
        local is not None
        and yingshou is not None
        and common.is_cny(rec.get("currency") or "")
        and float(local) > float(yingshou) + 0.01
    ):
        result["bucket"] = "exception"
        result["code"] = "E4"
        result["reason"] = "超额核销：本次本币大于应收"
        return result
    arrival_cap = common.to_number(rec.get("arrival_amount"))
    if (
        local is not None
        and arrival_cap is not None
        and float(local) > float(arrival_cap) + 0.01
    ):
        result["bucket"] = "exception"
        result["code"] = "E4"
        result["reason"] = "超额核销：核销额大于到账"
        return result

    # 尾差：交付 vs 核销
    deliver = rec.get("deliver_local")
    if deliver is not None and local is not None and thr is not None:
        if abs(float(deliver) - float(local)) > thr + 1e-9:
            # 完整核销场景下差额；阈值 0 时任何不等都挂
            st = rec.get("status") or ""
            if st in common.SETTLED and abs(float(deliver) - float(local)) > thr + 1e-9:
                # 外币：交付本币是系统汇率，不能与到账本币比——跳过
                if common.is_cny(rec.get("currency") or ""):
                    result["bucket"] = "hold"
                    result["code"] = "E5"
                    result["reason"] = "金额与交付额不一致（尾差阈值=0）"
                    return result

    if local is None:
        result["bucket"] = "exception"
        result["code"] = "E6"
        result["reason"] = "本币金额算不出"
        return result

    settled = (rec.get("status") or "") in common.SETTLED
    r_time = common.receipt_time(rec.get("shoukuan_date"), rec.get("hexiao_date"))
    way = common.pay_way(
        rec.get("status") or "",
        rec.get("huikuan_type") or "",
        common.norm_date(rec.get("shoukuan_date")),
        common.norm_date(rec.get("hexiao_date")),
    )
    local_f = float(local)

    five = {
        "计提": round(local_f, 2) if settled else None,
        "回款明细": round(local_f, 2),
        "是否结账": "是" if settled else "否",
        "收款时间": r_time.isoformat() if r_time else None,
        "收款方式": way,
        "实收SOD": sod or None,
    }
    result["five_cols"] = five
    result["bucket"] = "auto"
    result["code"] = ""
    result["reason"] = f"通道={result['channel']} 匹配={how}"
    if so:
        result["locate_hint"] = f"在盈亏『明细』按「新智云单号」筛选：{so}（禁止用行号）"
    if row is not None:
        result["current_values"] = {
            "计提": snap.get("jiti"),
            "回款明细": snap.get("huikuan"),
            "是否结账": str(snap.get("jiezhang") or "").strip() if snap.get("jiezhang") is not None else "",
            "收款时间": str(snap.get("shoukuan_time") or "")[:10],
            "收款方式": snap.get("shoukuan_way"),
            "实收SOD": snap.get("sod"),
        }
        result["ledger_row_ref"] = row  # 仅内部；清单不展示为定位键
    return result


def classify_records(
    records: List[dict],
    ledger: Optional[LedgerIndex] = None,
    rates: Optional[Dict[str, float]] = None,
) -> dict:
    """ledger 为 None 时每条进 hold E2（禁止无表瞎填五列）。"""
    rates = rates or {}
    thr = common.tail_threshold()
    year_now = common.current_year()
    results = [classify_one(rec, ledger, rates, thr, year_now) for rec in records]

    # —— 跨行校验：同一 AR 的核销合计不得超过到账（明妹明确约束；E4 单条判不出来）——
    over = check_ar_totals(records, thr)
    for r in results:
        info = over.get(r.get("ar") or "")
        if not info:
            continue
        r["bucket"] = "exception"
        r["code"] = "E4"
        r["reason"] = info
        r["five_cols"] = {}

    auto = [r for r in results if r["bucket"] == "auto"]
    hold = [r for r in results if r["bucket"] == "hold"]
    exc = [r for r in results if r["bucket"] == "exception"]
    return {
        "auto": auto,
        "hold": hold,
        "exception": exc,
        "counts": {
            "auto": len(auto),
            "hold": len(hold),
            "exception": len(exc),
            "total": len(records),
        },
        "e_code_dist": _dist(results),
        "ar_summary": build_ar_summary(results),
    }


def check_ar_totals(records: List[dict], thr: float = 0.0) -> Dict[str, str]:
    """
    同一 AR 的 Σ本次核销 ≤ 到账金额（含手续费两种口径都比，取宽松的一种）。
    超出 → 返回 {ar: 原因}，该 AR 全部行进 E4。
    只有拿得到 arrival_total 的 AR 才校验；拿不到就不判（不猜）。
    """
    sums: Dict[str, float] = {}
    arrival: Dict[str, float] = {}
    fees: Dict[str, float] = {}
    for rec in records:
        ar = rec.get("ar") or ""
        if not ar:
            continue
        amt = rec.get("amount_orig")
        if amt is not None:
            sums[ar] = round(sums.get(ar, 0.0) + float(amt), 2)
        at = rec.get("arrival_total")
        if at is not None:
            arrival[ar] = float(at)
        fee = rec.get("fee")
        if fee:
            fees[ar] = float(fee)
    bad: Dict[str, str] = {}
    for ar, total in sums.items():
        base = arrival.get(ar)
        if base is None:
            continue
        gross = base + fees.get(ar, 0.0)
        cap = max(base, gross) + max(thr, 0.0)
        if total > cap + 0.005:
            bad[ar] = (
                f"同一到账的核销合计 {total:.2f} 超过到账金额 "
                f"{base:.2f}（含手续费 {gross:.2f}）→ 超额核销，停下找销售核对"
            )
    return bad


def build_ar_summary(results: List[dict]) -> List[dict]:
    """
    按 AR 汇总每个 SO 的状态，并派生到账流转表「是否更新应收款」三态：
    全部 SO 可填 → 是；部分 → 部分；一个都没有 → 空白。
    （混合案例正是她表里出现"部分"的原因，实测微信表 2 行、汇款表 1 行。）
    """
    try:
        from flow_ledger import derive_flow_status
    except Exception:  # pragma: no cover - 仅防御
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
        out.append(
            {
                "ar": ar,
                "so_count": len(items),
                "buckets": states,
                "流转表_是否更新应收款_建议": derive_flow_status(states),
                "flow_locate": next((i.get("flow_locate") for i in items if i.get("flow_locate")), ""),
                "待处理SO": [i.get("so") for i in items if i["bucket"] != "auto" and i.get("so")],
            }
        )
    return out


def _dist(items: List[dict]) -> Dict[str, int]:
    d: Dict[str, int] = {}
    for it in items:
        c = it.get("code") or "OK"
        d[c] = d.get(c, 0) + 1
    return d


def records_from_fixture(
    fixture: dict,
    key: str = "day",
    yucun_details: Optional[List[dict]] = None,
) -> List[dict]:
    """从夹具构造规范化记录。day/std 走对账通道。"""
    fields = fixture.get("fields_duizhang") or []
    fields_hk = fixture.get("fields_huikuan") or []
    ar_rows = fixture.get("ar") or []
    ar_meta = parse_huikuan_meta(ar_rows, fields_hk) if ar_rows and fields_hk else {}

    raw = fixture.get(key)
    if raw is None:
        raise KeyError(f"夹具无键 {key}")
    if not isinstance(raw, list):
        raw = [raw]

    # 若 key 是预存合成
    if key == "yucun" or (raw and isinstance(raw[0], dict) and raw[0].get("channel") == "yucun"):
        return _records_yucun_synthetic(raw, yucun_details)

    recs = parse_duizhang_rows(raw, fields, ar_meta)
    return recs


def _records_yucun_synthetic(
    ar_list: List[dict],
    details: Optional[List[dict]],
) -> List[dict]:
    """合成预存：AR + 明细（已按日期过滤的由调用方保证）。"""
    details = details or []
    out = []
    for ar in ar_list:
        target = common.norm_date(ar.get("hexiao_date") or ar.get("核销日期"))
        filtered = filter_yucun_details_by_date(details, target) if details else []
        # 若 ar 自带 detail 列表
        if ar.get("details"):
            filtered = filter_yucun_details_by_date(
                parse_detail_rows(ar["details"]), target
            )
        if not filtered:
            # 无明细 → hold 类异常
            out.append(
                {
                    "ar": ar.get("ar") or "",
                    "so": ar.get("so") or "",
                    "sod": ar.get("sod") or "",
                    "customer": ar.get("customer") or "",
                    "amount_orig": None,
                    "currency": ar.get("currency") or "人民币CNY",
                    "hexiao_date": target,
                    "shoukuan_date": common.norm_date(ar.get("shoukuan_date")),
                    "status": ar.get("status") or "预存已核销",
                    "huikuan_type": "预存回款",
                    "channel": "yucun",
                    "fee": ar.get("fee") or 0,
                    "deliver_local": ar.get("deliver_local"),
                }
            )
            continue
        for d in filtered:
            out.append(
                {
                    "ar": ar.get("ar") or d.get("ar") or "",
                    "so": d.get("so") or ar.get("so") or "",
                    "sod": ar.get("sod") or "",
                    "customer": ar.get("customer") or "",
                    "amount_orig": d.get("amount"),
                    "currency": d.get("currency") or ar.get("currency") or "人民币CNY",
                    "hexiao_date": d.get("hexiao_date") or target,
                    "shoukuan_date": common.norm_date(ar.get("shoukuan_date")),
                    "status": ar.get("status") or "预存已核销",
                    "huikuan_type": "预存回款",
                    "channel": "yucun",
                    "fee": ar.get("fee") or 0,
                    "deliver_local": ar.get("deliver_local"),
                    "rate_arrival": d.get("rate"),
                }
            )
    return out


def load_excel_exports(workspace: Path) -> Optional[List[dict]]:
    """尝试从 01_智云导出 读 Excel（列名模糊匹配）。缺关键列则 raise。"""
    export_dir = workspace / "01_智云导出"
    if not export_dir.is_dir():
        return None
    import openpyxl

    aliases = common.load_aliases()
    # 优先对账表
    candidates = sorted(export_dir.glob("*.xlsx"))
    duizhang_path = None
    huikuan_path = None
    detail_path = None
    for p in candidates:
        n = p.name
        if "对账" in n:
            duizhang_path = p
        elif "明细" in n or "核销明细" in n:
            detail_path = p
        elif "回款记录" in n or "回款" in n:
            huikuan_path = p

    if not duizhang_path and not huikuan_path:
        return None

    ar_meta: Dict[str, dict] = {}
    if huikuan_path:
        wb = openpyxl.load_workbook(str(huikuan_path), read_only=True, data_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        headers = [str(c).strip() if c is not None else "" for c in next(rows_iter)]
        cols = common.resolve_columns(
            headers,
            "回款记录",
            ["AR", "回款类型"],
            aliases,
        )
        # optional cols
        opt_keys = ["核销日期", "到账日期", "到账金额原币", "手续费", "原币币种", "核销状态", "开票客户"]
        opt = {}
        role = aliases.get("回款记录", {})
        for k in opt_keys:
            idx = common.fuzzy_find_col(headers, role.get(k, [k]))
            if idx is not None:
                opt[k] = idx
        for row in rows_iter:
            vals = list(row)
            ar = str(vals[cols["AR"]] or "").strip()
            if not ar:
                continue
            ar_meta[ar] = {
                "huikuan_type": str(vals[cols["回款类型"]] or "").strip(),
                "currency": str(vals[opt["原币币种"]] or "").strip() if "原币币种" in opt else "",
                "arrival_date": common.norm_date(vals[opt["到账日期"]]) if "到账日期" in opt else None,
                "fee": common.to_number(vals[opt["手续费"]]) if "手续费" in opt else 0.0,
                "amount_orig": common.to_number(vals[opt["到账金额原币"]])
                if "到账金额原币" in opt
                else None,
                "status": str(vals[opt["核销状态"]] or "").strip() if "核销状态" in opt else "",
                "customer": str(vals[opt["开票客户"]] or "").strip() if "开票客户" in opt else "",
                "hexiao_date": common.norm_date(vals[opt["核销日期"]]) if "核销日期" in opt else None,
            }
        wb.close()

    records: List[dict] = []
    if duizhang_path:
        wb = openpyxl.load_workbook(str(duizhang_path), read_only=True, data_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        headers = [str(c).strip() if c is not None else "" for c in next(rows_iter)]
        cols = common.resolve_columns(
            headers,
            "回款核销对账",
            ["SO", "回款额原币", "核销日期"],
            aliases,
        )
        role = aliases.get("回款核销对账", {})
        opt = {}
        for k in ["SOD", "AR", "客户名称", "回款额本币", "交付额本币", "到账日期", "核销状态", "下单币种"]:
            idx = common.fuzzy_find_col(headers, role.get(k, [k]))
            if idx is not None:
                opt[k] = idx
        for row in rows_iter:
            vals = list(row)
            rec = {
                "so": str(vals[cols["SO"]] or "").strip(),
                "sod": str(vals[opt["SOD"]] or "").strip() if "SOD" in opt else "",
                "ar": str(vals[opt["AR"]] or "").strip() if "AR" in opt else "",
                "customer": str(vals[opt["客户名称"]] or "").strip() if "客户名称" in opt else "",
                "amount_orig": common.to_number(vals[cols["回款额原币"]]),
                "amount_local": common.to_number(vals[opt["回款额本币"]]) if "回款额本币" in opt else None,
                "deliver_local": common.to_number(vals[opt["交付额本币"]]) if "交付额本币" in opt else None,
                "hexiao_date": common.norm_date(vals[cols["核销日期"]]),
                "shoukuan_date": common.norm_date(vals[opt["到账日期"]]) if "到账日期" in opt else None,
                "status": str(vals[opt["核销状态"]] or "").strip() if "核销状态" in opt else "手动核销",
                "currency": str(vals[opt["下单币种"]] or "").strip() if "下单币种" in opt else "人民币CNY",
                "huikuan_type": "",
                "channel": "duizhang",
                "fee": 0.0,
            }
            if rec["ar"] and rec["ar"] in ar_meta:
                meta = ar_meta[rec["ar"]]
                rec["huikuan_type"] = meta.get("huikuan_type") or ""
                rec["channel"] = _channel_for_type(rec["huikuan_type"])
                rec["fee"] = meta.get("fee") or 0.0
                rec["arrival_total"] = meta.get("amount_orig")
                if not rec["currency"]:
                    rec["currency"] = meta.get("currency") or "人民币CNY"
            # 分笔：即使在对账表出现也 hold
            if rec["channel"] == "fenbi":
                pass
            records.append(rec)
        wb.close()

    # 预存：仅在回款记录有类型且有明细时展开
    if ar_meta and detail_path:
        wb = openpyxl.load_workbook(str(detail_path), read_only=True, data_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        headers = [str(c).strip() if c is not None else "" for c in next(rows_iter)]
        cols = common.resolve_columns(
            headers, "核销明细", ["回款记录NUM", "核销日期", "本次核销金额"], aliases
        )
        role = aliases.get("核销明细", {})
        opt = {}
        for k in ["币种", "汇率", "SO"]:
            idx = common.fuzzy_find_col(headers, role.get(k, [k]))
            if idx is not None:
                opt[k] = idx
        all_details = []
        for row in rows_iter:
            vals = list(row)
            so_raw = str(vals[opt["SO"]] or "").strip() if "SO" in opt else ""
            # 单元格里可能是「ZG SO2606…」或纯 SO 号
            so_val = ""
            for token in so_raw.replace("\n", " ").replace("、", " ").split():
                t = token.strip()
                if t.upper().startswith("SO") and len(t) >= 6:
                    so_val = t
                    break
            if not so_val and so_raw.upper().startswith("SO"):
                so_val = so_raw.split()[0]
            all_details.append(
                {
                    "ar": str(vals[cols["回款记录NUM"]] or "").strip(),
                    "hexiao_date": common.norm_date(vals[cols["核销日期"]]),
                    "amount": common.to_number(vals[cols["本次核销金额"]]),
                    "currency": str(vals[opt["币种"]] or "").strip() if "币种" in opt else "人民币CNY",
                    "rate": common.to_number(vals[opt["汇率"]]) if "汇率" in opt else None,
                    "so": so_val,
                }
            )
        wb.close()
        # 对预存 AR：展开过滤后的明细
        existing_ars = {r.get("ar") for r in records}
        for ar, meta in ar_meta.items():
            if _channel_for_type(meta.get("huikuan_type") or "") != "yucun":
                continue
            if ar in existing_ars:
                # 对账表不应有预存，若有也忽略对账行
                records = [r for r in records if r.get("ar") != ar]
            filtered = filter_yucun_details_by_date(
                [d for d in all_details if d["ar"] == ar],
                meta.get("hexiao_date") or meta.get("arrival_date"),
            )
            for d in filtered:
                records.append(
                    {
                        "ar": ar,
                        "so": d.get("so") or "",
                        "sod": "",
                        "customer": meta.get("customer") or "",
                        "amount_orig": d.get("amount"),
                        "currency": d.get("currency") or meta.get("currency") or "人民币CNY",
                        "hexiao_date": d.get("hexiao_date"),
                        "shoukuan_date": meta.get("arrival_date"),
                        "status": meta.get("status") or "预存已核销",
                        "huikuan_type": meta.get("huikuan_type") or "预存回款",
                        "channel": "yucun",
                        "fee": meta.get("fee") or 0.0,
                        "arrival_total": meta.get("amount_orig"),
                        "deliver_local": None,
                    }
                )

    # 纯分笔：对账无行时从 AR 直接 hold
    if ar_meta:
        covered = {r.get("ar") for r in records}
        for ar, meta in ar_meta.items():
            if ar in covered:
                continue
            ch = _channel_for_type(meta.get("huikuan_type") or "")
            if ch == "fenbi":
                records.append(
                    {
                        "ar": ar,
                        "so": "",
                        "sod": "",
                        "customer": meta.get("customer") or "",
                        "amount_orig": meta.get("amount_orig"),
                        "currency": meta.get("currency") or "人民币CNY",
                        "hexiao_date": meta.get("hexiao_date"),
                        "shoukuan_date": meta.get("arrival_date"),
                        "status": meta.get("status") or "",
                        "huikuan_type": meta.get("huikuan_type"),
                        "channel": "fenbi",
                        "fee": meta.get("fee") or 0.0,
                    }
                )
    return records


def serialize_result(result: dict) -> dict:
    """JSON 友好。"""

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
    ap = argparse.ArgumentParser(description="核销判定（三通道·三栏）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--fixture", default="", help="离线夹具 JSON")
    ap.add_argument("--key", default="day", help="夹具键 day/std")
    ap.add_argument("--ledger", default="", help="盈亏核算表副本（只读）")
    ap.add_argument("--rate", action="append", default=[], help="外币汇率 美元USD=7.0")
    ap.add_argument("--out", default="", help="判定结果 json 路径")
    ap.add_argument("--synthetic", default="", help="合成记录 JSON（测试用）")
    ap.add_argument("--flow", default="", help="到账流转表副本（只读）；不给则扫 02_我的表副本/")
    ap.add_argument(
        "--flow-complete",
        action="store_true",
        help="声明当天所有渠道的流转表都已给全；只有这时才判 E0（对不到账）",
    )
    args = ap.parse_args(argv)

    ws = Path(args.workspace)
    common.ensure_out_dirs(ws)
    rates = common.parse_rate_args(args.rate)
    # 夹具 std 默认美元汇率常测 7.0 —— 仅当用户显式传入才用；测试会传

    ledger = None
    if args.ledger:
        ledger = LedgerIndex(Path(args.ledger))
    else:
        # 工作区 02
        cand = list((ws / "02_我的表副本").glob("*盈亏*")) if (ws / "02_我的表副本").is_dir() else []
        if cand:
            ledger = LedgerIndex(cand[0])
        else:
            print(
                "WARN: 未提供盈亏表（--ledger 或 02_我的表副本/*盈亏*）；"
                "全部无法匹配的单将 hold E2，不会瞎填五列",
                file=sys.stderr,
            )
            ledger = None  # 明确 None → classify 一律 E2，禁止空索引 auto

    records: List[dict] = []
    if args.synthetic:
        records = json.loads(Path(args.synthetic).read_text(encoding="utf-8"))
    elif args.fixture:
        fix = load_fixture(Path(args.fixture))
        records = records_from_fixture(fix, args.key)
    else:
        try:
            records = load_excel_exports(ws) or []
        except ValueError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return 2
        if not records:
            # 兜底：01_智云导出/ 里放的是 .json 取数夹具时也认（测试/演练常这么放）
            js = sorted((ws / "01_智云导出").glob("*.json")) if (ws / "01_智云导出").is_dir() else []
            for j in js:
                try:
                    fix = load_fixture(j)
                    if isinstance(fix, dict) and (fix.get("day") or fix.get("std")):
                        records = records_from_fixture(fix, args.key)
                        print(f"用取数夹具：{j.name}（key={args.key}）")
                        break
                except Exception:
                    continue
        if not records:
            print(
                "ERROR: 第 6 步跑不了 —— 01_智云导出/ 里没有可用的智云导出。\n"
                "  需要：文件名含「回款」的回款记录 + 含「对账」的回款核销对账（预存类另需含「明细」的一份）。\n"
                "  请她导出后放进 01_智云导出/ 再跑；不要改用别的文件凑合。",
                file=sys.stderr,
            )
            return 2

    # —— 到账流转表三键匹配（R2）：没有它，E0/E12 判不出来，清单也说不出"在流转表哪一行" ——
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
    today = dt.date.today().strftime("%Y%m%d")
    # 产出必须跟随 --workspace（此前写死 common.OUT_DIR，换工作区会把结果写回技能自带目录）
    out_path = Path(args.out) if args.out else (ws / "04_产出" / f"判定结果_{today}.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(serialize_result(result), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    c = result["counts"]
    print(f"判定完成 total={c['total']} auto={c['auto']} hold={c['hold']} exception={c['exception']}")
    print(f"E码分布: {result['e_code_dist']}")
    print(f"结果: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
