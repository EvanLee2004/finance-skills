#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
到账流转表接入（R2 三键匹配）。

为什么单独一个模块：判定要回答"这笔核销对应到账流转表的哪一行、该写什么单号"，
这是第 6 步的一半。此前 classify 只读 rec["flow_hits"] 而没人给它赋值 → E0/E12 是死码。

设计要点
- **按表头内容认表，不按文件名**（她各渠道分开好几张：汇款/微信/支付宝/美元户）。
- 匹配键按可靠度排序（参考李尚 business-rules §3）：
  1. 到账日期 + 净到账额 + 公司名/汇款人
  2. 到账日期 + （净额 + 手续费）+ 公司名  ← 微信/支付宝把客户支付额拆成净额与手续费
  3. 到账日期 + 金额（任一口径），公司名对不上 → **弱命中**，仍要人确认
- 命中 0 → E0；命中 >1 → E12；命中 1 → 给出行号与建议单号。
- **只读**：本模块从不写流转表。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

# 认表特征。**必须有流转表专有列**，否则盈亏「明细」会被误认成流转表
# （它也有"单号/客户名称/项目下单日期/应收金额"，实测把 5000 行明细当成了流转行）。
_REQUIRED_HINTS = ("单号",)
_SIGNATURE_COLS = ("是否更新应收款", "是否已登记系统", "收款形式")
_NAME_HINTS = ("公司名称", "汇款人", "对方户名")
# 盈亏表专有列：出现即判定"这不是流转表"
_LEDGER_MARKERS = ("计提金额", "新智云单号", "回款明细", "是否结账")

_NOISE = re.compile(r"[（）()\s·、,，.。\-—_]|有限公司|股份|集团|分公司|总公司")


def normalize_name(s: str) -> str:
    """公司名归一：去噪声词与标点，便于子串比对。"""
    return _NOISE.sub("", str(s or "")).strip().lower()


def name_similar(a: str, b: str) -> bool:
    """一方是另一方子串即算相似（汇款人与开票客户常不完全一致）。"""
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return False
    return na == nb or na in nb or nb in na


def amounts_equal(a: Optional[float], b: Optional[float], tol: float = 0.005) -> bool:
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) <= tol


class FlowLedger:
    """一份或多份到账流转表的只读索引。"""

    def __init__(self, rows: Optional[List[dict]] = None):
        self.rows: List[dict] = rows or []
        self.sources: List[str] = []

    # ---------- 装载 ----------

    @staticmethod
    def _looks_like_flow(headers: Sequence[str]) -> bool:
        norm = [common._norm(h) for h in headers]
        joined = "|".join(norm)
        if any(m in joined for m in _LEDGER_MARKERS):
            return False  # 盈亏明细，不是流转表
        if not all(h in joined for h in _REQUIRED_HINTS):
            return False
        if not any(s in joined for s in _SIGNATURE_COLS):
            return False  # 没有流转表专有列 → 不认
        return any(h in joined for h in _NAME_HINTS)

    @classmethod
    def from_paths(cls, paths: Sequence[Path]) -> "FlowLedger":
        import openpyxl

        aliases = common.load_aliases()
        inst = cls()
        for p in paths:
            try:
                wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
            except Exception:
                continue
            for ws in wb.worksheets:
                all_rows = list(ws.iter_rows(values_only=True))
                if not all_rows:
                    continue
                hrow, headers = common.find_header_row(
                    all_rows, "到账流转", ["日期", "公司名称", "金额", "单号"], aliases
                )
                if not cls._looks_like_flow(headers):
                    continue
                it = iter(all_rows[hrow + 1:])
                try:
                    cols = common.resolve_columns(
                        headers, "到账流转", ["日期", "公司名称", "金额", "单号"], aliases
                    )
                except ValueError:
                    continue
                opt = {}
                for k in ["收款形式", "是否更新应收款", "是否已登记系统"]:
                    idx = common.fuzzy_find_col(
                        headers, aliases.get("到账流转", {}).get(k, [k])
                    )
                    if idx is not None:
                        opt[k] = idx
                rno = hrow + 1
                for row in it:
                    rno += 1
                    vals = list(row)

                    def cell(i):
                        return vals[i] if i is not None and i < len(vals) else None

                    date = common.norm_date(cell(cols["日期"]))
                    amount = common.to_number(cell(cols["金额"]))
                    if date is None and amount is None:
                        continue
                    inst.rows.append(
                        {
                            "file": p.name,
                            "sheet": ws.title,
                            "row_no": rno,  # 1-based，含表头
                            "date": date,
                            "payer": str(cell(cols["公司名称"]) or "").strip(),
                            "amount": amount,
                            "order_cell": str(cell(cols["单号"]) or "").strip(),
                            "form": str(cell(opt.get("收款形式")) or "").strip()
                            if "收款形式" in opt
                            else "",
                            "updated": str(cell(opt.get("是否更新应收款")) or "").strip()
                            if "是否更新应收款" in opt
                            else "",
                            "registered": str(cell(opt.get("是否已登记系统")) or "").strip()
                            if "是否已登记系统" in opt
                            else "",
                        }
                    )
                inst.sources.append(f"{p.name}#{ws.title}")
            wb.close()
        return inst

    @classmethod
    def from_workspace(cls, workspace: Path) -> "FlowLedger":
        """从 02_我的表副本 按内容识别所有流转表（不看文件名）。"""
        d = Path(workspace) / "02_我的表副本"
        if not d.is_dir():
            return cls()
        return cls.from_paths(sorted(d.glob("*.xlsx")))

    # ---------- 匹配 ----------

    def match(
        self,
        arrival_date,
        amount_net: Optional[float],
        customer: str = "",
        fee: Optional[float] = 0.0,
    ) -> dict:
        """
        返回 {"hits": n, "rows": [...], "matched_by": str}
        matched_by ∈ 三键 / 三键(含手续费) / 日期+金额(名字不符) / ""
        """
        date = common.norm_date(arrival_date)
        gross = None
        if amount_net is not None and fee:
            gross = round(float(amount_net) + float(fee), 2)

        def by_date(r):
            return date is not None and r["date"] == date

        cand_net = [r for r in self.rows if by_date(r) and amounts_equal(r["amount"], amount_net)]
        cand_gross = (
            [r for r in self.rows if by_date(r) and amounts_equal(r["amount"], gross)]
            if gross is not None
            else []
        )

        for pool, tag in ((cand_net, "三键"), (cand_gross, "三键(含手续费)")):
            named = [r for r in pool if name_similar(r["payer"], customer)]
            if named:
                return {"hits": len(named), "rows": named, "matched_by": tag}

        weak = cand_net + [r for r in cand_gross if r not in cand_net]
        if weak:
            return {
                "hits": len(weak),
                "rows": weak,
                "matched_by": "日期+金额(名字不符)",
            }
        return {"hits": 0, "rows": [], "matched_by": ""}

    # ---------- 输出辅助 ----------

    @staticmethod
    def locate_text(hit: dict) -> str:
        """给她看的定位说明（不是行号定位盈亏表，这里是流转表自己的行）。"""
        if not hit or not hit.get("rows"):
            return ""
        r = hit["rows"][0]
        return f"{r['file']}#{r['sheet']} 第{r['row_no']}行（{hit.get('matched_by','')}）"

    @staticmethod
    def suggest_order_cell(so_list: Sequence[str], existing: str = "") -> str:
        """建议写进流转表「单号」列的文本：多 SO 换行分隔，已有的不重复加。"""
        have = {s.strip() for s in re.split(r"[\s\n,，;；]+", existing or "") if s.strip()}
        add = [s for s in dict.fromkeys([s.strip() for s in so_list if s and s.strip()]) if s not in have]
        if not add:
            return existing or ""
        parts = ([existing.strip()] if existing and existing.strip() else []) + add
        return "\n".join(parts)


def derive_flow_status(so_states: Sequence[str]) -> str:
    """
    到账流转表「是否更新应收款」三态（李尚 business-rules §9 + 明妹 7-23 录音）。

    传进来每个 state 是这笔到账下某个 SOD 今天**能不能在她盈亏表里更新**：
      · "ready" —— 订单已在她表里，这次能填（auto），或只是要拆行/指认（E5/E8）——她照样更得动
      · "wait"  —— 订单还没交付进表（E2/E3），或数据有问题（异常），今天更新不了

    规则：全部 ready → 是；一个 ready 都没有 → 空白；有的 ready 有的 wait → 部分。

    ⚠ 2026-07-24 关键修正（明妹 7-23 录音原话："部分是指某些订单没交付、我没法更新"）：
      **「部分核销」(E5) 不再算 wait**。她那笔到账的钱已经全部核销落地，只是要在表里
      把一个订单拆成「已收/未收」两行——这算她**更新了**。以前把 E5 也算成"没更新"，
      导致一笔明明全核掉的到账被误标「部分」（实测 AR26070112）。
      只有「订单没交付进表」(E2/E3) 才让这笔到账变「部分」。
      兼容旧入参：历史上传的是 bucket("auto"/"hold")，"auto" 仍按 ready 计。
    """
    def _ready(s: str) -> bool:
        return s in ("ready", "auto")

    states = [s for s in so_states if s]
    if not states:
        return ""
    ready = sum(1 for s in states if _ready(s))
    if ready == 0:
        return ""
    if ready == len(states):
        return "是"
    return "部分"


def flow_status_policy(status: str) -> dict:
    """
    明妹口径（2026-07-23 核销清单回填讨论）：
    - 全部订单没检索到 → 「是否更新」**留空**，**不要公式**
    - 全部订单检索到并更新 → 填「是」；表若有惯用公式可套公式（本程序建议字面「是」）
    - 部分更新 → 填「部分」，**颜色标出**哪些 SO 更了 / 没更
    程序第一版只给建议，不自动写她的流转表。
    """
    s = (status or "").strip()
    if s in ("", "（空白）", "空白", "空"):
        return {
            "填法": "留空",
            "公式策略": "不要公式、不要写「否」",
            "颜色标注": "无需（整笔都还没更）",
            "人话": "这笔到账关联的订单这次一个都还没能更新 → 「是否更新应收款」空着",
        }
    if s == "是":
        return {
            "填法": "是",
            "公式策略": "可套你表里惯用公式；没有公式就填字面「是」",
            "颜色标注": "无需（全部已更新）",
            "人话": "这笔到账下订单全部检索到并已更新 → 填「是」",
        }
    if s in ("部分", "部份"):
        return {
            "填法": "部分",
            "公式策略": "不要写成「是」的公式；字面「部分」",
            "颜色标注": "已更新 SO 正常色；未更新 SO 标红（在单号列或备注里区分）",
            "人话": "同一笔到账有的单更新了、有的还没 → 填「部分」，红字标没更新的单号",
        }
    return {
        "填法": s,
        "公式策略": "按明妹现场口径",
        "颜色标注": "—",
        "人话": f"状态={s}",
    }


def annotate_records(
    records: List[dict],
    flow: Optional[FlowLedger],
    complete: bool = False,
) -> List[dict]:
    """
    给每条记录补 flow_hits / flow_locate / flow_order_suggest（原地）。

    **默认不判 E0**（complete=False）。她各渠道分开好几张表（汇款/微信/支付宝/美元户），
    手上若只有其中一两张，其余渠道的到账当然找不到 —— 那是"没给全数据"，不是"对不到账"。
    实测：只给汇款+微信两张表跑 7-08 的夹具，28 笔会被全部误判成 E0。

    只有运行时显式声明"当天所有渠道的流转表都给全了"（--flow-complete）才判 E0；
    否则找不到就标注"未在现有流转表中找到"，让记录按原通道规则继续判，不冤枉她。
    """
    if flow is None or not flow.rows:
        return records
    dates = [r["date"] for r in flow.rows if r.get("date")]
    dmin, dmax = (min(dates), max(dates)) if dates else (None, None)
    cache: Dict[tuple, dict] = {}
    for rec in records:
        # 三键里的「金额」必须是**这笔到账的总额**，不是单个 SO/SOD 的金额。
        # 流转表一行 = 银行一笔到账；一笔到账挂 N 个订单时拿分项金额去比，永远对不上。
        # （旧版就是拿 amount_orig 比的 → 多 SO 的到账全都匹配不到流转行。）
        amount = rec.get("arrival_total")
        if amount is None:
            amount = rec.get("amount_orig")
        key = (
            str(rec.get("shoukuan_date")),
            amount,
            rec.get("customer") or "",
            rec.get("fee") or 0.0,
        )
        if key not in cache:
            cache[key] = flow.match(
                rec.get("shoukuan_date"),
                amount,
                rec.get("customer") or "",
                rec.get("fee") or 0.0,
            )
        hit = cache[key]
        d = common.norm_date(rec.get("shoukuan_date"))
        covered = dmin is not None and d is not None and dmin <= d <= dmax
        if not covered:
            rec["flow_hits"] = None
            rec["flow_matched_by"] = "流转表未覆盖该日期(未做三键)"
            rec["flow_locate"] = ""
            continue
        if hit["hits"] == 0 and not complete:
            rec["flow_hits"] = None  # 不判 E0：可能只是这个渠道的表没给
            rec["flow_matched_by"] = "未在现有流转表中找到(可能缺该渠道的表)"
            rec["flow_locate"] = ""
            continue
        rec["flow_hits"] = hit["hits"]
        rec["flow_matched_by"] = hit["matched_by"]
        rec["flow_locate"] = FlowLedger.locate_text(hit)
        if hit["hits"] == 1 and rec.get("so"):
            rec["flow_order_suggest"] = FlowLedger.suggest_order_cell(
                [rec.get("so")], hit["rows"][0].get("order_cell", "")
            )
    return records


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="到账流转表索引自检（只读）")
    ap.add_argument("--workspace", default=str(common.WORK))
    ap.add_argument("--paths", nargs="*", default=[])
    args = ap.parse_args(argv)

    flow = (
        FlowLedger.from_paths([Path(p) for p in args.paths])
        if args.paths
        else FlowLedger.from_workspace(Path(args.workspace))
    )
    print(f"流转表来源: {flow.sources or '（无）'}")
    print(f"可用行数: {len(flow.rows)}")
    if not flow.rows:
        print("WARN: 没有认出任何到账流转表 → 三键匹配不可用，E0/E12 无法判定", file=sys.stderr)
        return 1
    filled = sum(1 for r in flow.rows if r.get("order_cell"))
    multi = sum(1 for r in flow.rows if len(re.split(r"[\s\n]+", r["order_cell"].strip())) > 1)
    print(f"已填单号行: {filled}；其中多 SO 行: {multi}")
    print(json.dumps({"sources": flow.sources, "rows": len(flow.rows)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
