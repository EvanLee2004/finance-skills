#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第 1 步：智云只读取数 → 工作区 01_智云导出/

取数架构（2026-07-23 明妹当面口径 + Claude 接口实调重写，v2）
────────────────────────────────────────────────────────────
**入口只有一个**：财务管理 →「回款记录」，**只筛一个条件：核销日期 = T-1**。
然后把每笔回款记录**点进去的明细页**原样抓下来——这正是明妹手工在做的事：

  回款记录（1 笔到账）
    ├─ 关联「下单」            → 这笔到账关联了哪几个 SO + 每个 SO 的交付额
    ├─ 关联「订单同币种核销明细信息」→ 逐 SO 的本次核销金额
    │                            **0 行 = 全额核销**（明妹原话：没有这张就说明到账=交付）
    └─ 由 SO 去「订单明细」表   → 每个 SO 下的 SOD + 逐 SOD 交付额
                                 （SOD 的唯一来源；她盈亏表一行 = 一个 SOD）

**已废弃**：按核销日期单独筛「回款核销对账」表。原因（2026-07-23 实调）：
  1. 它对预存/预收类返回 0 行，靠"按回款类型走不同通道"来补，而通道路由正是
     2026-07-22 那次**3 笔到账 6.5 万静默消失**的根因（预收类既被踢出对账通道、
     又在明细子表拿不到行，两头落空还不报错）。
  2. 它拿不到 SOD 与逐 SOD 交付额，判不出"一个 SO 拆 N 个 SOD"这个 45% 的主场景。
  3. 它是独立筛的第二张表，与回款记录可能不同步；关联读法天然同步。

红线：只读（GetFilterRows / getRowRelationRows / getWorksheetInfo，无任何写接口）；
      账号密码每次运行时提供，绝不写进代码 / config / git。

用法：
  export ZHIYUN_USER='你的智云账号'
  export ZHIYUN_PASS   # 在 shell 里 export，勿写进任何文件；用完 unset
  python3 scripts/fetch_zhiyun.py --date 2026-07-22 --workspace 工作区/

  或交互（推荐，密码不回显、不落盘）：
  python3 scripts/fetch_zhiyun.py --date yesterday --workspace 工作区/

依赖：playwright（登录）+ requests（取数）+ openpyxl（写出 xlsx）
  pip install playwright requests openpyxl && playwright install chromium
"""
from __future__ import annotations

import argparse
import getpass
import json
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

# ── 常量（表 ID / 字段 ID 来自 2026-07-09/22/23 勘探，非密钥）────────────────
BASE_DEFAULT = "http://192.168.10.167:18880"
APP_ID = "6ff4fb2e-e68c-4ee9-83a0-836de8f72c11"

WS_HUIKUAN = "6555d2b1f9460e517040ba6c"  # 回款记录（唯一入口）

# 回款记录字段（只保留新架构真正用到的）
F_HK = {
    "arrival_date": "6555d2b1f9460e517040ba70",
    "amount_orig": "6555d2b1f9460e517040ba71",
    "currency": "663c495f4a2483fd86249c90",
    "ar": "663c4b8a4a2483fd86249ca8",
    "amount_local": "663c4d354a2483fd86249cbb",
    "status": "663ca2204a2483fd8624a107",
    "hexiao_date": "664c571b4a2483fd86250fa5",
    "huikuan_type": "66b1f6986429811e5f304a0b",
    "fee": "675119ab327314202700730a",
    "customer_rel": "676a3314327314202700e7fc",
    "customer_txt": "676b6d01327314202700ed59",
}

# 关联字段按**中文名**取（controlId 会随配置变，名字不会）
REL_XIADAN = "下单"
REL_HEXIAO_MINGXI = "订单同币种核销明细信息"
REL_SODLINE = "订单明细"  # 只借它的 dataSource 定位「订单明细」表

# 关联子表里要取的列（按中文名，取不到就空，不猜）
XIADAN_COLS = ["SO", "交付额/原币", "汇率", "结算币种", "订单名称"]
# ⚠「同币种核销明细信息」表里**没有**叫 SO 的字段，SO 藏在关联字段「订单NUM」的 name 里
#   （2026-07-23 实调：该表字段 = 核销记录NUM/回款记录NUM/订单NUM/本次核销金额/…）。
#   旧版靠"按金额跟下单栏配对"猜 SO，金额一撞就配错；这一版直接从订单NUM读，不猜。
MINGXI_COLS = [
    "订单NUM", "本次核销金额", "本次核销金额本币", "核销日期",
    "币种", "汇率", "订单名称", "是否已撤销",
]
SODLINE_COLS = ["SO", "SOD", "交付额/原币", "币种", "项目状态"]


class LoginError(RuntimeError):
    pass


class FetchError(RuntimeError):
    pass


def _plain(v: Any, options: Optional[Dict[str, str]] = None) -> str:
    """明道云单元格 → 纯文本。"""
    if v is None:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(v)
    s = str(v).strip()
    if not s:
        return ""
    if s.startswith("["):
        try:
            arr = json.loads(s)
        except ValueError:
            return s
        out: List[str] = []
        for item in arr:
            if isinstance(item, dict):
                out.append(
                    item.get("name")
                    or item.get("fullname")
                    or item.get("sourcevalue")
                    or item.get("departmentName")
                    or ""
                )
            elif options and str(item) in options:
                out.append(options[str(item)])
            else:
                out.append(str(item))
        return "、".join(x for x in out if x)
    if options and s in options:
        return options[s]
    return s


def extract_so(text: str) -> str:
    """从 'ZG SO26060433' / 'SO26060433' / '[{...name:SO...}]' 里抠出 SO 号。"""
    s = _plain(text)
    for token in s.replace("\n", " ").replace("、", " ").replace(",", " ").split():
        t = token.strip().upper()
        if t.startswith("SO") and not t.startswith("SOD") and len(t) >= 8:
            return token.strip()
    return ""


def resolve_date(s: str) -> str:
    s = (s or "").strip().lower()
    if s in ("yesterday", "t-1", "昨天"):
        return (date.today() - timedelta(days=1)).isoformat()
    if s in ("today", "今天"):
        return date.today().isoformat()
    datetime.strptime(s, "%Y-%m-%d")
    return s


def login_with_password(
    base_url: str, username: str, password: str, headless: bool = True
) -> Tuple[str, Optional[str]]:
    """账号密码 → (md_pss_id, account_id)。不落盘。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise LoginError(
            "未安装 playwright。请: pip install playwright && playwright install chromium"
        ) from e

    account_sel = "#txtMobilePhone"
    password_sel = "input[type=password]"
    btn_sels = ["text=登 录", "text=登录", ".loginBtn"]

    try:
        with sync_playwright() as p:
            br = p.chromium.launch(headless=headless)
            try:
                ctx = br.new_context(ignore_https_errors=True)
                pg = ctx.new_page()
                pg.goto(base_url, wait_until="networkidle", timeout=45000)
                pg.fill(account_sel, username)
                pg.fill(password_sel, password)
                clicked = False
                for sel in btn_sels:
                    try:
                        pg.click(sel, timeout=2500)
                        clicked = True
                        break
                    except Exception:
                        continue
                if not clicked:
                    pg.keyboard.press("Enter")
                pg.wait_for_timeout(7000)
                token = None
                for c in ctx.cookies():
                    if c.get("name") == "md_pss_id" and c.get("value"):
                        token = c["value"]
                        break
                if not token:
                    raise LoginError(
                        f"登录后未拿到 md_pss_id（url={pg.url}）。账号密码可能错误或页面结构变了。"
                    )
                acct = None
                try:
                    acct = pg.evaluate(
                        "() => { try { return (md && md.global && md.global.Account && "
                        "md.global.Account.accountId) || null; } catch(e) { return null; } }"
                    )
                except Exception:
                    acct = None
                return token, acct
            finally:
                br.close()
    except LoginError:
        raise
    except Exception as e:
        raise LoginError(f"登录异常 {type(e).__name__}: {e}") from e


class ZhiyunClient:
    def __init__(self, base: str, cookie: str, account_id: str = "", page_size: int = 200):
        import requests

        self.base = base.rstrip("/")
        self.page_size = page_size
        self.session = requests.Session()
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"md_pss_id {cookie}",
            "X-Requested-With": "XMLHttpRequest",
        }
        if account_id:
            self.headers["AccountId"] = account_id
        self._tpl_cache: Dict[str, List[dict]] = {}

    def post(self, path: str, body: dict, timeout: int = 90) -> dict:
        url = f"{self.base}/wwwapi/{path.lstrip('/')}"
        r = self.session.post(url, headers=self.headers, json=body, timeout=timeout)
        r.raise_for_status()
        j = r.json()
        if isinstance(j, dict) and "data" in j:
            return j["data"] if j["data"] is not None else {}
        return j

    # ── 模板 / 字段 ──────────────────────────────────────────────
    def controls(self, worksheet_id: str) -> List[dict]:
        if worksheet_id in self._tpl_cache:
            return self._tpl_cache[worksheet_id]
        info = self.post(
            "worksheet/getWorksheetInfo",
            {"worksheetId": worksheet_id, "appId": APP_ID, "getTemplate": True},
        )
        ctrls = (info.get("template") or {}).get("controls") or info.get("controls") or []
        self._tpl_cache[worksheet_id] = ctrls
        return ctrls

    @staticmethod
    def name_map(controls: Sequence[dict]) -> Dict[str, str]:
        """controlId → 中文名。"""
        return {
            (c.get("controlId") or c.get("id") or ""): (c.get("controlName") or c.get("name") or "")
            for c in controls
        }

    @staticmethod
    def id_by_name(controls: Sequence[dict], name: str) -> str:
        for c in controls:
            if (c.get("controlName") or c.get("name") or "") == name:
                return c.get("controlId") or c.get("id") or ""
        return ""

    @staticmethod
    def option_maps(controls: Sequence[dict]) -> Dict[str, Dict[str, str]]:
        """controlId → {optionKey: 显示值}。"""
        out: Dict[str, Dict[str, str]] = {}
        for c in controls:
            opts = c.get("options") or []
            if opts:
                cid = c.get("controlId") or c.get("id") or ""
                out[cid] = {str(o.get("key")): str(o.get("value")) for o in opts}
        return out

    def datasource_of(self, worksheet_id: str, control_name: str) -> str:
        for c in self.controls(worksheet_id):
            if (c.get("controlName") or c.get("name") or "") == control_name:
                return c.get("dataSource") or ""
        return ""

    # ── 取行 ────────────────────────────────────────────────────
    def filter_rows_by_date(
        self, worksheet_id: str, date_control_id: str, day: str, page_size: Optional[int] = None
    ) -> Tuple[List[dict], int]:
        """按日期字段筛单日（filterType 11 + dateRange 18 已实测）。**只此一个条件。**"""
        ps = page_size or self.page_size
        rows: List[dict] = []
        page = 1
        total = 0
        while True:
            d = self.post(
                "worksheet/getFilterRows",
                {
                    "worksheetId": worksheet_id,
                    "appId": APP_ID,
                    "pageSize": ps,
                    "pageIndex": page,
                    "status": 1,
                    "sortControls": [],
                    "notGetTotal": False,
                    "searchType": 1,
                    "keyWords": "",
                    "filterControls": [
                        {
                            "controlId": date_control_id,
                            "dataType": 30,
                            "spliceType": 1,
                            "filterType": 11,
                            "dateRange": 18,
                            "minValue": day,
                            "maxValue": day,
                            "value": "",
                            "values": [],
                        }
                    ],
                    "fastFilters": [],
                    "navGroupFilters": [],
                },
            )
            batch = d.get("data") if isinstance(d, dict) else d
            batch = batch if isinstance(batch, list) else []
            if isinstance(d, dict):
                total = int(d.get("count") or total or 0)
            rows.extend(batch)
            if not batch or (total and len(rows) >= total) or len(batch) < ps:
                break
            page += 1
            if page > 500:
                raise FetchError(f"翻页超过 500，worksheet={worksheet_id}")
        return rows, total or len(rows)

    def search_rows(self, worksheet_id: str, keyword: str, page_size: int = 200) -> List[dict]:
        """全文检索（用于按 SO 找订单明细）。调用方必须再做精确过滤。"""
        rows: List[dict] = []
        page = 1
        while page <= 20:
            d = self.post(
                "worksheet/getFilterRows",
                {
                    "worksheetId": worksheet_id,
                    "appId": APP_ID,
                    "pageSize": page_size,
                    "pageIndex": page,
                    "status": 1,
                    "sortControls": [],
                    "notGetTotal": False,
                    "searchType": 1,
                    "keyWords": keyword,
                    "filterControls": [],
                    "fastFilters": [],
                    "navGroupFilters": [],
                },
            )
            batch = d.get("data") if isinstance(d, dict) else d
            batch = batch if isinstance(batch, list) else []
            rows.extend(batch)
            if len(batch) < page_size:
                break
            page += 1
        return rows

    def relation_rows(
        self, worksheet_id: str, row_id: str, control_id: str, page_size: int = 100
    ) -> Tuple[List[dict], List[dict]]:
        """返回 (行列表, 目标表 controls)。"""
        rows: List[dict] = []
        controls: List[dict] = []
        page = 1
        while page <= 50:
            d = self.post(
                "worksheet/getRowRelationRows",
                {
                    "worksheetId": worksheet_id,
                    "rowId": row_id,
                    "controlId": control_id,
                    "pageIndex": page,
                    "pageSize": page_size,
                    "appId": APP_ID,
                    "getWorksheet": page == 1,
                },
            )
            batch: Any = []
            if isinstance(d, dict):
                batch = d.get("data") or d.get("rows") or []
                if isinstance(batch, dict):
                    batch = batch.get("data") or []
                if page == 1:
                    w = d.get("worksheet") or {}
                    controls = (w.get("template") or {}).get("controls") or w.get("controls") or []
            batch = batch if isinstance(batch, list) else []
            rows.extend(batch)
            if len(batch) < page_size:
                break
            page += 1
        return rows, controls


def pick_named(
    row: dict, names: Dict[str, str], opts: Dict[str, Dict[str, str]], wanted: Sequence[str]
) -> Dict[str, str]:
    """按中文列名从原始行取值 → {列名: 纯文本}。取不到的列留空，不猜。"""
    by_name = {v: k for k, v in names.items() if v}
    out: Dict[str, str] = {}
    for w in wanted:
        cid = by_name.get(w, "")
        out[w] = _plain(row.get(cid), opts.get(cid)) if cid else ""
    return out


def write_xlsx(path: Path, headers: List[str], rows: List[List[Any]]) -> None:
    from openpyxl import Workbook

    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(str(path))


def fetch_day(client: ZhiyunClient, day: str, out_dir: Path) -> dict:
    """拉一天的四张表 + 摘要 json。返回计数摘要（无客户名/金额明细）。"""
    out_dir.mkdir(parents=True, exist_ok=True)

    hk_ctrls = client.controls(WS_HUIKUAN)
    hk_opts = client.option_maps(hk_ctrls)
    cid_xiadan = client.id_by_name(hk_ctrls, REL_XIADAN)
    cid_mingxi = client.id_by_name(hk_ctrls, REL_HEXIAO_MINGXI)
    ws_sodline = client.datasource_of(WS_HUIKUAN, REL_SODLINE)
    if not cid_xiadan:
        raise FetchError(f"回款记录里找不到「{REL_XIADAN}」关联字段——智云表结构变了，停下别猜")

    hk_rows, hk_total = client.filter_rows_by_date(WS_HUIKUAN, F_HK["hexiao_date"], day)

    # ── ① 回款记录 ────────────────────────────────────────────
    hk_headers = [
        "回款记录ID", "核销日期", "到账日期", "到账金额/原币", "到账金额/本币",
        "手续费/原币", "原币币种", "回款类型", "核销状态", "开票客户", "rowid",
    ]
    hk_out: List[List[Any]] = []
    type_counts: Dict[str, int] = {}
    payments: List[dict] = []
    for row in hk_rows:
        ar = _plain(row.get(F_HK["ar"]))
        htype = _plain(row.get(F_HK["huikuan_type"]), hk_opts.get(F_HK["huikuan_type"]))
        type_counts[htype or "(空)"] = type_counts.get(htype or "(空)", 0) + 1
        rec = {
            "ar": ar,
            "hexiao_date": _plain(row.get(F_HK["hexiao_date"])),
            "arrival_date": _plain(row.get(F_HK["arrival_date"])),
            "amount_orig": _plain(row.get(F_HK["amount_orig"])),
            "amount_local": _plain(row.get(F_HK["amount_local"])),
            "fee": _plain(row.get(F_HK["fee"])),
            "currency": _plain(row.get(F_HK["currency"]), hk_opts.get(F_HK["currency"])),
            "huikuan_type": htype,
            "status": _plain(row.get(F_HK["status"]), hk_opts.get(F_HK["status"])),
            "customer": _plain(row.get(F_HK["customer_txt"])) or _plain(row.get(F_HK["customer_rel"])),
            "rowid": row.get("rowid") or "",
        }
        payments.append(rec)
        hk_out.append([rec[k] for k in (
            "ar", "hexiao_date", "arrival_date", "amount_orig", "amount_local",
            "fee", "currency", "huikuan_type", "status", "customer", "rowid")])

    day_tag = day.replace("-", "")
    write_xlsx(out_dir / f"回款记录_{day_tag}.xlsx", hk_headers, hk_out)

    # ── ② 下单栏（每笔 → SO + 交付额）+ ③ 同币种核销明细 ───────
    xd_out: List[List[Any]] = []
    mx_out: List[List[Any]] = []
    all_so: List[str] = []
    ars_without_orders: List[str] = []

    for rec in payments:
        rid = str(rec.get("rowid") or "")
        if not rid:
            ars_without_orders.append(rec["ar"])
            continue

        xd_rows, xd_ctrls = client.relation_rows(WS_HUIKUAN, rid, cid_xiadan)
        xd_names, xd_opts = client.name_map(xd_ctrls), client.option_maps(xd_ctrls)
        got_so = 0
        for r in xd_rows:
            v = pick_named(r, xd_names, xd_opts, XIADAN_COLS)
            so = extract_so(v.get("SO") or "")
            if not so:
                continue
            got_so += 1
            if so not in all_so:
                all_so.append(so)
            xd_out.append([
                rec["ar"], so, v.get("交付额/原币"), v.get("汇率"),
                v.get("结算币种"), v.get("订单名称"),
            ])
        if got_so == 0:
            ars_without_orders.append(rec["ar"])

        if cid_mingxi:
            mx_rows, mx_ctrls = client.relation_rows(WS_HUIKUAN, rid, cid_mingxi)
            mx_names, mx_opts = client.name_map(mx_ctrls), client.option_maps(mx_ctrls)
            for r in mx_rows:
                v = pick_named(r, mx_names, mx_opts, MINGXI_COLS)
                if (v.get("是否已撤销") or "").strip() == "是":
                    continue  # 撤销掉的核销不算数
                so = extract_so(v.get("订单NUM") or "")
                if not so:
                    print(
                        f"WARN: 核销明细有一行读不出 SO（AR={rec['ar']} 订单={v.get('订单名称')}）"
                        "，判定会把这笔挂起不会瞎填",
                        file=sys.stderr,
                    )
                if so and so not in all_so:
                    all_so.append(so)
                mx_out.append([
                    rec["ar"], v.get("核销日期") or rec["hexiao_date"],
                    v.get("本次核销金额"), v.get("本次核销金额本币"),
                    v.get("币种") or rec["currency"], v.get("汇率"), so, v.get("订单名称"),
                ])

    write_xlsx(
        out_dir / f"订单交付_{day_tag}.xlsx",
        ["回款记录ID", "SO", "交付额/原币", "汇率", "结算币种", "订单名称"],
        xd_out,
    )
    write_xlsx(
        out_dir / f"核销明细_{day_tag}.xlsx",
        ["回款记录NUM", "核销日期", "本次核销金额", "本次核销金额/本币",
         "币种", "汇率", "SO", "订单名称"],
        mx_out,
    )

    # ── ④ 订单明细（SO → SOD + 逐 SOD 交付额）─────────────────
    sod_out: List[List[Any]] = []
    so_without_sod: List[str] = []
    if ws_sodline and all_so:
        sl_ctrls = client.controls(ws_sodline)
        sl_names, sl_opts = client.name_map(sl_ctrls), client.option_maps(sl_ctrls)
        for so in all_so:
            try:
                hits = client.search_rows(ws_sodline, so)
            except Exception as e:
                print(f"WARN: 订单明细检索失败 SO={so}: {type(e).__name__}", file=sys.stderr)
                hits = []
            n = 0
            for r in hits:
                v = pick_named(r, sl_names, sl_opts, SODLINE_COLS)
                # 全文检索会带出订单名里含该串的别的单 → 必须精确过滤
                if extract_so(v.get("SO") or "") != so:
                    continue
                sod = (v.get("SOD") or "").strip()
                if not sod:
                    continue
                n += 1
                sod_out.append([so, sod, v.get("交付额/原币"), v.get("币种"), v.get("项目状态")])
            if n == 0:
                so_without_sod.append(so)
    elif all_so:
        print(f"WARN: 找不到「{REL_SODLINE}」表，SOD 取不到", file=sys.stderr)
        so_without_sod = list(all_so)

    write_xlsx(
        out_dir / f"订单明细_{day_tag}.xlsx",
        ["SO", "SOD", "交付额/原币", "币种", "项目状态"],
        sod_out,
    )

    summary = {
        "day": day,
        "架构": "单入口(回款记录·核销日期=T-1)+关联子表",
        "回款记录笔数": len(hk_rows),
        "回款记录_接口报总数": hk_total,
        "下单行数": len(xd_out),
        "涉及SO数": len(all_so),
        "核销明细行数": len(mx_out),
        "订单明细SOD行数": len(sod_out),
        "回款类型分布": type_counts,
        "无下单行的AR": ars_without_orders,
        "查不到SOD的SO": so_without_sod,
        "files": [
            f"回款记录_{day_tag}.xlsx",
            f"订单交付_{day_tag}.xlsx",
            f"核销明细_{day_tag}.xlsx",
            f"订单明细_{day_tag}.xlsx",
        ],
        "read_only": True,
    }
    (out_dir / f"取数摘要_{day_tag}.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return summary


def resolve_credentials(args) -> Tuple[str, str]:
    user = (args.user or os.environ.get("ZHIYUN_USER") or "").strip()
    pwd = (args.password or os.environ.get("ZHIYUN_PASS") or "").strip()
    if args.cookie_only:
        return "", ""
    if not user:
        user = input("智云账号（邮箱/手机）: ").strip()
    if not pwd:
        pwd = getpass.getpass("智云密码（不回显、不落盘）: ")
    if not user or not pwd:
        raise SystemExit("ERROR: 需要账号和密码（或改用 --cookie-only + MD_PSS_ID）")
    return user, pwd


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="智云只读取数（单入口：回款记录按核销日期 + 关联子表）"
    )
    ap.add_argument("--date", default="yesterday", help="YYYY-MM-DD / yesterday / today")
    ap.add_argument("--workspace", default="", help="技能工作区根（含 01_智云导出）")
    ap.add_argument("--out", default="", help="直接指定导出目录（优先于 workspace）")
    ap.add_argument("--base-url", default=os.environ.get("ZHIYUN_BASE", BASE_DEFAULT))
    ap.add_argument("--user", default="", help="账号；也可用环境变量 ZHIYUN_USER")
    ap.add_argument(
        "--password", default="",
        help="密码（不推荐写在命令行历史）；优先用 ZHIYUN_PASS 或交互 getpass",
    )
    ap.add_argument("--cookie-only", action="store_true", help="不登录，只用 MD_PSS_ID")
    ap.add_argument("--account-id", default=os.environ.get("ZHIYUN_ACCOUNT_ID", ""))
    ap.add_argument("--headed", action="store_true", help="有头浏览器登录（调试）")
    args = ap.parse_args(argv)

    day = resolve_date(args.date)
    if args.out:
        out_dir = Path(args.out)
    elif args.workspace:
        out_dir = Path(args.workspace) / "01_智云导出"
    else:
        out_dir = Path(__file__).resolve().parent.parent / "工作区" / "01_智云导出"

    cookie = (os.environ.get("MD_PSS_ID") or "").strip()
    account_id = (args.account_id or "").strip()

    if not cookie:
        user, pwd = resolve_credentials(args)
        print(f"正在登录智云 {args.base_url} …（密码不打印）")
        try:
            cookie, acct = login_with_password(args.base_url, user, pwd, headless=not args.headed)
        except LoginError as e:
            print(f"ERROR: 登录失败 — {e}", file=sys.stderr)
            return 2
        if acct and not account_id:
            account_id = acct
        del pwd
        print("登录成功，开始只读取数…")

    try:
        client = ZhiyunClient(args.base_url, cookie, account_id=account_id)
        try:
            client.controls(WS_HUIKUAN)
        except Exception as e:
            print(f"ERROR: 取字段失败（内网不通/无权限？）: {e}", file=sys.stderr)
            return 2
        summary = fetch_day(client, day, out_dir)
    finally:
        cookie = ""
        del cookie

    print("✅ 智云只读取数完成（未写系统）")
    print(f"📁 目录: {out_dir.resolve()}")
    print(
        f"   回款记录 {summary['回款记录笔数']} 笔 · 下单 {summary['下单行数']} 行"
        f"（{summary['涉及SO数']} 个 SO）· 核销明细 {summary['核销明细行数']} 行"
        f" · 订单明细 {summary['订单明细SOD行数']} 个 SOD"
    )
    print(f"   回款类型分布: {summary['回款类型分布']}")
    if summary["无下单行的AR"]:
        print(f"   ⚠ 有 {len(summary['无下单行的AR'])} 笔到账没抓到下单行，判定会报异常不会漏")
    if summary["查不到SOD的SO"]:
        print(f"   ⚠ 有 {len(summary['查不到SOD的SO'])} 个 SO 查不到 SOD，将退化成按 SO 匹配")
    print("👉 下一步：把盈亏/流转表副本放进 02_我的表副本/ 后说「跑昨天的核销」")
    return 0


if __name__ == "__main__":
    sys.exit(main())
