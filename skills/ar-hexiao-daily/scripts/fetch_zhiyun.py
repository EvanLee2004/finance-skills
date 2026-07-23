#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第 1 步（可选）：智云只读取数 → 写入工作区 01_智云导出/

设计：
  - 账号密码**每次运行时提供**，绝不写进代码 / config / git
  - 只读：GetFilterRows / getRowRelationRows / getWorksheetInfo，无写接口
  - 三通道：回款记录(入口) + 回款核销对账(整笔/自动) + 关联子表(预存)
  - 分笔在系统无逐单金额 → 只出现在回款记录里，留给 classify hold

用法（任选一种提供账密）：
  export ZHIYUN_USER='xxx@besteasy.com'
  export ZHIYUN_PASS='****'          # 勿写进文件
  python3 scripts/fetch_zhiyun.py --date 2026-07-21 --workspace 工作区/

  或交互：
  python3 scripts/fetch_zhiyun.py --date yesterday --workspace 工作区/
  # 提示输入账号密码（密码 getpass 不回显）

  已有 cookie（调试）：
  export MD_PSS_ID='...'
  python3 scripts/fetch_zhiyun.py --date 2026-07-21 --cookie-only

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
from typing import Any, Dict, List, Optional, Tuple

# ── 常量（表 ID / 字段 ID 来自 2026-07-09/22 勘探，非密钥）────────────────
BASE_DEFAULT = "http://192.168.10.167:18880"
APP_ID = "6ff4fb2e-e68c-4ee9-83a0-836de8f72c11"

WS_HUIKUAN = "6555d2b1f9460e517040ba6c"  # 回款记录
WS_DUIZHANG = "66540a9f4a2483fd8625343e"  # 回款核销对账
WS_MINGXI = "689954801b702f9906a29504"  # 同币种核销明细信息（直查可能 0 行）

# 回款记录
F_HK = {
    "arrival_date": "6555d2b1f9460e517040ba70",
    "amount_orig": "6555d2b1f9460e517040ba71",
    "currency": "663c495f4a2483fd86249c90",
    "ar": "663c4b8a4a2483fd86249ca8",
    "amount_local": "663c4d354a2483fd86249cbb",
    "status": "663ca2204a2483fd8624a107",
    "xiadan": "663daf0f4a2483fd8624a5b9",  # 下单 关联
    "hexiao_date": "664c571b4a2483fd86250fa5",
    "huikuan_type": "66b1f6986429811e5f304a0b",
    "fee": "675119ab327314202700730a",
    "customer_rel": "676a3314327314202700e7fc",  # 开票客户 关联
    "customer_txt": "676b6d01327314202700ed59",  # 开票客户 文本
    "mingxi_rel": "689959a11b702f9906a29563",  # 订单同币种核销明细信息
}

# 回款核销对账
F_DZ = {
    "so": "65d70c96150d8ec2b4ec8478",
    "sod": "663c35a64a2483fd86249a23",
    "deliver_orig": "65f156ee82758c90eee78702",
    "deliver_local": "65f15c4c82758c90eee7875b",
    "order_name": "66540eff4a2483fd86253546",
    "customer": "6654119b4a2483fd86253580",
    "ar": "6654120c4a2483fd86253587",
    "hexiao_date": "665417fe4a2483fd862535d9",
    "status": "67c12e8fa82e575f4c9bc285",
    "settle": "67c13453a82e575f4c9bc382",
    "amount_orig": "67c14059a82e575f4c9bc453",
    "amount_local": "67c14059a82e575f4c9bc454",
    "shoukuan_date": "68230784c84cb09d26c27f1c",
    "currency": "65a9d21bdd2dc6df7283c2dc",
}

# 同币种核销明细（关联子表行）
F_MX = {
    "hx_num": "689959a11b702f9906a29561",
    "ar_rel": "689959a11b702f9906a29562",
    "order_rel": "689959a11b702f9906a29564",
    "hexiao_date": "689959a11b702f9906a29567",
    "amount": "689959a11b702f9906a29568",
    "currency": "689959a11b702f9906a29569",
    "rate": "689959a11b702f9906a2956a",
    "amount_local": "689959a11b702f9906a2956b",
    "status": "689959a11b702f9906a2956c",
    "order_name": "68ac04c31b702f9906a2fba1",
}


class LoginError(RuntimeError):
    pass


class FetchError(RuntimeError):
    pass


def _plain(v: Any, options: Optional[Dict[str, str]] = None) -> str:
    """明道云单元格 → 纯文本（与看板/旧 fetch 同思路）。"""
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


def resolve_date(s: str) -> str:
    s = (s or "").strip().lower()
    if s in ("yesterday", "t-1", "昨天"):
        return (date.today() - timedelta(days=1)).isoformat()
    if s in ("today", "今天"):
        return date.today().isoformat()
    # YYYY-MM-DD
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
    def __init__(
        self,
        base: str,
        cookie: str,
        account_id: str = "",
        page_size: int = 200,
    ):
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

    def post(self, path: str, body: dict, timeout: int = 90) -> dict:
        # 大小写：GetFilterRows 用文档大写路径；部分环境小写也能通，优先文档写法
        url = f"{self.base}/wwwapi/{path.lstrip('/')}"
        r = self.session.post(url, headers=self.headers, json=body, timeout=timeout)
        r.raise_for_status()
        j = r.json()
        # 明道云常见：{state, data:{data,count}} 或 data 直接是列表
        if isinstance(j, dict) and "data" in j:
            return j["data"] if j["data"] is not None else {}
        return j

    def get_fields(self, worksheet_id: str) -> List[dict]:
        # 小写路径与旧 fetch 兼容
        info = self.post(
            "worksheet/getWorksheetInfo",
            {"worksheetId": worksheet_id, "appId": APP_ID, "getTemplate": True},
        )
        controls = (info.get("template") or {}).get("controls") or info.get("controls") or []
        out = []
        for c in controls:
            opts = [
                {"key": o.get("key"), "value": o.get("value")}
                for o in (c.get("options") or [])
            ]
            out.append(
                {
                    "id": c.get("controlId") or c.get("id"),
                    "name": c.get("controlName") or c.get("name") or "",
                    "type": c.get("type"),
                    "options": opts,
                }
            )
        return out

    def option_map(self, fields: List[dict], control_id: str) -> Dict[str, str]:
        for f in fields:
            if f["id"] == control_id:
                return {str(o["key"]): str(o["value"]) for o in f.get("options") or []}
        return {}

    def filter_rows_by_date(
        self,
        worksheet_id: str,
        date_control_id: str,
        day: str,
        page_size: Optional[int] = None,
    ) -> Tuple[List[dict], int]:
        """按日期字段筛单日（filterType 11 + dateRange 18 已实测）。"""
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
            if batch is None:
                batch = []
            if not isinstance(batch, list):
                batch = []
            total = int(d.get("count") or total or 0) if isinstance(d, dict) else total
            rows.extend(batch)
            if not batch or (total and len(rows) >= total) or len(batch) < ps:
                break
            page += 1
            if page > 500:
                raise FetchError(f"翻页超过 500，worksheet={worksheet_id}")
        return rows, total or len(rows)

    def relation_rows(
        self, worksheet_id: str, row_id: str, control_id: str, page_size: int = 100
    ) -> List[dict]:
        rows: List[dict] = []
        page = 1
        while True:
            d = self.post(
                "worksheet/getRowRelationRows",
                {
                    "worksheetId": worksheet_id,
                    "rowId": row_id,
                    "controlId": control_id,
                    "pageIndex": page,
                    "pageSize": page_size,
                    "appId": APP_ID,
                    "getWorksheet": True,
                },
            )
            batch = []
            if isinstance(d, dict):
                batch = d.get("data") or d.get("rows") or []
                if isinstance(batch, dict):
                    batch = batch.get("data") or []
            if not isinstance(batch, list):
                batch = []
            rows.extend(batch)
            if not batch or len(batch) < page_size:
                break
            page += 1
            if page > 50:
                break
        return rows


def _so_from_xiadan_row(row: dict) -> Tuple[str, Optional[float]]:
    """从下单关联行尽量抽出 SO 与金额。"""
    # 常见：name 字段带 SO；或 control 里有 SO
    so = _plain(row.get("name")) or ""
    # 有时 name 是「SO2606… 某项目」
    for token in so.replace("\n", " ").split():
        if token.upper().startswith("SO") and len(token) >= 6:
            so = token
            break
    amount = None
    for k, v in row.items():
        if k in ("rowid", "sid", "allowedit", "allowdelete"):
            continue
        # 启发式：数值字段
        try:
            if isinstance(v, (int, float)):
                amount = float(v)
        except Exception:
            pass
        s = str(v) if v is not None else ""
        if s.upper().startswith("SO") and len(s) >= 8 and " " not in s:
            so = s
    return so, amount


def pair_yucun_so(
    detail_rows: List[dict], xiadan_rows: List[dict], fields_mx_opts: Dict[str, Dict[str, str]]
) -> List[dict]:
    """预存明细 × 下单栏：按金额配对补 SO（无 SOD）。金额对不上则 SO 空。"""
    xiadan_pool = []
    for xr in xiadan_rows:
        so, amt = _so_from_xiadan_row(xr)
        # 尝试从关联对象字段再挖
        if not so:
            so = _plain(xr.get("sourcevalue") or xr.get("name"))
        xiadan_pool.append({"so": so, "amount": amt, "raw": xr, "used": False})

    out = []
    for dr in detail_rows:
        amount = _plain(dr.get(F_MX["amount"]))
        try:
            amt_f = float(amount) if amount not in ("", None) else None
        except ValueError:
            amt_f = None
        so = ""
        if amt_f is not None:
            for xp in xiadan_pool:
                if xp["used"]:
                    continue
                if xp["amount"] is not None and abs(xp["amount"] - amt_f) < 0.02:
                    so = xp["so"]
                    xp["used"] = True
                    break
        # 订单NUM 关联有时带 name=SO
        if not so:
            so = _plain(dr.get(F_MX["order_rel"]))
            for token in so.replace("、", " ").split():
                if token.upper().startswith("SO"):
                    so = token
                    break
        out.append(
            {
                "ar": "",  # 调用方填
                "hexiao_date": _plain(dr.get(F_MX["hexiao_date"])),
                "amount": amount,
                "currency": _plain(dr.get(F_MX["currency"])),
                "rate": _plain(dr.get(F_MX["rate"])),
                "so": so,
                "order_name": _plain(dr.get(F_MX["order_name"])),
            }
        )
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


def fetch_day(
    client: ZhiyunClient,
    day: str,
    out_dir: Path,
    *,
    skip_yucun_details: bool = False,
) -> dict:
    """拉一天三通道数据，写 Excel + 摘要 json。返回计数摘要（无敏感明细）。"""
    out_dir.mkdir(parents=True, exist_ok=True)

    fields_hk = client.get_fields(WS_HUIKUAN)
    opts_type = client.option_map(fields_hk, F_HK["huikuan_type"])
    opts_status = client.option_map(fields_hk, F_HK["status"])
    opts_cur = client.option_map(fields_hk, F_HK["currency"])

    hk_rows, hk_total = client.filter_rows_by_date(
        WS_HUIKUAN, F_HK["hexiao_date"], day
    )
    dz_rows, dz_total = client.filter_rows_by_date(
        WS_DUIZHANG, F_DZ["hexiao_date"], day
    )

    # —— 回款记录 xlsx ——
    hk_headers = [
        "回款记录ID",
        "核销日期",
        "到账日期",
        "到账金额/原币",
        "到账金额/本币",
        "手续费/原币",
        "原币币种",
        "回款类型",
        "核销状态",
        "开票客户",
        "rowid",
    ]
    hk_out_rows = []
    type_counts: Dict[str, int] = {}
    yucun_candidates: List[dict] = []

    for row in hk_rows:
        ar = _plain(row.get(F_HK["ar"]))
        htype = _plain(row.get(F_HK["huikuan_type"]), opts_type)
        type_counts[htype or "(空)"] = type_counts.get(htype or "(空)", 0) + 1
        customer = _plain(row.get(F_HK["customer_txt"])) or _plain(
            row.get(F_HK["customer_rel"])
        )
        rec = {
            "ar": ar,
            "hexiao_date": _plain(row.get(F_HK["hexiao_date"])),
            "arrival_date": _plain(row.get(F_HK["arrival_date"])),
            "amount_orig": _plain(row.get(F_HK["amount_orig"])),
            "amount_local": _plain(row.get(F_HK["amount_local"])),
            "fee": _plain(row.get(F_HK["fee"])),
            "currency": _plain(row.get(F_HK["currency"]), opts_cur),
            "huikuan_type": htype,
            "status": _plain(row.get(F_HK["status"]), opts_status),
            "customer": customer,
            "rowid": row.get("rowid") or "",
        }
        hk_out_rows.append(
            [
                rec["ar"],
                rec["hexiao_date"],
                rec["arrival_date"],
                rec["amount_orig"],
                rec["amount_local"],
                rec["fee"],
                rec["currency"],
                rec["huikuan_type"],
                rec["status"],
                rec["customer"],
                rec["rowid"],
            ]
        )
        if "预存" in (htype or "") or "预收" in (htype or ""):
            yucun_candidates.append(rec)

    day_tag = day.replace("-", "")
    hk_path = out_dir / f"回款记录_{day_tag}.xlsx"
    write_xlsx(hk_path, hk_headers, hk_out_rows)

    # —— 对账表 xlsx ——
    fields_dz = client.get_fields(WS_DUIZHANG)
    opts_dz_status = client.option_map(fields_dz, F_DZ["status"])
    dz_headers = [
        "回款记录ID",
        "SO",
        "SOD",
        "回款额/原币",
        "回款额/本币",
        "交付额/原币",
        "交付额/本币",
        "核销日期",
        "回款日期",
        "客户名称",
        "结算周期",
        "核销状态",
        "下单币种",
        "订单名称",
    ]
    dz_out = []
    for row in dz_rows:
        dz_out.append(
            [
                _plain(row.get(F_DZ["ar"])),
                _plain(row.get(F_DZ["so"])),
                _plain(row.get(F_DZ["sod"])),
                _plain(row.get(F_DZ["amount_orig"])),
                _plain(row.get(F_DZ["amount_local"])),
                _plain(row.get(F_DZ["deliver_orig"])),
                _plain(row.get(F_DZ["deliver_local"])),
                _plain(row.get(F_DZ["hexiao_date"])),
                _plain(row.get(F_DZ["shoukuan_date"])),
                _plain(row.get(F_DZ["customer"])),
                _plain(row.get(F_DZ["settle"])),
                _plain(row.get(F_DZ["status"]), opts_dz_status),
                _plain(row.get(F_DZ["currency"])),
                _plain(row.get(F_DZ["order_name"])),
            ]
        )
    dz_path = out_dir / f"回款核销对账_{day_tag}.xlsx"
    write_xlsx(dz_path, dz_headers, dz_out)

    # —— 预存：关联子表 ——
    mx_headers = [
        "回款记录NUM",
        "核销日期",
        "本次核销金额",
        "币种",
        "汇率",
        "SO",
        "订单名称",
    ]
    mx_out: List[List[Any]] = []
    yucun_detail_count = 0
    yucun_with_so = 0
    if not skip_yucun_details:
        for rec in yucun_candidates:
            rid = rec.get("rowid")
            if not rid:
                continue
            try:
                details = client.relation_rows(
                    WS_HUIKUAN, str(rid), F_HK["mingxi_rel"]
                )
            except Exception as e:
                print(f"WARN: 预存明细子表失败 AR={rec.get('ar')}: {type(e).__name__}", file=sys.stderr)
                details = []
            try:
                xiadan = client.relation_rows(WS_HUIKUAN, str(rid), F_HK["xiadan"])
            except Exception:
                xiadan = []
            paired = pair_yucun_so(details, xiadan, {})
            for p in paired:
                p["ar"] = rec["ar"]
                yucun_detail_count += 1
                if p.get("so"):
                    yucun_with_so += 1
                mx_out.append(
                    [
                        rec["ar"],
                        p.get("hexiao_date") or rec.get("hexiao_date"),
                        p.get("amount"),
                        p.get("currency") or rec.get("currency"),
                        p.get("rate"),
                        p.get("so") or "",
                        p.get("order_name") or "",
                    ]
                )

    mx_path = out_dir / f"核销明细_{day_tag}.xlsx"
    write_xlsx(mx_path, mx_headers, mx_out)

    # 摘要（无客户名/金额明细）
    summary = {
        "day": day,
        "huikuan_rows": len(hk_rows),
        "huikuan_count_api": hk_total,
        "duizhang_rows": len(dz_rows),
        "duizhang_count_api": dz_total,
        "yucun_ar_count": len(yucun_candidates),
        "yucun_detail_rows": yucun_detail_count,
        "yucun_detail_with_so": yucun_with_so,
        "type_counts": type_counts,
        "files": [hk_path.name, dz_path.name, mx_path.name],
        "read_only": True,
    }
    sum_path = out_dir / f"取数摘要_{day_tag}.json"
    sum_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
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
        description="智云只读取数 → 01_智云导出（账密运行时提供，不写死）"
    )
    ap.add_argument("--date", default="yesterday", help="YYYY-MM-DD / yesterday / today")
    ap.add_argument("--workspace", default="", help="技能工作区根（含 01_智云导出）")
    ap.add_argument("--out", default="", help="直接指定导出目录（优先于 workspace）")
    ap.add_argument("--base-url", default=os.environ.get("ZHIYUN_BASE", BASE_DEFAULT))
    ap.add_argument("--user", default="", help="账号；也可用环境变量 ZHIYUN_USER")
    ap.add_argument(
        "--password",
        default="",
        help="密码（不推荐写在命令行历史）；优先用 ZHIYUN_PASS 或交互 getpass",
    )
    ap.add_argument("--cookie-only", action="store_true", help="不登录，只用 MD_PSS_ID")
    ap.add_argument("--account-id", default=os.environ.get("ZHIYUN_ACCOUNT_ID", ""))
    ap.add_argument("--headed", action="store_true", help="有头浏览器登录（调试）")
    ap.add_argument("--skip-yucun-details", action="store_true", help="跳过预存子表（更快）")
    args = ap.parse_args(argv)

    day = resolve_date(args.date)
    if args.out:
        out_dir = Path(args.out)
    elif args.workspace:
        out_dir = Path(args.workspace) / "01_智云导出"
    else:
        # 默认：技能自带工作区
        skill = Path(__file__).resolve().parent.parent
        out_dir = skill / "工作区" / "01_智云导出"

    cookie = (os.environ.get("MD_PSS_ID") or "").strip()
    account_id = (args.account_id or "").strip()

    if not args.cookie_only or not cookie:
        if not cookie:
            user, pwd = resolve_credentials(args)
            print(f"正在登录智云 {args.base_url} …（密码不打印）")
            try:
                cookie, acct = login_with_password(
                    args.base_url, user, pwd, headless=not args.headed
                )
            except LoginError as e:
                print(f"ERROR: 登录失败 — {e}", file=sys.stderr)
                return 2
            if acct and not account_id:
                account_id = acct
            # 立刻丢掉明文引用
            del pwd
            print("登录成功，开始只读取数…")

    if not cookie:
        print("ERROR: 无 md_pss_id", file=sys.stderr)
        return 2

    try:
        client = ZhiyunClient(args.base_url, cookie, account_id=account_id)
        # 连通性探活
        try:
            client.get_fields(WS_HUIKUAN)
        except Exception as e:
            print(f"ERROR: 取数字段失败（内网/权限？）: {e}", file=sys.stderr)
            return 2

        summary = fetch_day(
            client, day, out_dir, skip_yucun_details=args.skip_yucun_details
        )
    finally:
        # 会话密钥不驻留
        cookie = ""
        del cookie

    print("✅ 智云只读取数完成（未写系统）")
    print(f"📁 目录: {out_dir.resolve()}")
    print(
        f"   回款记录 {summary['huikuan_rows']} 笔 · "
        f"对账 {summary['duizhang_rows']} 行 · "
        f"预存AR {summary['yucun_ar_count']} · "
        f"预存明细 {summary['yucun_detail_rows']}（含SO {summary['yucun_detail_with_so']}）"
    )
    print(f"   回款类型分布: {summary['type_counts']}")
    print(f"   文件: {', '.join(summary['files'])}")
    print("👉 下一步：把盈亏/流转表副本放进 02_我的表副本/ 后说「跑昨天的核销」")
    return 0


if __name__ == "__main__":
    sys.exit(main())
