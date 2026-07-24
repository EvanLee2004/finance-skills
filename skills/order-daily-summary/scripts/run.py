#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""九点下单统计 CLI：date_window → fetch → summarize → write_report。"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date, datetime, time as dtime
from pathlib import Path

# 每日 9 点快照口径：抓取时刻晚于此（含宽限）→ 视为“晚跑”，提示昨日订单可能已被改期
SNAPSHOT_HOUR = 9
LATE_RUN_AFTER = dtime(9, 15)

# allow `python scripts/run.py` from skill root or any cwd
_SCRIPTS = Path(__file__).resolve().parent
_SKILL_ROOT = _SCRIPTS.parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from date_window import date_window, is_weekend  # noqa: E402
from fetch_orders import ZHIYUN_DEFAULTS, LoginError, fetch_orders  # noqa: E402
from summarize import (  # noqa: E402
    DEPT_DISPLAY_ORDER,
    dept_totals,
    filter_records_by_date_window,
    load_org_map_from_xlsx,
    load_records_from_order_xlsx,
    summarize_records,
)
from write_report import write_report  # noqa: E402


def _parse_today(s: str | None) -> date:
    if not s:
        return date.today()
    return datetime.strptime(s.strip()[:10], "%Y-%m-%d").date()


def load_config(config_path: Path | None) -> dict:
    cfg: dict = {
        "base_url": ZHIYUN_DEFAULTS["base_url"],
        "app_id": ZHIYUN_DEFAULTS["app_id"],
        "orders_worksheet_id": ZHIYUN_DEFAULTS["orders_worksheet_id"],
    }
    # env first for secrets
    env_user = os.environ.get("ZHIYUN_USER") or os.environ.get("ZHIYUN_USERNAME")
    env_pwd = os.environ.get("ZHIYUN_PASSWORD")
    if env_user:
        cfg["username"] = env_user
    if env_pwd:
        cfg["password"] = env_pwd

    path = config_path
    if path is None:
        candidate = _SKILL_ROOT / "config.local.json"
        if candidate.is_file():
            path = candidate
    if path is not None and Path(path).is_file():
        try:
            # utf-8-sig 兼容 Windows PowerShell 默认写出的 BOM，避免 json 解析炸
            file_cfg = json.loads(Path(path).read_text(encoding="utf-8-sig"))
        except (OSError, ValueError) as e:
            print(f"错误：无法读取配置文件 {path}：{e}", file=sys.stderr)
            sys.exit(2)
        if isinstance(file_cfg, dict):
            for k, v in file_cfg.items():
                if v not in (None, ""):
                    # env secrets win over file if already set for user/pass
                    if k in ("username", "password") and k in cfg and cfg[k]:
                        continue
                    cfg[k] = v
    return cfg


def _has_credentials(cfg: dict) -> bool:
    if cfg.get("md_pss_id"):
        return True
    return bool(cfg.get("username") and cfg.get("password"))


def _creds_help() -> str:
    example = _SKILL_ROOT / "config" / "config.local.example.json"
    local = _SKILL_ROOT / "config.local.json"
    return (
        "缺少智云登录凭据，无法抓数。\n"
        "请任选其一准备（密码不会回显到日志）：\n"
        f"  1) 复制 {example} → {local}，填写 username / password\n"
        "  2) 环境变量 export ZHIYUN_USER='邮箱' 与 ZHIYUN_PASSWORD='密码'\n"
        "注意：config.local.json 已 gitignore，切勿提交或打进技能包。\n"
        "本技能需能访问内网 http://192.168.10.167:18880 。"
    )


def _late_run_warning(today: date, fetch_time: datetime) -> str | None:
    """跑的是“今天的实时窗口”且晚于 9:15 → 提示昨日订单可能已被改期。

    智云「下单日期」可被编辑，昨天的订单当天会被改期/改单号，导致“昨日下单额”
    随当天推移而变。9 点定时跑即为当日快照；晚跑得到的数字可能与 9 点不同。
    仅当 today == 实际今天（在跑活数据）才提示；--today 回溯历史日不提示。
    """
    if today != date.today():
        return None
    if fetch_time.time() <= LATE_RUN_AFTER:
        return None
    return (
        f"本次于 {fetch_time.strftime('%H:%M')} 实时抓取，晚于每日 {SNAPSHOT_HOUR:02d}:00 快照口径；"
        "智云「下单日期」可被改期/改单号，晚跑数字常小于 9 点（例：大单改到今天则昨日报表少一截）。"
        "要对齐「九点下单/旧 exe」：请在 09:00–09:15 跑本技能，或用 --from-xlsx 喂 9 点人工导出的下单表。"
    )


def _force_utf8_stdio() -> None:
    """Windows 控制台/重定向下中文乱码兜底：把 stdout/stderr 切到 UTF-8。"""
    os.environ.setdefault("PYTHONUTF8", "1")
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            pass


def main(argv: list[str] | None = None) -> int:
    _force_utf8_stdio()
    parser = argparse.ArgumentParser(
        description="九点下单统计：登录智云抓下单表 → 组织架构归部门 → 输出万元汇总 xlsx"
    )
    parser.add_argument(
        "--out",
        required=True,
        help="输出目录（绝对路径更稳）",
    )
    parser.add_argument(
        "--today",
        default=None,
        help="模拟运行日 YYYY-MM-DD（默认今天）",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="config.local.json 路径（默认 skill 根目录下 config.local.json）",
    )
    parser.add_argument(
        "--detail",
        action="store_true",
        help="额外输出明细 sheet",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="周末也强制跑（默认周末可跑，本开关保留兼容）",
    )
    parser.add_argument(
        "--org",
        default=None,
        help="组织架构 xlsx 路径（默认 config/销售组织架构.xlsx）",
    )
    parser.add_argument(
        "--from-xlsx",
        default=None,
        help="离线模式：跳过智云，直接读本地「下单」导出 xlsx（与旧 exe 同口径；"
        "用于 9 点人工导出后汇总，或与 exe 对数）。不需账号。",
    )
    parser.add_argument(
        "--no-date-filter",
        action="store_true",
        help="配合 --from-xlsx：不过滤日期窗口，导出表里有什么算什么"
        "（导出已是单日时可开）",
    )
    args = parser.parse_args(argv)

    today = _parse_today(args.today)
    start, end = date_window(today)
    if is_weekend(today) and not args.force:
        # still allow run; task book says 若仍触发 with weekend window; --force for default scripts
        pass

    org_path = Path(args.org) if args.org else (_SKILL_ROOT / "config" / "销售组织架构.xlsx")
    if not org_path.is_file():
        print(f"错误：未找到组织架构文件：{org_path}", file=sys.stderr)
        return 2

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"运行日={today.isoformat()} 窗口={start.isoformat()}～{end.isoformat()}")
    offline = bool(args.from_xlsx)
    late_warning: str | None = None
    source_note = ""

    if offline:
        xlsx_path = Path(args.from_xlsx)
        print(f"离线模式：读取本地下单表 {xlsx_path}")
        try:
            raw = load_records_from_order_xlsx(xlsx_path)
        except (OSError, ValueError) as e:
            print(f"读本地下单表失败：{e}", file=sys.stderr)
            return 4
        if args.no_date_filter:
            records = raw
            source_note = f"离线导出 {xlsx_path.name}（未滤日期窗口，共 {len(raw)} 行）"
        else:
            records = filter_records_by_date_window(raw, start, end)
            source_note = (
                f"离线导出 {xlsx_path.name}（窗口内 {len(records)}/{len(raw)} 行）"
            )
        fetch_time = datetime.now()
        # 离线=吃固定导出，不算「晚跑实时抓」
        late_warning = None
    else:
        cfg_path = Path(args.config) if args.config else None
        cfg = load_config(cfg_path)
        if not _has_credentials(cfg):
            print(_creds_help(), file=sys.stderr)
            return 2
        print("正在登录智云并抓取下单…（需内网）")
        try:
            records = fetch_orders(start, end, cfg)
        except LoginError as e:
            print(f"登录失败：{e}", file=sys.stderr)
            print("请检查账号密码与内网连通，勿在对话中粘贴密码。", file=sys.stderr)
            return 3
        except Exception as e:  # noqa: BLE001
            print(f"抓数失败：{type(e).__name__}: {e}", file=sys.stderr)
            return 4
        fetch_time = datetime.now()  # 抓取时刻 = 数据快照的“截至”时间
        late_warning = _late_run_warning(today, fetch_time)
        source_note = "智云实时 GetFilterRows"

    org_map = load_org_map_from_xlsx(org_path)
    result = summarize_records(records, org_map)

    stamp = fetch_time.strftime("%Y%m%d_%H%M%S")
    out_file = out_dir / f"下单数据_{start.isoformat()}_{end.isoformat()}_{stamp}.xlsx"
    write_report(
        result,
        out_file,
        window_start=start,
        window_end=end,
        api_row_count=len(records),
        include_detail=True if offline else bool(args.detail),  # 离线默认带明细+SO 便于对 exe
        data_asof=fetch_time,
        late_warning=late_warning,
        extra_log={"取数来源": source_note},
    )

    print(f"明细行数={result.detail_row_count}")
    print(f"总计万元={result.grand_total_wan}")
    dtot = dept_totals(result)
    breakdown = "  ".join(f"{name}={dtot.get(name, 0.0):.2f}" for name in DEPT_DISPLAY_ORDER)
    print(f"分部门万元：{breakdown}")
    print(f"金额字段={result.amount_field_used or '（未解析）'}")
    print(f"取数来源={source_note}")
    if result.unmatched_sales:
        print(f"未匹配销售({len(result.unmatched_sales)})：{'、'.join(result.unmatched_sales)}")
        print("请更新 config/销售组织架构.xlsx 后重跑。")
    else:
        print("未匹配销售=无")
    if offline:
        print(f"数据来自本地导出（固定快照），输出={out_file}")
    else:
        print(f"数据截至={fetch_time.strftime('%Y-%m-%d %H:%M:%S')}（智云为实时数据，本表为该时刻快照）")
        if late_warning:
            print(f"⚠ {late_warning}")
        print(f"输出={out_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
