#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""九点下单统计 CLI：date_window → fetch → summarize → write_report。"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date, datetime
from pathlib import Path

# allow `python scripts/run.py` from skill root or any cwd
_SCRIPTS = Path(__file__).resolve().parent
_SKILL_ROOT = _SCRIPTS.parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from date_window import date_window, is_weekend  # noqa: E402
from fetch_orders import ZHIYUN_DEFAULTS, LoginError, fetch_orders  # noqa: E402
from summarize import load_org_map_from_xlsx, summarize_records  # noqa: E402
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
            file_cfg = json.loads(Path(path).read_text(encoding="utf-8"))
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


def main(argv: list[str] | None = None) -> int:
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
    args = parser.parse_args(argv)

    today = _parse_today(args.today)
    start, end = date_window(today)
    if is_weekend(today) and not args.force:
        # still allow run; task book says 若仍触发 with weekend window; --force for default scripts
        pass

    cfg_path = Path(args.config) if args.config else None
    cfg = load_config(cfg_path)
    if not _has_credentials(cfg):
        print(_creds_help(), file=sys.stderr)
        return 2

    org_path = Path(args.org) if args.org else (_SKILL_ROOT / "config" / "销售组织架构.xlsx")
    if not org_path.is_file():
        print(f"错误：未找到组织架构文件：{org_path}", file=sys.stderr)
        return 2

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"运行日={today.isoformat()} 窗口={start.isoformat()}～{end.isoformat()}")
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

    org_map = load_org_map_from_xlsx(org_path)
    result = summarize_records(records, org_map)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = out_dir / f"下单数据_{start.isoformat()}_{end.isoformat()}_{stamp}.xlsx"
    write_report(
        result,
        out_file,
        window_start=start,
        window_end=end,
        api_row_count=len(records),
        include_detail=bool(args.detail),
    )

    print(f"明细行数={result.detail_row_count}")
    print(f"总计万元={result.grand_total_wan}")
    print(f"金额字段={result.amount_field_used or '（未解析）'}")
    if result.unmatched_sales:
        print(f"未匹配销售({len(result.unmatched_sales)})：{'、'.join(result.unmatched_sales)}")
        print("请更新 config/销售组织架构.xlsx 后重跑。")
    else:
        print("未匹配销售=无")
    print(f"输出={out_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
