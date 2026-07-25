# -*- coding: utf-8 -*-
"""离线 --from-xlsx：读导出表 + 汇总。有金标文件则验 4.26。"""
from __future__ import annotations

import sys
import os
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from summarize import (  # noqa: E402
    load_org_map_from_xlsx,
    load_records_from_order_xlsx,
    summarize_records,
)

# 金标：某次 08:50 人工导出的下单表（真实数据，不进仓库）。
# 位置可用环境变量覆盖，方便换机器/换人跑；默认找运行者的 Downloads。
# ⚠ 不能只判 is_file()：文件在、但**读不了**（macOS 隐私授权、权限、被占用）时
#    旧写法会让整套测试硬失败。公开仓里谁 clone 都跑不了 = 没人跑测试。
GOLD = Path(
    os.environ.get("ORDER_GOLD_XLSX")
    or (Path.home() / "Downloads" / "下单汇总_20260724_085036.xlsx")
)
ORG = Path(__file__).resolve().parents[1] / "config" / "销售组织架构.xlsx"


def _gold_readable() -> bool:
    """文件既在、又真的读得开，才算能跑金标。"""
    try:
        with GOLD.open("rb"):
            return True
    except OSError:
        return False


def test_load_and_sum_synthetic(tmp_path):
    from openpyxl import Workbook

    p = tmp_path / "下单.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "下单"
    ws.append(["销售", "SO", "订单名称", "下单日期", "下单预估额/本币"])
    ws.append(["甲", "SO1", "项目A", "2026-07-23", 10000])
    ws.append(["乙", "SO2", "项目B", "2026-07-23", 5000])
    wb.save(p)
    recs = load_records_from_order_xlsx(p)
    assert len(recs) == 2
    org = {"甲": "本地化", "乙": "数据"}
    r = summarize_records(recs, org)
    assert r.detail_row_count == 2
    assert abs(r.grand_total_wan - 1.5) < 0.001
    assert r.detail_rows[0].get("SO") == "SO1"


@pytest.mark.skipif(not _gold_readable(), reason="本机无 08:50 金标导出（或读不了）")
def test_gold_085036_is_4_26_wan():
    recs = load_records_from_order_xlsx(GOLD)
    assert len(recs) == 11
    org = load_org_map_from_xlsx(ORG)
    r = summarize_records(recs, org)
    assert r.detail_row_count == 11
    assert abs(r.grand_total_wan - 4.26) < 0.01
    sos = {d.get("SO") for d in r.detail_rows}
    assert "SO26070435" in sos  # 于占国 Apier WS 2.15 万那单
