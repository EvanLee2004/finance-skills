# -*- coding: utf-8 -*-
"""迭代 v2 新增：流转表三键匹配 / per-AR 合计校验 / 案例ID / 真重判 / 源文件只读。"""
import datetime as dt
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from conftest import FIXTURE, LEDGER_FULL, TEST_DATA  # noqa: E402

import flow_ledger as FL  # noqa: E402
import classify_hexiao as C  # noqa: E402
import rescan_holds as R  # noqa: E402
import verify_sources as V  # noqa: E402

FLOW_HK = TEST_DATA / "步骤7_回填" / "到账流转表_汇款7月_副本.xlsx"
FLOW_WX = TEST_DATA / "步骤7_回填" / "到账流转表_微信全年_副本.xlsx"


# ---------- 公司名归一 ----------

def test_name_similar_ignores_suffix_and_punct():
    assert FL.name_similar("追觅创新科技（苏州）有限公司", "追觅创新科技(苏州)")
    assert FL.name_similar("交通银行股份有限公司北京市分行", "交通银行北京市分行")
    assert not FL.name_similar("北京多语", "深圳正浩")


def test_name_similar_empty_is_false():
    assert not FL.name_similar("", "任意公司")


# ---------- 三键匹配 ----------

def _flow(rows):
    f = FL.FlowLedger()
    f.rows = rows
    return f


def _row(**kw):
    base = {
        "file": "t.xlsx", "sheet": "S", "row_no": 2, "date": dt.date(2026, 7, 14),
        "payer": "某某科技有限公司", "amount": 100.0, "order_cell": "", "form": "汇款",
        "updated": "", "registered": "",
    }
    base.update(kw)
    return base


def test_match_three_keys_hit_one():
    f = _flow([_row()])
    hit = f.match(dt.date(2026, 7, 14), 100.0, "某某科技")
    assert hit["hits"] == 1 and hit["matched_by"] == "三键"


def test_match_uses_gross_when_fee_present():
    """微信/支付宝：流转表登的是含手续费的客户支付总额，净额匹配不上要用合计。"""
    f = _flow([_row(amount=300.0)])
    hit = f.match(dt.date(2026, 7, 14), 298.38, "某某科技", fee=1.62)
    assert hit["hits"] == 1 and hit["matched_by"] == "三键(含手续费)"


def test_match_zero_is_e0_signal():
    f = _flow([_row()])
    assert f.match(dt.date(2026, 7, 14), 999.0, "某某科技")["hits"] == 0


def test_match_multi_is_e12_signal():
    f = _flow([_row(row_no=2), _row(row_no=3)])
    assert f.match(dt.date(2026, 7, 14), 100.0, "某某科技")["hits"] == 2


def test_match_name_mismatch_is_weak_not_silent():
    """名字对不上但日期金额命中：仍返回，但标注需人工确认，不许当强命中。"""
    f = _flow([_row(payer="完全无关公司")])
    hit = f.match(dt.date(2026, 7, 14), 100.0, "某某科技")
    assert hit["hits"] == 1 and hit["matched_by"] == "日期+金额(名字不符)"


def test_suggest_order_cell_appends_without_dup():
    assert FL.FlowLedger.suggest_order_cell(["SO2"], "SO1") == "SO1\nSO2"
    assert FL.FlowLedger.suggest_order_cell(["SO1"], "SO1") == "SO1"


# ---------- 流转表三态派生 ----------

def test_derive_flow_status_tri_state():
    # 新词表：ready=行在表里她更得动；wait=没交付进表/异常，今天更不了
    assert FL.derive_flow_status(["ready", "ready"]) == "是"
    assert FL.derive_flow_status(["ready", "wait"]) == "部分"
    assert FL.derive_flow_status(["wait", "wait"]) == ""
    assert FL.derive_flow_status([]) == ""
    # 兼容旧入参：历史上传 bucket，"auto" 仍按 ready 计
    assert FL.derive_flow_status(["auto", "auto"]) == "是"
    assert FL.derive_flow_status(["auto", "hold"]) == "部分"


def test_flow_status_policy_no_formula_when_empty():
    p = FL.flow_status_policy("")
    assert p["填法"] == "留空"
    assert "不要公式" in p["公式策略"]
    p2 = FL.flow_status_policy("部分")
    assert "红" in p2["颜色标注"]


# ---------- per-AR 核销合计校验 ----------
# 2026-07-23：check_ar_totals 已并入 classify_hexiao.expand_payment（ΣH ≤ 到账 → E4/E5），
# 对应用例见 test_classify.py::test_writeoff_over_arrival_is_e4 / test_deliver_over_arrival_is_e5_split_row


# ---------- 案例ID ----------

def test_case_id_is_ar_times_so():
    r = C.classify_one(
        {"ar": "AR1", "so": "SO1", "amount_orig": 1.0, "currency": "人民币CNY",
         "status": "手动核销", "channel": "duizhang", "customer": "X",
         "hexiao_date": dt.date(2026, 7, 8), "shoukuan_date": dt.date(2026, 7, 8)},
        None, {}, 0.0, 2026,
    )
    assert r["case_id"] == "AR1|SO1"


def test_ar_summary_marks_partial():
    # 有的能填(auto)、有的**没交付进表**(E2) → 部分
    results = [
        {"ar": "AR1", "so": "SO1", "bucket": "auto"},
        {"ar": "AR1", "so": "SO2", "bucket": "hold", "code": "E2"},
    ]
    s = C.build_ar_summary(results)[0]
    assert s["流转表_是否更新应收款_建议"] == "部分"
    assert s["待处理SO"] == ["SO2"]


def test_ar_summary_e5_partial_writeoff_is_yes():
    """2026-07-24 明妹口径 + 实测 AR26070112：一笔到账里有「部分核销」(E5) 的单，
    钱已全核落地、她拆行即算更新 → 流转表应是「是」，不能因 E5 掉成「部分」。
    这是 E5 被当成"没更新"的老 bug 的防回归。"""
    results = [
        {"ar": "AR1", "so": "SO1", "bucket": "auto"},
        {"ar": "AR1", "so": "SO2", "bucket": "hold", "code": "E5"},
    ]
    s = C.build_ar_summary(results)[0]
    assert s["流转表_是否更新应收款_建议"] == "是"


def test_ar_summary_all_not_delivered_is_blank():
    """整笔到账全是没交付进表(E2/E3) → 一个都更不了 → 空白。"""
    results = [
        {"ar": "AR1", "so": "SO1", "bucket": "hold", "code": "E3"},
        {"ar": "AR1", "so": "SO2", "bucket": "hold", "code": "E2"},
    ]
    s = C.build_ar_summary(results)[0]
    assert s["流转表_是否更新应收款_建议"] == ""


# ---------- 台账：多 SO 不再被吞 ----------

def test_ledger_keeps_each_so_separately():
    result = {
        "auto": [],
        "hold": [
            {"ar": "AR1", "so": "SO1", "case_id": "AR1|SO1", "code": "E2", "reason": "未交付"},
            {"ar": "AR1", "so": "SO2", "case_id": "AR1|SO2", "code": "E1", "reason": "分笔"},
        ],
        "exception": [],
    }
    out = R.merge_from_classify([], result, "2026-07-22")
    assert len(out) == 2
    assert {r["SO"] for r in out} == {"SO1", "SO2"}


def test_auto_one_so_does_not_release_whole_ar():
    """一个 SO 转 auto 不得把同 AR 的另一个 SO 也标成可补做。"""
    rows = [
        {"案例ID": "AR1|SO1", "AR": "AR1", "SO": "SO1", "状态": "挂起", "E码": "E2",
         "原因": "", "复查条件": "", "挂起日": "2026-07-01", "重扫次数": 0, "最近重扫日": ""},
        {"案例ID": "AR1|SO2", "AR": "AR1", "SO": "SO2", "状态": "挂起", "E码": "E1",
         "原因": "", "复查条件": "", "挂起日": "2026-07-01", "重扫次数": 0, "最近重扫日": ""},
    ]
    result = {"auto": [{"ar": "AR1", "so": "SO1", "case_id": "AR1|SO1"}], "hold": [], "exception": []}
    out = {r["案例ID"]: r for r in R.merge_from_classify(rows, result, "2026-07-22")}
    assert out["AR1|SO1"]["状态"] == "可补做"
    assert out["AR1|SO2"]["状态"] == "挂起"


def test_revisit_condition_is_actionable():
    assert "月初" in R.revisit_condition("E2")
    assert R.revisit_condition("E1")


# ---------- 真重判 ----------

class _FakeLedger:
    def __init__(self, so_index):
        self.so_index = so_index


def test_reclassify_promotes_when_so_now_in_ledger():
    """月初贴完交付 → 之前 E2 挂起的笔应自动升级为可补做（第8步的核心价值）。"""
    rows = [{"案例ID": "AR1|SO1", "AR": "AR1", "SO": "SO1", "E码": "E2",
             "状态": "挂起", "原因": "", "复查条件": "", "重扫次数": 1, "最近重扫日": ""}]
    stat = R.reclassify_against_ledger(rows, _FakeLedger({"SO1": [10]}), "2026-08-01")
    assert rows[0]["状态"] == "可补做" and stat["升级可补做"] == 1


def test_reclassify_multi_row_so_becomes_e8_not_auto():
    rows = [{"案例ID": "AR1|SO1", "AR": "AR1", "SO": "SO1", "E码": "E2",
             "状态": "挂起", "原因": "", "复查条件": "", "重扫次数": 1, "最近重扫日": ""}]
    R.reclassify_against_ledger(rows, _FakeLedger({"SO1": [10, 22]}), "2026-08-01")
    assert rows[0]["状态"] == "挂起" and rows[0]["E码"] == "E8"


def test_reclassify_skips_codes_needing_fresh_export():
    """E1（分笔待回满）本地判不动，必须如实保留，不许乐观放行。"""
    rows = [{"案例ID": "AR1|SO1", "AR": "AR1", "SO": "SO1", "E码": "E1",
             "状态": "挂起", "原因": "", "复查条件": "", "重扫次数": 1, "最近重扫日": ""}]
    stat = R.reclassify_against_ledger(rows, _FakeLedger({"SO1": [10]}), "2026-08-01")
    assert rows[0]["状态"] == "挂起" and stat["本地判不动"] == 1


def test_reclassify_without_ledger_does_nothing():
    rows = [{"案例ID": "A|B", "AR": "A", "SO": "B", "E码": "E2", "状态": "挂起",
             "原因": "", "复查条件": "", "重扫次数": 0, "最近重扫日": ""}]
    stat = R.reclassify_against_ledger(rows, None, "2026-08-01")
    assert rows[0]["状态"] == "挂起" and stat["跳过"] == 1


# ---------- 源文件只读保证 ----------

def test_verify_sources_detects_modification(tmp_path):
    ws = tmp_path / "工作区"
    (ws / "02_我的表副本").mkdir(parents=True)
    f = ws / "02_我的表副本" / "a.xlsx"
    f.write_bytes(b"hello")
    assert V.do_snapshot(ws) == 0
    assert V.do_verify(ws) == 0
    f.write_bytes(b"hello!")  # 模拟被改
    assert V.do_verify(ws) == 1


def test_verify_sources_detects_missing(tmp_path):
    ws = tmp_path / "工作区"
    (ws / "01_智云导出").mkdir(parents=True)
    f = ws / "01_智云导出" / "b.xlsx"
    f.write_bytes(b"x")
    V.do_snapshot(ws)
    f.unlink()
    assert V.do_verify(ws) == 1


# ---------- 真实表集成 ----------

@pytest.mark.skipif(not FLOW_HK.is_file() or not FLOW_WX.is_file(), reason="无真实流转表副本")
def test_real_flow_tables_load_and_match():
    flow = FL.FlowLedger.from_paths([FLOW_HK, FLOW_WX])
    assert len(flow.rows) > 500, "两张真实流转表应有数百行"
    # 用表里真实存在的一行反查，必须能命中自己
    probe = next(r for r in flow.rows if r["amount"] and r["payer"] and r["date"])
    hit = flow.match(probe["date"], probe["amount"], probe["payer"])
    assert hit["hits"] >= 1 and hit["matched_by"].startswith("三键")


@pytest.mark.skipif(not FLOW_HK.is_file(), reason="无真实流转表副本")
def test_annotate_records_fills_flow_fields():
    flow = FL.FlowLedger.from_paths([FLOW_HK])
    probe = next(r for r in flow.rows if r["amount"] and r["payer"] and r["date"])
    recs = [{"ar": "ARX", "so": "SOX", "shoukuan_date": probe["date"],
             "amount_orig": probe["amount"], "customer": probe["payer"], "fee": 0.0}]
    FL.annotate_records(recs, flow)
    assert recs[0]["flow_hits"] >= 1
    assert "第" in recs[0]["flow_locate"] and "行" in recs[0]["flow_locate"]


# ---------- 收款方式：冲预收双义（回放校准抓到的规则缺口）----------

def test_pay_way_prepaid_type_is_hui_now():
    assert C.common.pay_way("预存已核销") == "汇"  # 口径修正：类型不再判冲预收


def test_pay_way_same_month_is_hui():
    d = dt.date(2026, 7, 8)
    assert C.common.pay_way("手动核销", d, d) == "汇"


def test_pay_way_cross_month_is_chongyushou():
    """晚核销标签：6月到账、7月核销 → 冲预收（实测明妹就是这么填的）。"""
    assert C.common.pay_way("手动核销", dt.date(2026, 6, 26), dt.date(2026, 7, 8)) == "冲预收"


def test_pay_way_missing_dates_falls_back():
    assert C.common.pay_way("手动核销") == "汇"
