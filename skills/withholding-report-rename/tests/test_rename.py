#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
回归测试：5 个样本 PDF -> 5 个已知正确新名（含金额）。真 pytest。

样本 PDF 含境外客户名，不进 git（见 .gitignore），放在本地资料家：
    财务部skills/技能/代扣代缴申报表重命名/测试数据/
也可用环境变量 WR_TESTDATA 指向别处。测试数据不在时 skip，
不让缺数据把别的机器上的回归搞红。
"""
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SKILL_DIR, "scripts"))

import rename  # noqa: E402

# 默认测试数据位置：finance-skills/skills/<id> -> 上溯到 甲骨易实习 -> 技能家
DEFAULT_TESTDATA = os.path.normpath(os.path.join(
    SKILL_DIR, "..", "..", "..", "技能", "代扣代缴申报表重命名", "测试数据"))
TESTDATA = os.environ.get("WR_TESTDATA", DEFAULT_TESTDATA)

# 原文件名关键片段 -> 期望新名（不含 .pdf）
GOLDEN = {
    "bridge": "BRIDGE TECHNOLOGY LIMITED134133.8",
    "(79)": "D'ArteMediaLLC21496.8",
    "153136.105": "LUCALIZE MANAGEMENT CONSULTANCIES CO.L.L.C1500.4",
    "102206.830": "TerraTranslationsLLC7170.69",
    "104132.977": "WordPowerS.r.l3718.65",
}


def _pdfs_available():
    if not os.path.isdir(TESTDATA):
        return False
    return any(f.lower().endswith(".pdf") for f in os.listdir(TESTDATA))


@pytest.mark.skipif(not _pdfs_available(), reason=f"测试 PDF 不在 {TESTDATA}（设 WR_TESTDATA）")
def test_golden_rename_five_samples():
    """原 main() 的 5 条金标断言，语义不删不放宽。"""
    pdfs = [os.path.join(TESTDATA, f) for f in os.listdir(TESTDATA)
            if f.lower().endswith(".pdf")]
    assert pdfs, f"{TESTDATA} 里没有 PDF"

    pattern = rename.build_suffix_pattern(rename.load_suffix_words())
    overrides = rename.load_overrides()
    plans = [rename.plan_one(p, pattern, overrides) for p in pdfs]
    rename.dedup(plans)
    got = {os.path.basename(p["src"]): p for p in plans}

    fails = []
    for frag, expect in GOLDEN.items():
        match = [b for b in got if frag in b]
        if not match:
            fails.append(f"找不到含 {frag!r} 的样本 PDF")
            continue
        rec = got[match[0]]
        if rec["status"] != "ok":
            fails.append(f"{match[0]}: 进了待人工（{rec['note']}），期望可重命名")
        elif rec["newbase"] != expect:
            fails.append(f"{match[0]}:\n    期望 {expect!r}\n    实际 {rec['newbase']!r}")
        else:
            print(f"  ✓ {frag} -> {rec['newbase']}")

    assert not fails, "回归失败：\n" + "\n".join(f"  - {x}" for x in fails)
    assert len(GOLDEN) == 5


def test_load_suffix_words_nonempty():
    """无 PDF 也能跑的轻量断言：词表可加载。"""
    words = rename.load_suffix_words()
    assert isinstance(words, list) and len(words) >= 5
    assert any(w == "LIMITED" or w == "TECHNOLOGY" for w in words)
