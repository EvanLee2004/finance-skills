#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
回归测试：5 个样本 PDF -> 5 个已知正确新名（含金额）。

样本 PDF 含境外客户名，不进 git（见 .gitignore），放在本地资料家：
    财务部skills/技能/代扣代缴申报表重命名/测试数据/
也可用环境变量 WR_TESTDATA 指向别处。测试数据不在时优雅跳过（exit 0），
不让缺数据把别的机器上的回归搞红。

跑：python3 tests/test_rename.py
"""
import os
import sys

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


def main():
    if not os.path.isdir(TESTDATA):
        print(f"[skip] 测试数据不在：{TESTDATA}（设 WR_TESTDATA 指向 5 个样本 PDF）")
        return 0
    pdfs = [os.path.join(TESTDATA, f) for f in os.listdir(TESTDATA)
            if f.lower().endswith(".pdf")]
    if not pdfs:
        print(f"[skip] {TESTDATA} 里没有 PDF")
        return 0

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

    if fails:
        print("\n✗ 回归失败：")
        for x in fails:
            print("  -", x)
        return 1
    print(f"\n✓ 全部 {len(GOLDEN)} 个样本通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
