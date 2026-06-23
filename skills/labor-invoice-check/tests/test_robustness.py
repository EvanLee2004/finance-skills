#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""劳务发票核对 · 回归测试
合成用例（金标=亮晶口头确认的口径）+ 真实数据冒烟（若 测试数据/ 在）。
跑：python3 tests/test_robustness.py
"""
import os
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))
import check  # noqa


def _inv(pairs):
    """pairs: list[(idno, amount)] → 发票聚合结构（按身份证求和）。"""
    sum_by_id, cnt_by_id = defaultdict(float), defaultdict(int)
    for idno, amt in pairs:
        sum_by_id[check.norm_id(idno)] += amt
        cnt_by_id[check.norm_id(idno)] += 1
    return dict(sum_by_id=sum_by_id, cnt_by_id=cnt_by_id, sum_by_name=defaultdict(float))


PASS = 0
FAIL = 0


def check_eq(label, got, want):
    global PASS, FAIL
    if got == want:
        PASS += 1
    else:
        FAIL += 1
        print(f"  ✗ {label}: got={got!r} want={want!r}")


def status_of(clist, inv, name):
    for r in check.classify(clist, inv):
        if r["供应商姓名"] == name:
            return r["状态"]
    return None


def test_core_rules():
    print("· 合成金标用例")
    # 默认参数（不读 md 的话）
    check.CONFIG.update({"THRESHOLD": 800.0, "TOLERANCE": 1.0, "INTERN_KEYWORDS": ["Intern", "实习"]})
    clist = [
        dict(name="杨仲舒", pay=3682.36, note="Freelancer", idno="32080219941127202X"),  # 多票求和=应付→可付
        dict(name="费祥汝", pay=1148, note="Freelancer", idno="370612199805162721"),       # 无票→未开票标黄
        dict(name="张依炜", pay=200, note="Intern，个税起征点特殊", idno="110106200105082723"),  # 实习生→豁免
        dict(name="王嘉乐", pay=1600, note="Intern，个税起征点特殊", idno="370829200400000000"),  # 实习生即使>800也豁免
        dict(name="小额哥", pay=800, note="Freelancer", idno="110000000000000001"),         # =800→可付(不要求开票)
        dict(name="缺票哥", pay=1000, note="Freelancer", idno="110000000000000002"),        # 开600<1000→缺票标黄
        dict(name="多开哥", pay=1000, note="Freelancer", idno="110000000000000003"),        # 开1200>1000→可付(多开)
        dict(name="尾差哥", pay=1000, note="Freelancer", idno="110000000000000004"),        # 开999.5差0.5≤容差→可付
        dict(name="JOHN SMITH", pay=5000, note="Freelancer", idno=""),                      # 纯英文→外国人豁免
        dict(name="零票哥", pay=1000, note="Freelancer", idno="110000000000000005"),        # 身份证匹到但合计0→未开票
    ]
    inv = _inv([
        ("32080219941127202X", 2091.52), ("32080219941127202X", 1590.84),  # 杨仲舒两票
        ("110000000000000002", 600.0),
        ("110000000000000003", 1200.0),
        ("110000000000000004", 999.5),
        ("110000000000000005", 0.0),
    ])
    check_eq("杨仲舒 多票求和=应付→可付", status_of(clist, inv, "杨仲舒"), "可付")
    check_eq("费祥汝 无票→未开票", status_of(clist, inv, "费祥汝"), "标黄-未开票")
    check_eq("张依炜 实习生→豁免", status_of(clist, inv, "张依炜"), "豁免-实习生")
    check_eq("王嘉乐 实习生>800仍豁免", status_of(clist, inv, "王嘉乐"), "豁免-实习生")
    check_eq("小额哥 =800→可付", status_of(clist, inv, "小额哥"), "可付")
    check_eq("缺票哥 开600<1000→缺票", status_of(clist, inv, "缺票哥"), "标黄-缺票")
    check_eq("多开哥 开1200>1000→可付", status_of(clist, inv, "多开哥"), "可付")
    check_eq("尾差哥 差0.5≤容差→可付", status_of(clist, inv, "尾差哥"), "可付")
    check_eq("JOHN SMITH 纯英文→外国人豁免", status_of(clist, inv, "JOHN SMITH"), "豁免-外国人")
    check_eq("零票哥 匹到但合计0→未开票", status_of(clist, inv, "零票哥"), "标黄-未开票")


def test_id_match_beats_name():
    print("· 身份证号匹配（重名/姓名特殊字符靠它区分）")
    check.CONFIG.update({"THRESHOLD": 800.0, "TOLERANCE": 1.0, "INTERN_KEYWORDS": ["Intern"]})
    # 两个同名"张伟"，靠身份证区分：一个开够、一个没开
    clist = [
        dict(name="张伟", pay=1000, note="Freelancer", idno="110000000000000010"),
        dict(name="张伟", pay=1000, note="Freelancer", idno="110000000000000011"),
        dict(name="流畅（阿拉伯语）", pay=1000, note="Freelancer", idno="110000000000000012"),  # 姓名特殊字符
    ]
    inv = _inv([("110000000000000010", 1000.0), ("110000000000000012", 1000.0)])
    res = check.classify(clist, inv)
    s10 = [r["状态"] for r in res if r["供应商姓名"] == "张伟"]
    check_eq("同名张伟 一个可付一个未开票", sorted(s10), sorted(["可付", "标黄-未开票"]))
    check_eq("姓名特殊字符 靠身份证匹到→可付", res[2]["状态"], "可付")


def test_to_number():
    print("· 金额解析鲁棒性")
    import datetime
    check_eq("逗号数字", check.to_number("1,234.5"), 1234.5)
    check_eq("空", check.to_number(""), None)
    check_eq("横杠", check.to_number("-"), None)
    check_eq("datetime坏值→None", check.to_number(datetime.datetime(1900, 1, 1)), None)
    check_eq("norm_id去空格大写", check.norm_id("3208 0219941127202x"), "32080219941127202X")


def test_sheet_isolation():
    """抗干扰：发票文件里台账 sheet 不叫Sheet1、且塞了多个杂 sheet（含迷惑性的'财务核对'），
       程序应只认含『纳税人识别号+合计金额』的真台账，无视其余。清单侧同理。"""
    print("· 多 sheet 抗干扰（按列认台账，无视其它 sheet）")
    import tempfile
    import openpyxl as opx

    # --- 造发票文件：台账放在非Sheet1的、随便命名的 sheet；另加杂 sheet ---
    inv_path = os.path.join(tempfile.gettempdir(), "_t_inv.xlsx")
    wb = opx.Workbook()
    wb.active.title = "封面说明"                       # 杂 sheet 1
    wb.active["A1"] = "本月发票统计 仅供内部"
    sc = wb.create_sheet("财务核对")                    # 迷惑 sheet：像清单不像台账
    sc.append(["序号", "供应商姓名", "应付金额", "备注", "开票金额", "check"])
    sc.append([1, "张三", 1000, "Freelancer", 1000, 0])
    led = wb.create_sheet("发票明细202607")            # 真台账：不叫 Sheet1
    led.append(["序号", "销售方信息名称", "销售方信息纳税人识别号", "金额（元）", "合计金额（元）"])
    led.append([1, "张三", "110000000000000020", 970, 1000])
    led.append([2, "张三", "110000000000000020", 485, 500])   # 同一身份证两张票
    wb.create_sheet("Sheet3").append(["姓名", "有票", "无票"])  # 杂 sheet 2
    wb.save(inv_path); wb.close()

    # --- 造清单文件：国内个人 sheet 改了名 + 标题行 + 一个杂 sheet ---
    lst_path = os.path.join(tempfile.gettempdir(), "_t_list.xlsx")
    wb2 = opx.Workbook()
    wb2.active.title = "汇总透视"                        # 杂 sheet
    wb2.active["A1"] = "别读我"
    pay = wb2.create_sheet("个人译费2026")              # 真清单：非默认名、带标题行
    pay.append(["2026年7月 国内个人译费", None, None, None])
    pay.append(["序号", "供应商姓名", "应付金额", "备注", "身份证号/护照号"])
    pay.append([1, "张三", 1500, "Freelancer", "110000000000000020"])
    wb2.save(lst_path); wb2.close()

    check.CONFIG.update({"THRESHOLD": 800.0, "TOLERANCE": 1.0, "INTERN_KEYWORDS": ["Intern"]})
    la, ia = check.load_aliases()
    clist, lhdr, lw = check.read_list(lst_path, la)
    inv, ihdr, iw = check.read_invoices(inv_path, ia)
    check_eq("清单认出真sheet(跳过杂sheet/标题行)", [c["name"] for c in clist], ["张三"])
    check_eq("台账认出真sheet(非Sheet1)", "销售方信息纳税人识别号" in ihdr, True)
    check_eq("台账没误读'财务核对'(那表无纳税人识别号)", inv["sum_by_id"].get("110000000000000020"), 1500.0)
    check_eq("两张票求和=1500", round(inv["sum_by_id"]["110000000000000020"], 2), 1500.0)
    res = check.classify(clist, inv)
    check_eq("张三 应付1500 开票1500 → 可付", res[0]["状态"], "可付")
    for p in (inv_path, lst_path):
        try:
            os.remove(p)
        except OSError:
            pass


def test_real_smoke():
    """真实数据冒烟：能跑通、关键案例对、标黄数稳定。测试数据缺则跳过。"""
    # HERE=.../finance-skills/skills/labor-invoice-check/tests → 上溯4层到 财务部skills
    财务部skills = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(HERE))))
    D = os.path.join(财务部skills, "技能", "劳务发票核对", "测试数据")
    lp = os.path.join(D, "待支付译费.xlsx")
    ip = os.path.join(D, "个人发票统计.xlsx")
    if not (os.path.isfile(lp) and os.path.isfile(ip)):
        print(f"· 真实冒烟：跳过（测试数据不在 {D}）")
        return
    print("· 真实数据冒烟")
    check.load_rules()
    la, ia = check.load_aliases()
    clist, _, lw = check.read_list(lp, la)
    inv, _, iw = check.read_invoices(ip, ia)
    check_eq("清单读到464人(已剔除合计行)", len(clist), 464)
    res = check.classify(clist, inv)
    by_name = {r["供应商姓名"]: r for r in res}
    check_eq("杨仲舒(真实)→可付", by_name.get("杨仲舒", {}).get("状态"), "可付")
    check_eq("费祥汝(真实)→未开票", by_name.get("费祥汝", {}).get("状态"), "标黄-未开票")
    from collections import Counter
    cnt = Counter(r["状态"] for r in res)
    flagged = sum(v for k, v in cnt.items() if k.startswith("标黄"))
    print(f"    实测分布: {dict(cnt)} | 标黄={flagged}")
    check_eq("标黄合计=18(剔除合计行后)", flagged, 18)


if __name__ == "__main__":
    test_to_number()
    test_core_rules()
    test_id_match_beats_name()
    test_sheet_isolation()
    test_real_smoke()
    print(f"\n{'='*40}\nPASS={PASS}  FAIL={FAIL}")
    sys.exit(1 if FAIL else 0)
