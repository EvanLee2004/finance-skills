# -*- coding: utf-8 -*-
"""
应收账款合并 · merge.py  （财务部 Agent 技能 · 样板）
================================================================
按 SOP（references/SOP_应收账款合并.md）实现 S1–S8：
  S1 合并各分年 sheet → 一张表（年度标签 + 新智云单号补位 + 按列名对齐）
  S2 账龄(月份) = (年差×12 + 月差) − 1
  S3 按新智云单号从回填源 VLOOKUP 回填标注列
  S4 销售归属（维护表驱动）：离职→接手「接手-离职」复合名；坏账桶光名；客户重分配优先
  S5 删行：应收金额=0（已回款/核销）
  S6 复核：年度自检 / #N/A 清单 / 名称残留
  S7 排序：按年度降序
  S8 透视汇总 Sheet2：应收金额 按 销售人员→客户

S1–S3 移植自验收过的 merge_receivables.py（账龄/交付月份与赵成品逐行 100% 一致）。

用法（agent 识别好文件后传路径；也可自动扫工作区 input/）：
  python3 merge.py --source 源台账.xlsx --ref 回填源.xlsx --rules 营销人员应收匹配规则.xlsx --out 应收all.xlsx
  python3 merge.py --inspect            # 只打印识别到的三个输入与表头
其中 --ref / --rules 可省（省 ref 跳过回填；省 rules 跳过归属）。
"""
import os
import re
import sys
import json
import argparse
import datetime
from collections import defaultdict
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
CONFIG_DIR = os.path.join(SKILL_DIR, "config")
WORK_INPUT = os.path.join(SKILL_DIR, "工作区", "input")
WORK_OUTPUT = os.path.join(SKILL_DIR, "工作区", "output")

# ========================= 配置 =========================
CONFIG = {
    "REPORT_MONTH": None,          # 账龄基准月 YYYYMM；None=从源文件名解析、再退电脑当月
    "AGING_MINUS_ONE": True,
    "KEY": "新智云单号",
    "SHEET_YEAR_LABEL": {"6月批量": "技术统一确认"},
    "EXCLUDE_SHEETS": ["Sheet1", "Sheet2", "Sheet3"],
    "REF_SHEET": None,
}

# 列名别名（会变的配置）：优先读 config/列名别名.json，缺则用内置默认。
_DEFAULT_ALIASES = {
    "销售人员": ["销售人员", "销售"],
    "客户名称": ["客户名称", "客户"],
    "单号": ["单号", "老单号", "订单号", "订单编号"],
    "新智云单号": ["新智云单号"],
    "翻译类型": ["翻译类型", "订单类型"],
    "文件名": ["文件名", "名称"],
    "应收金额": ["应收金额", "订单折合本币", "金额"],
    "_交付月份直取": ["交付月份", "项目交付", "销售确认"],
    "_交付日期回退": ["完成时间", "项目交付日期", "交件日期", "客户交付日期", "截止时间", "下单时间"],
}
_DEFAULT_ANNOTATION = {
    "结算阶段": ["结算阶段"],
    "预计回款日期": ["预计回款"],
    "销售解释说明": ["销售解释"],
    "有无合同": ["有无合同"],
    "合同分类": ["合同分类"],
    "框架合同PO记录": ["PO单", "下单记录", "PO"],
    "客户确认": ["客户正式确认", "客户确认"],
    "客户结算周期": ["客户结算周期", "结算周期"],
    "是否按月给客户发结算单": ["按月给客户发结算单", "按月", "发结算单"],
}


def load_aliases():
    p = os.path.join(CONFIG_DIR, "列名别名.json")
    if os.path.isfile(p):
        try:
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
            return d.get("COLUMN_ALIASES", _DEFAULT_ALIASES), d.get("ANNOTATION_COLS", _DEFAULT_ANNOTATION)
        except Exception as e:
            log(f"⚠ 读 列名别名.json 失败({e})，用内置默认。")
    return _DEFAULT_ALIASES, _DEFAULT_ANNOTATION


COLUMN_ALIASES, ANNOTATION_COLS = None, None  # 在 main 里加载

OUTPUT_HEADERS = [
    "年度", "销售人员", "客户名称", "新智云单号", "文件名", "应收金额", "交付月份", "账龄(月份)",
    "结算阶段", "预计回款日期", "销售解释说明", "有无合同", "合同分类",
    "框架合同PO记录", "客户确认", "客户结算周期", "是否按月给客户发结算单",
]


# --------------------------- 工具函数（移植自验收脚本） ---------------------------
def log(msg):
    print(msg, flush=True)


def to_yyyymm(v):
    def mk(y, m):
        return f"{int(y):04d}{int(m):02d}" if 1 <= int(m) <= 12 else None
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, (datetime.datetime, datetime.date, pd.Timestamp)):
        return mk(v.year, v.month)
    if isinstance(v, int) or (isinstance(v, float) and float(v).is_integer()):
        s = str(int(v))
        if len(s) in (6, 8) and s.isdigit():
            return mk(s[:4], s[4:6])
        return None
    s = str(v).strip()
    if not s or set(s) <= set("-/ 　"):
        return None
    m = re.match(r"^(\d{4})[\-/.](\d{1,2})", s)
    if m:
        return mk(m.group(1), m.group(2))
    if re.fullmatch(r"\d{6}", s) or re.fullmatch(r"\d{8}", s):
        return mk(s[:4], s[4:6])
    d = pd.to_datetime(s, errors="coerce")
    if pd.notna(d):
        return mk(d.year, d.month)
    return None


def aging_months(deliver_yyyymm, base_yyyymm, minus_one=True):
    d, b = deliver_yyyymm, base_yyyymm
    if not (isinstance(d, str) and isinstance(b, str)
            and len(d) >= 6 and len(b) >= 6 and d[:6].isdigit() and b[:6].isdigit()):
        return None
    by, bm = int(b[:4]), int(b[4:6])
    dy, dm = int(d[:4]), int(d[4:6])
    a = (by - dy) * 12 + (bm - dm)
    if minus_one:
        a -= 1
    return a


def clean_key(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    s = str(x).strip()
    return s or None


def year_label(sheet_name):
    s = str(sheet_name).strip()
    if re.fullmatch(r"\d{4}", s):
        return int(s)
    return CONFIG["SHEET_YEAR_LABEL"].get(s, s)


def year_sort_key(y):
    s = str(y)
    if re.fullmatch(r"\d{4}", s):
        return (0, -int(s))
    return (1, s)


def parse_report_month(source_path):
    b = os.path.basename(source_path)
    m = re.search(r"(\d{4})[.\-_年](\d{1,2})", b)
    if m:
        return f"{int(m.group(1)):04d}{int(m.group(2)):02d}"
    today = datetime.date.today()
    return f"{today.year:04d}{today.month:02d}"


# --------------------------- 输入识别（按内容） ---------------------------
EXCEL_EXTS = (".xlsx", ".xlsm")


def _scan_features(path):
    """提取判别特征：标注列种类数 / 是否含新智云单号 / 是否含年度列 / 年份分表数 / 是否像维护表。"""
    ann = 0
    has_key = has_year_col = is_rules = False
    year_sheets = 0
    try:
        with pd.ExcelFile(path) as xls:
            for s in xls.sheet_names:
                if re.fullmatch(r"\d{4}", str(s).strip()):
                    year_sheets += 1
                cols = [str(c).strip() for c in pd.read_excel(xls, sheet_name=s, nrows=0).columns]
                if any("新智云" in c for c in cols):
                    has_key = True
                if "年度" in cols:
                    has_year_col = True
                if any("离职" in c for c in cols) and any("接手" in c for c in cols):
                    is_rules = True
                hits = sum(1 for kws in ANNOTATION_COLS.values()
                           if any(any(kw in c for c in cols) for kw in kws))
                ann = max(ann, hits)
    except Exception:
        pass
    return {"ann": ann, "key": has_key, "ycol": has_year_col, "years": year_sheets, "rules": is_rules}


def find_inputs(input_dir):
    """从一个目录按内容认出【源台账 / 回填源 / 维护表】，不靠文件名。"""
    if not os.path.isdir(input_dir):
        return None, None, None
    cands = [os.path.join(input_dir, f) for f in sorted(os.listdir(input_dir))
             if f.lower().endswith(EXCEL_EXTS) and not f.startswith(("~$", "."))]
    if not cands:
        return None, None, None
    feats = {p: _scan_features(p) for p in cands}
    # 维护表：含 离职+接手 列
    rules = next((p for p in cands if feats[p]["rules"]), None)
    rest = [p for p in cands if p != rules]
    # 回填源：含标注列(>=3)或含年度列
    ref_cands = [p for p in rest if feats[p]["ann"] >= 3 or feats[p]["ycol"]]
    ref = max(ref_cands, key=lambda p: (feats[p]["ann"], feats[p]["ycol"])) if ref_cands else None
    # 源台账：剩下里年份分表最多的
    src_cands = [p for p in rest if p != ref] or rest
    source = max(src_cands, key=lambda p: feats[p]["years"]) if src_cands else None
    if ref is not None and source is not None and os.path.abspath(ref) == os.path.abspath(source):
        ref = None
    return source, ref, rules


# --------------------------- S1 合并 ---------------------------
def normalize_sheet(df, sheet_name):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    out = pd.DataFrame(index=df.index)
    for canon, names in COLUMN_ALIASES.items():
        hit = next((n for n in names if n in df.columns), None)
        out[canon] = df[hit] if hit else None
    ym1 = out["_交付月份直取"].map(to_yyyymm)
    ym2 = out["_交付日期回退"].map(to_yyyymm)
    out["交付月份"] = ym1.where(ym1.notna(), ym2)
    out["交付月份"] = out["交付月份"].map(lambda x: x if isinstance(x, str) else None)
    out.drop(columns=["_交付月份直取", "_交付日期回退"], inplace=True)
    out["年度"] = year_label(sheet_name)
    out["应收金额"] = pd.to_numeric(out["应收金额"], errors="coerce")
    out["新智云单号"] = out["新智云单号"].map(clean_key)
    out["单号"] = out["单号"].map(clean_key)
    # 新智云单号补位：缺则用 单号/老单号 补（S1 关键步）
    fb = out["新智云单号"].isna() & out["单号"].notna()
    out["新智云单号"] = out["新智云单号"].where(~fb, out["单号"])
    out["__主键回退"] = fb
    out["__来源sheet"] = str(sheet_name)
    key_cols = ["客户名称", "单号", "新智云单号", "应收金额"]
    out = out[~out[key_cols].isna().all(axis=1)].reset_index(drop=True)
    return out


def load_and_merge(source_path):
    frames, report = [], []
    xls = pd.ExcelFile(source_path)
    try:
        for s in list(xls.sheet_names):
            if s in CONFIG["EXCLUDE_SHEETS"]:
                report.append((s, "跳过", 0)); continue
            df = pd.read_excel(xls, sheet_name=s)
            if df.empty:
                report.append((s, "空表", 0)); continue
            norm = normalize_sheet(df, s)
            frames.append(norm)
            n_fb = int(norm["__主键回退"].sum())
            report.append((s, f"年度={year_label(s)}" + (f"; 主键回退{n_fb}行" if n_fb else ""), len(norm)))
    finally:
        xls.close()
    if not frames:
        raise RuntimeError("源台账没有可合并的 sheet。")
    return pd.concat(frames, ignore_index=True), report


# --------------------------- S2 账龄 ---------------------------
def add_aging(master, base_yyyymm):
    master = master.copy()
    master["账龄(月份)"] = master["交付月份"].map(
        lambda ym: aging_months(ym, base_yyyymm, CONFIG["AGING_MINUS_ONE"]))
    return master


# --------------------------- S3 回填 ---------------------------
def detect_ref_sheet(ref_path):
    xls = pd.ExcelFile(ref_path)
    try:
        best, best_score = xls.sheet_names[0], -1
        for s in xls.sheet_names:
            cols = [str(c).strip() for c in pd.read_excel(xls, sheet_name=s, nrows=0).columns]
            has_key = any(c == CONFIG["KEY"] for c in cols) or any("新智云" in c for c in cols)
            ann = sum(1 for kws in ANNOTATION_COLS.values()
                      if any(any(kw in c for c in cols) for kw in kws))
            score = ann + (10 if has_key else 0)
            if score > best_score:
                best, best_score = s, score
    finally:
        xls.close()
    return best


def blankify_annotation(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        try:
            return None if float(v) == 0 else v
        except (TypeError, ValueError):
            return v
    s = str(v).strip()
    return None if s in ("", "0", "0.0") else v


def load_ref_table(ref_path):
    ref_sheet = CONFIG["REF_SHEET"] or detect_ref_sheet(ref_path)
    d = pd.read_excel(ref_path, sheet_name=ref_sheet)
    d.columns = [str(c).strip() for c in d.columns]
    keycol = CONFIG["KEY"] if CONFIG["KEY"] in d.columns else next((c for c in d.columns if "新智云" in c), None)
    if keycol is None:
        raise RuntimeError(f"回填源 sheet「{ref_sheet}」找不到主键列。")
    ann_map = {}
    for canon, kws in ANNOTATION_COLS.items():
        found = next((c for c in d.columns if c == canon), None)
        if not found:
            for kw in kws:
                found = next((c for c in d.columns if kw in c), None)
                if found:
                    break
        ann_map[canon] = found
    ref = pd.DataFrame()
    ref[CONFIG["KEY"]] = d[keycol].map(clean_key)
    for canon, src in ann_map.items():
        ref[canon] = (d[src].map(blankify_annotation) if src else None)
    ref = ref.dropna(subset=[CONFIG["KEY"]]).drop_duplicates(subset=[CONFIG["KEY"]], keep="first")
    return ref, ref_sheet, ann_map


def backfill(master, ref_path):
    ref, ref_sheet, ann_map = load_ref_table(ref_path)
    master = master.copy()
    master[CONFIG["KEY"]] = master[CONFIG["KEY"]].map(clean_key)
    merged = master.merge(ref, on=CONFIG["KEY"], how="left")
    refkeys = set(ref[CONFIG["KEY"]].dropna())
    merged["__matched"] = merged[CONFIG["KEY"]].isin(refkeys) & merged[CONFIG["KEY"]].notna()
    meta = {"ref_sheet": ref_sheet, "ref_rows": len(ref), "ann_map": ann_map,
            "matched": int(merged["__matched"].sum())}
    return merged, meta


# --------------------------- S4 销售归属（维护表驱动） ---------------------------
def _base_name(s):
    """去结尾数字：高美杰1 → 高美杰。"""
    return re.sub(r"\d+$", "", str(s).strip()).strip()


def load_rules(rules_path):
    """读维护表 → (离职→接手, 客户→接手)。表头：离职销售|接手销售|(空)|客户需重新分配|接手销售。"""
    d = pd.read_excel(rules_path, header=None)
    departed_map, customer_map = {}, {}
    for _, r in d.iterrows():
        vals = [("" if pd.isna(x) else str(x).strip()) for x in r.tolist()]
        vals += [""] * (5 - len(vals))
        dep, dep_to, _, cust, cust_to = vals[0], vals[1], vals[2], vals[3], vals[4]
        if dep and dep_to and dep not in ("离职销售",):
            departed_map[dep] = dep_to
        if cust and cust_to and cust not in ("客户需重新分配",):
            customer_map[cust] = cust_to
    return departed_map, customer_map


# 只有这些离职人保留「接手-离职」复合痕迹（如高美杰=遗留坏账，要追溯）；
# 其余离职人直接换成接手人光名（与赵成品一致：张健-张林烨等不出现，折进张健）。
COMPOSITE_DEPARTED = {"高美杰"}


def attribute_sales(merged, departed_map, customer_map):
    """对每行定最终销售人员。规则（成品实证）：
       ① 客户重分配【全局】优先——客户在清单里就归指定人，不管原销售是谁（含在职）；
       ② 否则原销售是离职人 → 归默认接手；
       ③ 命名：仅 COMPOSITE_DEPARTED（高美杰）保留「接手-离职」复合名（接手是其变体如高美杰1→光名），
          其余一律用接手人光名。
       返回(新merged, 改动清单df)。"""
    merged = merged.copy()
    changes = []
    cust_keys = sorted([k for k in customer_map if k], key=len, reverse=True)  # 长键优先
    new_vals = []
    for orig_raw, cust_raw in zip(merged["销售人员"], merged["客户名称"]):
        orig = "" if orig_raw is None or (isinstance(orig_raw, float) and pd.isna(orig_raw)) else str(orig_raw).strip()
        cust = "" if cust_raw is None or (isinstance(cust_raw, float) and pd.isna(cust_raw)) else str(cust_raw).strip()
        # ① 客户重分配（全局，contains）
        owner = next((customer_map[k] for k in cust_keys if k and k in cust), None)
        # ② 离职默认接手
        if owner is None and orig in departed_map:
            owner = departed_map[orig]
        if owner is None:                      # 不动
            new_vals.append(orig_raw); continue
        # ③ 命名
        if orig in COMPOSITE_DEPARTED:
            final = owner if _base_name(owner) == _base_name(orig) else f"{owner}-{orig}"
        else:
            final = owner
        new_vals.append(final)
        if final != orig:
            changes.append({"原销售": orig, "客户名称": cust, "最终销售": final})
    merged["销售人员"] = new_vals
    return merged, pd.DataFrame(changes, columns=["原销售", "客户名称", "最终销售"])


# --------------------------- S5 删 0 行 ---------------------------
def drop_zero_receivable(merged):
    """删掉应收金额=0 或 空 的行（已回款/核销）。返回(保留, 被删df)。"""
    amt = pd.to_numeric(merged["应收金额"], errors="coerce")
    keep_mask = amt.notna() & (amt != 0)
    removed = merged.loc[~keep_mask].copy()
    return merged.loc[keep_mask].reset_index(drop=True), removed


# --------------------------- S6 复核 ---------------------------
def detect_name_variants(master):
    counts = master["销售人员"].dropna().map(lambda x: str(x).strip())
    counts = counts[counts != ""].value_counts()
    rows, groups = [], defaultdict(list)
    for name, cnt in counts.items():
        groups[_base_name(name)].append((name, int(cnt)))
    for base, variants in groups.items():
        if len(variants) > 1:
            for n, c in sorted(variants, key=lambda x: -x[1]):
                rows.append({"疑似组": base, "名称": n, "行数": c, "类型": "尾号/同名变体"})
    return pd.DataFrame(rows, columns=["疑似组", "名称", "行数", "类型"]).drop_duplicates().reset_index(drop=True)


def residual_departed(master, departed_map):
    """名称残留检查：成品里不应再有离职销售『光名』。"""
    s = master["销售人员"].dropna().map(lambda x: str(x).strip())
    left = s[s.isin(set(departed_map))].value_counts()
    return pd.DataFrame([{"离职光名残留": k, "行数": int(v)} for k, v in left.items()],
                        columns=["离职光名残留", "行数"])


def year_consistency(merged):
    rows = []
    for _, r in merged.iterrows():
        y, ym = r.get("年度"), r.get("交付月份")
        if isinstance(y, int) and isinstance(ym, str) and len(ym) == 6 and int(ym[:4]) != y:
            rows.append({"年度": y, "交付月份": ym, "销售人员": r.get("销售人员"),
                         "客户名称": r.get("客户名称"), "新智云单号": r.get("新智云单号")})
    return pd.DataFrame(rows, columns=["年度", "交付月份", "销售人员", "客户名称", "新智云单号"])


# --------------------------- S8 透视 ---------------------------
def build_pivot(master_out):
    """应收金额 按 销售人员→客户 求和（成品 Sheet2 同款）。"""
    d = master_out[["销售人员", "客户名称", "应收金额"]].copy()
    d["应收金额"] = pd.to_numeric(d["应收金额"], errors="coerce").fillna(0)
    piv = (d.groupby(["销售人员", "客户名称"], dropna=False)["应收金额"]
           .sum().reset_index().sort_values(["销售人员", "应收金额"], ascending=[True, False]))
    return piv


# --------------------------- 输出 ---------------------------
_ILLEGAL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _strip_illegal(df):
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda v: _ILLEGAL_RE.sub("", v) if isinstance(v, str) else v)
    return df


def write_workbook(out_path, master_out, pivot, unmatched, variants, residual, ycheck, removed, report_df):
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    sheets = {
        "主表": master_out, "透视汇总": pivot, "未匹配复核": unmatched,
        "名称变体": variants, "离职残留": residual, "年度校验": ycheck,
        "被删0行": removed, "运行报告": report_df,
    }
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        for name, df in sheets.items():
            _strip_illegal(df).to_excel(w, sheet_name=name, index=False)
    return out_path


# --------------------------- 主流程 ---------------------------
def run(source, ref, rules, out_path, base_month=None):
    log(f"· 源台账：{os.path.basename(source)}")
    log(f"· 回填源：{os.path.basename(ref) if ref else '（无，跳过回填）'}")
    log(f"· 维护表：{os.path.basename(rules) if rules else '（无，跳过归属）'}")

    base_yyyymm = base_month or CONFIG["REPORT_MONTH"] or parse_report_month(source)
    log(f"· 账龄基准月：{base_yyyymm}" + ("（月差再 -1）" if CONFIG["AGING_MINUS_ONE"] else ""))

    # S1 + S2
    master, report = load_and_merge(source)
    log(f"· S1 合并：{len(master)} 行")
    master = add_aging(master, base_yyyymm)
    log("· S2 账龄完成")

    # S3
    if ref:
        try:
            merged, meta = backfill(master, ref)
            log(f"· S3 回填：sheet「{meta['ref_sheet']}」命中 {meta['matched']} 行")
        except Exception as e:
            log(f"⚠ 回填失败跳过：{e}")
            merged, meta = master.copy(), {"ref_sheet": "", "ref_rows": 0, "ann_map": {}, "matched": 0}
            for canon in ANNOTATION_COLS:
                merged[canon] = None
            merged["__matched"] = False
    else:
        merged, meta = master.copy(), {"ref_sheet": "", "ref_rows": 0, "ann_map": {}, "matched": 0}
        for canon in ANNOTATION_COLS:
            merged[canon] = None
        merged["__matched"] = False

    # S4 归属
    departed_map, customer_map = ({}, {})
    reassign = pd.DataFrame(columns=["原销售", "客户名称", "最终销售"])
    if rules:
        departed_map, customer_map = load_rules(rules)
        merged, reassign = attribute_sales(merged, departed_map, customer_map)
        log(f"· S4 归属：维护表 离职{len(departed_map)}人/客户{len(customer_map)}条，改动 {len(reassign)} 行")

    # S5 删 0 行
    merged, removed = drop_zero_receivable(merged)
    log(f"· S5 删 0 行：删 {len(removed)} 行 → 余 {len(merged)} 行")

    # S6 复核
    variants = detect_name_variants(merged)
    residual = residual_departed(merged, departed_map)
    ycheck = year_consistency(merged)

    # S7 排序
    merged["__yorder"] = merged["年度"].map(year_sort_key)
    merged = merged.sort_values("__yorder", kind="stable").reset_index(drop=True)
    master_out = pd.DataFrame({h: (merged[h] if h in merged.columns else None) for h in OUTPUT_HEADERS})

    # 未匹配清单
    um = ~merged["__matched"].astype(bool) if "__matched" in merged.columns else merged[CONFIG["KEY"]].isna()
    unmatched = merged.loc[um, ["年度", "销售人员", "客户名称", "新智云单号", "应收金额", "交付月份"]].copy()

    # S8 透视
    pivot = build_pivot(master_out)

    removed_out = removed[["年度", "销售人员", "客户名称", "新智云单号", "应收金额", "交付月份"]].copy() \
        if len(removed) else pd.DataFrame(columns=["年度", "销售人员", "客户名称", "新智云单号", "应收金额", "交付月份"])

    rep = [["源台账", os.path.basename(source)],
           ["回填源", os.path.basename(ref) if ref else "（无）"],
           ["维护表", os.path.basename(rules) if rules else "（无）"],
           ["账龄基准月", base_yyyymm],
           ["S1 合并行数", len(master)],
           ["S4 归属改动行数", len(reassign)],
           ["S5 删0行数", len(removed)],
           ["主表最终行数", len(master_out)],
           ["未匹配(#N/A)", len(unmatched)],
           ["名称变体待复核", len(variants)],
           ["离职光名残留", len(residual)],
           ["年度不一致", len(ycheck)]]
    report_df = pd.DataFrame(rep, columns=["项目", "值"])

    write_workbook(out_path, master_out, pivot, unmatched, variants, residual, ycheck, removed_out, report_df)
    log(f"\n✓ 完成：{out_path}")
    log(f"  主表 {len(master_out)} 行 | 删0 {len(removed)} | 归属改动 {len(reassign)} | #N/A {len(unmatched)} | 离职残留 {len(residual)}")
    return master_out


def inspect_mode(input_dir):
    source, ref, rules = find_inputs(input_dir)
    for tag, path in [("源台账", source), ("回填源", ref), ("维护表", rules)]:
        log(f"\n[{tag}] {os.path.basename(path) if path else '（未识别）'}")
        if path:
            with pd.ExcelFile(path) as xls:
                for s in xls.sheet_names[:12]:
                    cols = [str(c).strip() for c in pd.read_excel(xls, sheet_name=s, nrows=0).columns]
                    log(f"   「{s}」({len(cols)}列): " + " | ".join(cols[:18]))


def main():
    global COLUMN_ALIASES, ANNOTATION_COLS
    COLUMN_ALIASES, ANNOTATION_COLS = load_aliases()
    ap = argparse.ArgumentParser(description="应收账款合并 S1–S8")
    ap.add_argument("--source"); ap.add_argument("--ref"); ap.add_argument("--rules")
    ap.add_argument("--out"); ap.add_argument("--base-month")
    ap.add_argument("--input-dir", default=WORK_INPUT)
    ap.add_argument("--inspect", action="store_true")
    a = ap.parse_args()
    if a.inspect:
        inspect_mode(a.input_dir); return
    source, ref, rules = a.source, a.ref, a.rules
    if not source:  # 没显式给 → 扫工作区按内容认
        source, ref, rules = find_inputs(a.input_dir)
        if not source:
            log(f"✗ 没找到源台账。请把文件放进 {a.input_dir}/ 或用 --source 指定。"); sys.exit(1)
    rules = rules or (os.path.join(CONFIG_DIR, "营销人员应收匹配规则.xlsx")
                      if os.path.isfile(os.path.join(CONFIG_DIR, "营销人员应收匹配规则.xlsx")) else None)
    out_path = a.out or os.path.join(WORK_OUTPUT, f"应收all_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx")
    run(source, ref, rules, out_path, a.base_month)


if __name__ == "__main__":
    main()
