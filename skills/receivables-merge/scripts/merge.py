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
  python3 merge.py --source 源台账.xlsx --ref 回填源.xlsx --rules config/销售归属维护表.md --out 应收all.xlsx
  python3 merge.py --inspect            # 只打印识别到的输入与表头
其中 --ref / --rules 可省（省 ref 跳过回填+结转；省 rules 跳过归属；--rules 默认用 config/销售归属维护表.md）。
缺文件/格式不对会清晰报错或优雅跳过，不裸崩。
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
    # 整批"技术统一确认"的 sheet：交付月份用统一确认月(固定值)，账龄从此算——
    # 不按每单"完成时间"各算（赵成品实证：6月批量整批交付月=202406）。
    "SHEET_DELIVER_MONTH": {"6月批量": "202406"},
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


def load_business_rules():
    """读 config/业务规则.md 的『特殊批次』『跳过的 sheet』表 → 覆盖 CONFIG 默认。
       缺文件/解析失败 → 保留内置默认，绝不崩。让"会变的批次规则"成为活的 MD。"""
    p = os.path.join(CONFIG_DIR, "业务规则.md")
    if not os.path.isfile(p):
        return
    try:
        with open(p, encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        log(f"⚠ 读 业务规则.md 失败({e})，用内置默认。")
        return
    ylab, deliver, exclude = {}, {}, []
    for part in re.split(r"(?m)^##\s+", text):
        ls = part.splitlines()
        head = ls[0] if ls else ""
        rows = _parse_md_table(ls)
        if "特殊批次" in head:
            for c in rows:
                name = (c[0] if len(c) > 0 else "").strip()
                lab = (c[1] if len(c) > 1 else "").strip()
                dm = (c[2] if len(c) > 2 else "").strip()
                if name and lab:
                    ylab[name] = lab
                if name and re.fullmatch(r"\d{6}", dm):
                    deliver[name] = dm
        elif "跳过" in head:
            for c in rows:
                if c and c[0].strip():
                    exclude.append(c[0].strip())
    if ylab:
        CONFIG["SHEET_YEAR_LABEL"] = ylab
    if deliver:
        CONFIG["SHEET_DELIVER_MONTH"] = deliver
    if exclude:
        CONFIG["EXCLUDE_SHEETS"] = exclude

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
    """提取判别特征：标注列种类数 / 是否含新智云单号 / 是否含年度列 / 年份分表数。"""
    ann = 0
    has_key = has_year_col = False
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
                hits = sum(1 for kws in ANNOTATION_COLS.values()
                           if any(any(kw in c for c in cols) for kw in kws))
                ann = max(ann, hits)
    except Exception:
        pass
    return {"ann": ann, "key": has_key, "ycol": has_year_col, "years": year_sheets}


def find_inputs(input_dir):
    """从一个目录按内容认出【源台账 / 回填源 / 维护表】，不靠文件名。
       维护表是 .md（一般放 config/，不在 input；此处兜底）；源台账/回填源是 .xlsx。"""
    if not os.path.isdir(input_dir):
        return None, None, None
    allf = sorted(os.listdir(input_dir))
    cands = [os.path.join(input_dir, f) for f in allf
             if f.lower().endswith(EXCEL_EXTS) and not f.startswith(("~$", "."))]
    md = [os.path.join(input_dir, f) for f in allf if f.lower().endswith(".md")]
    rules = md[0] if md else None       # 维护表只认 .md，绝不把 xlsx 当 rules（否则 MD 解析器会吃二进制崩）
    if not cands:
        return None, None, rules
    feats = {p: _scan_features(p) for p in cands}
    # 回填源：含标注列(>=3)或含年度列
    ref_cands = [p for p in cands if feats[p]["ann"] >= 3 or feats[p]["ycol"]]
    ref = max(ref_cands, key=lambda p: (feats[p]["ann"], feats[p]["ycol"])) if ref_cands else None
    # 源台账：剩下里年份分表最多的
    src_cands = [p for p in cands if p != ref] or cands
    source = max(src_cands, key=lambda p: feats[p]["years"]) if src_cands else None
    if ref is not None and source is not None and os.path.abspath(ref) == os.path.abspath(source):
        ref = None
    return source, ref, rules


# --------------------------- S1 合并 ---------------------------
def normalize_sheet(df, sheet_name):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    out = pd.DataFrame(index=df.index)
    matched = set()
    for canon, names in COLUMN_ALIASES.items():
        hit = next((n for n in names if n in df.columns), None)
        if hit is None:
            out[canon] = None
            continue
        matched.add(canon)
        col = df[hit]
        if isinstance(col, pd.DataFrame):   # 源表有重名列 → 取第一列，避免赋值崩
            col = col.iloc[:, 0]
        out[canon] = col
    ym1 = out["_交付月份直取"].map(to_yyyymm)
    ym2 = out["_交付日期回退"].map(to_yyyymm)
    out["交付月份"] = ym1.where(ym1.notna(), ym2)
    out["交付月份"] = out["交付月份"].map(lambda x: x if isinstance(x, str) else None)
    out.drop(columns=["_交付月份直取", "_交付日期回退"], inplace=True)
    # 整批"技术统一确认"的 sheet（如6月批量）：交付月份用固定的统一确认月，覆盖逐单推导。
    fixed_dm = CONFIG["SHEET_DELIVER_MONTH"].get(str(sheet_name).strip())
    if fixed_dm:
        out["交付月份"] = fixed_dm
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
    # 关键列认列自检：关键列没认出 → 这批数据会静默丢失/算错，必须告警（不能默默丢）。
    missing = [c for c in ("销售人员", "客户名称", "应收金额") if c not in matched]
    if "新智云单号" not in matched and "单号" not in matched:
        missing.append("新智云单号/单号(主键)")
    return out, missing


def load_and_merge(source_path):
    frames, report, col_warn = [], [], []
    xls = pd.ExcelFile(source_path)
    try:
        for s in list(xls.sheet_names):
            if s in CONFIG["EXCLUDE_SHEETS"]:
                report.append((s, "跳过", 0)); continue
            sn = str(s).strip()
            # 未知 sheet 告警：不是4位年份、又不在"特殊批次"表里 → 别猜，问人
            if not re.fullmatch(r"\d{4}", sn) and sn not in CONFIG["SHEET_YEAR_LABEL"]:
                col_warn.append({"sheet": sn, "问题": "未知 sheet（非年份、非已登记特殊批次）",
                                 "建议": "问人：是不是新的特殊批次？年度记啥、整批交付月哪个月？补进 业务规则.md"})
                log(f"⚠ 未知 sheet「{sn}」：非年份、非已登记特殊批次——按普通年份处理可能出错，建议问人。")
            df = pd.read_excel(xls, sheet_name=s)
            if df.empty:
                report.append((s, "空表", 0)); continue
            norm, missing = normalize_sheet(df, s)
            frames.append(norm)
            if missing:
                col_warn.append({"sheet": sn, "问题": f"关键列没认出来：{'、'.join(missing)}",
                                 "建议": "源表可能改了列名 → 这批数据会丢失/算错；在 列名别名.json 把新列名加进去"})
                log(f"⚠ sheet「{sn}」关键列没认出：{'、'.join(missing)}（这批数据可能丢失/算错！）")
            n_fb = int(norm["__主键回退"].sum())
            report.append((s, f"年度={year_label(s)}" + (f"; 主键回退{n_fb}行" if n_fb else ""), len(norm)))
    finally:
        xls.close()
    if not frames:
        raise RuntimeError("源台账没有可合并的 sheet。")
    return pd.concat(frames, ignore_index=True), report, col_warn


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


# --------------------------- S4.5 结转老坏账（从上一版 all） ---------------------------
def carry_forward(master_keys, ref_path, min_source_year):
    """从上一版 all 结转老坏账：ref 里【年度早于源表最早年份(min_source_year) 且 应收≠0】的行——
    这些是源台账已不再列出那些老年份(如2016-2018高美杰遗留)、但还没回款的老账。
    近期年份(源表覆盖范围内)若不在源表，说明已回款结清，**不结转**。
    账龄/销售归属/标注沿用上一版 all（它已是最终态，成品也不给结转行重算账龄）。
    返回 17 列(+__carried) 的 DataFrame。"""
    ref_sheet = CONFIG["REF_SHEET"] or detect_ref_sheet(ref_path)
    d = pd.read_excel(ref_path, sheet_name=ref_sheet)
    d.columns = [str(c).strip() for c in d.columns]
    keycol = CONFIG["KEY"] if CONFIG["KEY"] in d.columns else next((c for c in d.columns if "新智云" in c), None)
    if keycol is None:
        return pd.DataFrame()
    agingcol = next((c for c in d.columns if "账龄" in c), None)
    ann_src = {}
    for canon, kws in ANNOTATION_COLS.items():
        ann_src[canon] = (next((c for c in d.columns if c == canon), None)
                          or next((c for c in d.columns if any(kw in c for kw in kws)), None))
    rows = []
    for _, r in d.iterrows():
        key = clean_key(r.get(keycol))
        if not key or key in master_keys:
            continue
        # 只结转源表年份范围【之外】的老年份；近期年份不在源表=已结清，不结转。
        y = r.get("年度")
        try:
            yi = int(str(y).strip())
        except (TypeError, ValueError):
            continue
        if yi >= min_source_year:
            continue
        amt = pd.to_numeric(r.get("应收金额"), errors="coerce")
        if pd.isna(amt) or amt == 0:
            continue
        rec = {
            "年度": yi, "销售人员": r.get("销售人员"), "客户名称": r.get("客户名称"),
            "新智云单号": key, "文件名": r.get("文件名"), "应收金额": amt,
            "交付月份": to_yyyymm(r.get("交付月份")),
            "账龄(月份)": r.get(agingcol) if agingcol else None,   # 沿用上一版账龄(成品不重算)
            "__matched": True, "__carried": True,
        }
        for canon, src in ann_src.items():
            rec[canon] = blankify_annotation(r.get(src)) if src else None
        rows.append(rec)
    return pd.DataFrame(rows)


# --------------------------- S4 销售归属（维护表驱动） ---------------------------
def _base_name(s):
    """去结尾数字：高美杰1 → 高美杰。"""
    return re.sub(r"\d+$", "", str(s).strip()).strip()


def _parse_md_table(lines):
    """从若干行里抽出 markdown 表格的【数据行】(每行=单元格列表)，自动跳表头行与 |---| 分隔行。"""
    rows, seen_sep = [], False
    for ln in lines:
        s = ln.strip()
        if not s.startswith("|"):
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if cells and all(c and set(c) <= set("-: ") for c in cells):   # |---|---| 分隔行
            seen_sep = True
            continue
        if not seen_sep:                                               # 分隔行之前 = 表头行，跳过
            continue
        rows.append(cells)
    return rows


def load_rules(rules_path):
    """读【MD 维护表】→ (departed_map, customer_map, composite_set)。
       段一『按销售接手』：离职|接手|处理方式|备注；处理方式含"细分/追溯/复合" → 进 composite_set。
       段二『按公司接手』：客户(含关键词)|接手|备注。
       防御：缺文件/非.md/解码失败 → 警告并返回空(跳过归属)，绝不崩。"""
    empty = ({}, {}, set())
    if not rules_path or not os.path.isfile(rules_path):
        log(f"⚠ 维护表不存在，跳过销售归属：{rules_path}")
        return empty
    if not str(rules_path).lower().endswith((".md", ".markdown", ".txt")):
        log(f"⚠ 维护表应为 .md 文本（收到 {os.path.basename(rules_path)}），跳过销售归属。")
        return empty
    try:
        with open(rules_path, encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        log(f"⚠ 读维护表失败({e})，跳过销售归属。")
        return empty
    departed_map, customer_map, composite_set = {}, {}, set()
    for part in re.split(r"(?m)^##\s+", text):          # 按 ## 标题切段
        ls = part.splitlines()
        head = ls[0] if ls else ""
        rows = _parse_md_table(ls)
        if "销售接手" in head:
            for c in rows:
                dep = c[0] if len(c) > 0 else ""
                to = c[1] if len(c) > 1 else ""
                mode = c[2] if len(c) > 2 else ""
                if dep and to:
                    departed_map[dep] = to
                    if any(k in mode for k in ("细分", "追溯", "复合")):
                        composite_set.add(dep)
        elif "公司接手" in head:
            for c in rows:
                cust = c[0] if len(c) > 0 else ""
                to = c[1] if len(c) > 1 else ""
                if cust and to:
                    customer_map[cust] = to
    return departed_map, customer_map, composite_set


def attribute_sales(merged, departed_map, customer_map, composite_set):
    """对每行定最终销售人员（维护表驱动，确定性应用；不写死任何人名）：
       · 只动【离职销售】的行——在职销售当前手里的单一律不动；
       · 接手人 = 段二客户规则(优先，对所有离职行生效) 否则 段一默认接手；
       · 命名：处理方式=细分追溯(composite_set) → 复合名「接手-离职」(接手是其变体如高美杰1→光名)；
              处理方式=直接接手 → 接手人光名。
       composite_set 只决定【命名】，由维护表"处理方式"列驱动，规则变=改表不改码。
       返回(新merged, 改动清单df)。"""
    merged = merged.copy()
    changes = []
    cust_keys = sorted([k for k in customer_map if k], key=len, reverse=True)  # 长键优先
    new_vals = []
    for orig_raw, cust_raw in zip(merged["销售人员"], merged["客户名称"]):
        orig = "" if orig_raw is None or (isinstance(orig_raw, float) and pd.isna(orig_raw)) else str(orig_raw).strip()
        cust = "" if cust_raw is None or (isinstance(cust_raw, float) and pd.isna(cust_raw)) else str(cust_raw).strip()
        # 段二客户规则：对【所有行】生效（这些客户整体归指定人，不管现在谁持有）。
        owner = next((customer_map[k] for k in cust_keys if k and k in cust), None)
        if owner is None and orig in departed_map:     # 段一：离职销售默认接手
            owner = departed_map[orig]
        if owner is None:                              # 在职且无客户规则 → 不动
            new_vals.append(orig_raw); continue
        if orig in composite_set:                      # 处理方式=细分追溯 → 复合名
            final = owner if _base_name(owner) == _base_name(orig) else f"{owner}-{orig}"
        else:                                          # 处理方式=直接接手 → 光名
            final = owner
        new_vals.append(final)
        if final != orig:
            changes.append({"原销售": orig, "客户名称": cust, "最终销售": final})
    merged["销售人员"] = new_vals
    return merged, pd.DataFrame(changes, columns=["原销售", "客户名称", "最终销售"])


# --------------------------- S5 删 0 行 ---------------------------
def drop_zero_receivable(merged):
    """删掉应收金额=0 或 空 的行（已回款/核销）。返回(保留, 被删df)。
    区分『真的0=已回款』vs『解析不出=数据异常』，别让数据异常被误当已回款。"""
    amt = pd.to_numeric(merged["应收金额"], errors="coerce")
    keep_mask = amt.notna() & (amt != 0)
    removed = merged.loc[~keep_mask].copy()
    removed["删除原因"] = amt.loc[~keep_mask].map(
        lambda v: "应收=0(已回款/核销)" if pd.notna(v) else "应收为空/非数字(数据异常?待查)")
    n_nan = int(amt.loc[~keep_mask].isna().sum())
    if n_nan:
        log(f"⚠ {n_nan} 行应收金额解析不出(空/非数字)被删——可能数据异常，查『被删0行』删除原因列。")
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


def _lev1(a, b):
    """编辑距离是否==1（错别字检测，如 尹博健 vs 尹博建）。"""
    a, b = str(a), str(b)
    if a == b:
        return False
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if la == lb:
        return sum(1 for x, y in zip(a, b) if x != y) == 1
    short, long = (a, b) if la < lb else (b, a)
    i = j = diff = 0
    while i < len(short) and j < len(long):
        if short[i] == long[j]:
            i += 1; j += 1
        else:
            j += 1; diff += 1
            if diff > 1:
                return False
    return True


def suspect_attribution(orig_counts, departed_names):
    """归属存疑：源表里出现、与维护表离职名仅差一字的（疑似同一人错别字），交人工确认。"""
    deps = list(departed_names)
    rows = []
    for name, cnt in orig_counts.items():
        nm = str(name).strip()
        if not nm or nm in departed_names:
            continue
        near = [d for d in deps if _lev1(nm, d)]
        if near:
            rows.append({"源表销售": nm, "行数": int(cnt),
                         "存疑": f"与维护表离职名「{near[0]}」仅差一字，是否同一人(错别字)?",
                         "建议": "人工确认；若是→改维护表为该写法或加映射"})
    return pd.DataFrame(rows, columns=["源表销售", "行数", "存疑", "建议"])


# --------------------------- S8 透视 ---------------------------
def build_pivot(master_out):
    """应收金额 按 销售人员→客户 求和（成品 Sheet2 同款）。"""
    d = master_out[["销售人员", "客户名称", "应收金额"]].copy()
    d["应收金额"] = pd.to_numeric(d["应收金额"], errors="coerce").fillna(0)
    piv = (d.groupby(["销售人员", "客户名称"], dropna=False)["应收金额"]
           .sum().reset_index().sort_values(["销售人员", "应收金额"], ascending=[True, False]))
    piv["应收金额"] = piv["应收金额"].round(2)
    return piv


# --------------------------- 输出 ---------------------------
_ILLEGAL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _strip_illegal(df):
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda v: _ILLEGAL_RE.sub("", v) if isinstance(v, str) else v)
    return df


def write_workbook(out_path, master_out, pivot, reassign, suspect, col_warn, unmatched, variants, residual, ycheck, removed, report_df):
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    sheets = {
        "主表": master_out, "透视汇总": pivot,
        "认列告警": col_warn, "归属变更": reassign, "归属存疑": suspect,
        "未匹配复核": unmatched, "名称变体": variants, "离职残留": residual,
        "年度校验": ycheck, "被删0行": removed, "运行报告": report_df,
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
    master, report, col_warn = load_and_merge(source)
    col_warn_df = pd.DataFrame(col_warn, columns=["sheet", "问题", "建议"]) if col_warn \
        else pd.DataFrame(columns=["sheet", "问题", "建议"])
    log(f"· S1 合并：{len(master)} 行" + (f"｜⚠ 认列告警 {len(col_warn)} 项" if col_warn else ""))
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

    # S4 归属（确定性应用维护表；变更与存疑交人工确认）
    departed_map, customer_map = ({}, {})
    reassign = pd.DataFrame(columns=["原销售", "客户名称", "最终销售"])
    orig_counts = merged["销售人员"].map(
        lambda x: "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()).value_counts()
    if rules:
        departed_map, customer_map, composite_set = load_rules(rules)
        merged, reassign = attribute_sales(merged, departed_map, customer_map, composite_set)
        log(f"· S4 归属：维护表 离职{len(departed_map)}人/客户{len(customer_map)}条/细分追溯{len(composite_set)}人，改动 {len(reassign)} 行")
    if len(reassign):
        reassign_sum = (reassign.groupby(["原销售", "最终销售"]).size()
                        .reset_index(name="行数").sort_values("行数", ascending=False))
    else:
        reassign_sum = pd.DataFrame(columns=["原销售", "最终销售", "行数"])
    suspect = suspect_attribution(orig_counts, set(departed_map))
    if len(suspect):
        log(f"· ⚠ 归属存疑 {len(suspect)} 项（疑似错别字/离职名对不齐），需人工确认")

    # S4.5 结转老坏账（上一版 all 里、当周源没有、应收≠0 的行）
    carried_n = 0
    if ref:
        try:
            master_keys = set(merged[CONFIG["KEY"]].map(clean_key).dropna())
            yrs = [y for y in merged["年度"] if isinstance(y, int)]
            min_year = min(yrs) if yrs else 9999
            carried = carry_forward(master_keys, ref, min_year)
            if len(carried):
                merged = pd.concat([merged, carried], ignore_index=True)
                carried_n = len(carried)
                log(f"· S4.5 结转老坏账：{carried_n} 行（上一版 all 有、当周源无、应收≠0）")
        except Exception as e:
            log(f"⚠ 结转失败跳过：{e}")

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
    # 应收金额保留 2 位小数（去浮点噪声 1369.6800000000003→1369.68；财务金额本就是分）
    master_out["应收金额"] = pd.to_numeric(master_out["应收金额"], errors="coerce").round(2)

    # 未匹配清单
    um = ~merged["__matched"].fillna(False).astype(bool) if "__matched" in merged.columns else merged[CONFIG["KEY"]].isna()
    unmatched = merged.loc[um, ["年度", "销售人员", "客户名称", "新智云单号", "应收金额", "交付月份"]].copy()

    # S8 透视
    pivot = build_pivot(master_out)

    _rm_cols = ["年度", "销售人员", "客户名称", "新智云单号", "应收金额", "交付月份", "删除原因"]
    removed_out = removed[[c for c in _rm_cols if c in removed.columns]].copy() \
        if len(removed) else pd.DataFrame(columns=_rm_cols)

    rep = [["源台账", os.path.basename(source)],
           ["回填源", os.path.basename(ref) if ref else "（无）"],
           ["维护表", os.path.basename(rules) if rules else "（无）"],
           ["账龄基准月", base_yyyymm],
           ["S1 合并行数", len(master)],
           ["⚠ 认列告警(关键列/未知sheet)", len(col_warn)],
           ["S4 归属改动行数", len(reassign)],
           ["S4 归属存疑(待人工确认)", len(suspect)],
           ["S4.5 结转老坏账行数", carried_n],
           ["S5 删0行数", len(removed)],
           ["主表最终行数", len(master_out)],
           ["未匹配(#N/A)", len(unmatched)],
           ["名称变体待复核", len(variants)],
           ["离职光名残留", len(residual)],
           ["年度不一致", len(ycheck)]]
    report_df = pd.DataFrame(rep, columns=["项目", "值"])

    write_workbook(out_path, master_out, pivot, reassign_sum, suspect, col_warn_df, unmatched, variants, residual, ycheck, removed_out, report_df)
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


def check_mode(source, ref, rules):
    """预检：依赖 + 配置可解析 + 输入存在。跑正活前先 --check，问题早暴露。返回是否全过。"""
    ok = True
    log("=== 预检 (--check) ===")
    for mod in ("pandas", "openpyxl"):
        try:
            __import__(mod); log(f"  ✓ 依赖 {mod}")
        except ImportError:
            log(f"  ✗ 缺依赖 {mod}  → pip install {mod}"); ok = False
    for f in ("列名别名.json", "业务规则.md", "销售归属维护表.md"):
        p = os.path.join(CONFIG_DIR, f)
        log(f"  {'✓' if os.path.isfile(p) else '⚠'} 配置 {f}" + ("" if os.path.isfile(p) else "（缺 → 用内置默认/跳过该步）"))
    try:
        load_business_rules()
        log(f"  ✓ 业务规则.md 可解析（特殊批次 {len(CONFIG['SHEET_DELIVER_MONTH'])} 个、跳过 {len(CONFIG['EXCLUDE_SHEETS'])} sheet）")
    except Exception as e:
        log(f"  ✗ 业务规则.md 解析失败：{e}"); ok = False
    rp = rules or os.path.join(CONFIG_DIR, "销售归属维护表.md")
    if os.path.isfile(rp):
        try:
            dm, cm, cs = load_rules(rp)
            log(f"  ✓ 维护表可解析（离职 {len(dm)} / 客户 {len(cm)} / 细分追溯 {len(cs)}）")
        except Exception as e:
            log(f"  ✗ 维护表解析失败：{e}"); ok = False
    for tag, p in [("源台账", source), ("回填源", ref)]:
        if p:
            exists = os.path.isfile(p)
            log(f"  {'✓' if exists else '✗'} {tag}：{p}")
            ok = ok and exists
    log("=== " + ("预检通过 ✓ 可以跑" if ok else "预检发现问题 ✗ 先解决") + " ===")
    return ok


def main():
    global COLUMN_ALIASES, ANNOTATION_COLS
    COLUMN_ALIASES, ANNOTATION_COLS = load_aliases()
    load_business_rules()    # 用活的 业务规则.md 覆盖特殊批次/跳过sheet（缺则用默认）
    ap = argparse.ArgumentParser(description="应收账款合并 S1–S8")
    ap.add_argument("--source"); ap.add_argument("--ref"); ap.add_argument("--rules")
    ap.add_argument("--out"); ap.add_argument("--base-month")
    ap.add_argument("--input-dir", default=WORK_INPUT)
    ap.add_argument("--inspect", action="store_true")
    ap.add_argument("--check", action="store_true", help="预检：依赖/配置/输入，不跑正活")
    a = ap.parse_args()
    if a.inspect:
        inspect_mode(a.input_dir); return
    source, ref, rules = a.source, a.ref, a.rules
    if a.check:
        sys.exit(0 if check_mode(source, ref, rules) else 1)
    if not source:  # 没显式给 → 扫工作区按内容认
        source, ref, rules = find_inputs(a.input_dir)
        if not source:
            log(f"✗ 没找到源台账。请把文件放进 {a.input_dir}/ 或用 --source 指定。"); sys.exit(1)
    rules = rules or (os.path.join(CONFIG_DIR, "销售归属维护表.md")
                      if os.path.isfile(os.path.join(CONFIG_DIR, "销售归属维护表.md")) else None)
    # —— 预检：清晰报错而非裸崩 ——
    if source and not os.path.isfile(source):
        log(f"✗ 源台账文件不存在：{source}"); sys.exit(1)
    if ref and not os.path.isfile(ref):
        log(f"⚠ 回填源不存在，跳过回填+结转：{ref}"); ref = None
    if a.base_month and not (re.fullmatch(r"\d{6}", str(a.base_month)) and 1 <= int(str(a.base_month)[4:6]) <= 12):
        log(f"✗ --base-month 应为 YYYYMM 六位合法月份（收到 {a.base_month}）。"); sys.exit(1)
    out_path = a.out or os.path.join(WORK_OUTPUT, f"应收all_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx")
    run(source, ref, rules, out_path, a.base_month)


if __name__ == "__main__":
    main()
