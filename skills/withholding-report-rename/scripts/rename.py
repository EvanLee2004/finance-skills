#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
代扣代缴申报表重命名 (withholding-report-rename)

把「代扣代缴、代收代缴税款报告表」PDF 批量按规则重命名：

    新文件名 = {被代扣代缴代收代缴纳税人名称}{计税依据合计 + 实代扣代缴代收代缴税额合计}.pdf

数据来源（PDF 由税务系统导出，固定 2 页）：
  · 第 1 页合计行（含「合」或「总」字的行）：列 5 = 计税依据合计，列 10 = 实缴税额合计
  · 第 2 页数据行：列 2 = 纳税人名称（单元格内被换行切碎，需合并再按业务词表分词）

默认 copy 模式：改好名的副本写进输出文件夹，**原 PDF 原封不动**；
同时产出「对照表.csv」便于复核 / 回滚。抽不到名/金额或名称可疑的，
进「待人工」清单、**不**改名，留给人确认——绝不瞎命名。

用法：
    python3 rename.py --input <PDF所在文件夹> [--out-dir <输出夹>] \
        [--mode copy|rename] [--dry-run]
"""
import argparse
import csv
import datetime
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
CONFIG_DIR = os.path.join(SKILL_DIR, "config")

# ── pdfplumber 守卫：缺了给人话，别裸崩 ─────────────────────────────
try:
    import pdfplumber
except ImportError:
    print("✗ 缺 pdfplumber。装：pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pdfplumber")
    sys.exit(1)


# ── 业务词表兜底（config/名称分词词表.md 缺失时用这批） ──────────────
DEFAULT_SUFFIX_WORDS = [
    'MANAGEMENT', 'CONSULTANCIES', 'CONSULTANCY', 'CONSULTING',
    'LIMITED', 'INCORPORATED', 'CORPORATION',
    'COMPANY', 'HOLDINGS', 'INTERNATIONAL', 'ENTERPRISES',
    'SERVICES', 'SOLUTIONS', 'TECHNOLOGIES', 'TECHNOLOGY',
    'LOGISTICS', 'TRADING', 'DEVELOPMENT', 'INVESTMENT',
    'PARTNERS', 'ENGINEERING', 'CONSTRUCTION', 'TRANSPORTATION',
    'COMMUNICATIONS', 'INDUSTRIES', 'RESOURCES', 'PROPERTIES',
    'TRANSLATIONS', 'TRANSLATION', 'PRODUCTIONS', 'RECORDING',
    'STUDIO', 'STUDIOS', 'PRODUCTION', 'OFFSHORE',
]

# Windows 文件名非法字符（macOS 也一并清掉，跨系统安全）
ILLEGAL = '<>:"/\\|?*'


def log(msg):
    """中文路径在 Windows GBK 控制台可能 UnicodeEncodeError，包一层。"""
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode('utf-8', 'replace').decode('ascii', 'replace'))


# ── 读 config 活表 ──────────────────────────────────────────────────
def load_suffix_words():
    """从 config/名称分词词表.md 读业务词（一行一个全大写词），缺了用兜底。"""
    path = os.path.join(CONFIG_DIR, "名称分词词表.md")
    words = []
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                w = line.strip().lstrip('-*').strip()
                if re.fullmatch(r'[A-Z]{3,}', w):
                    words.append(w)
    return words or list(DEFAULT_SUFFIX_WORDS)


def load_overrides():
    """从 config/名称修正表.md 读「原文件名 → 正确新名」手工映射（markdown 表）。

    键归一化为去掉扩展名的原文件名。算法搞不定的疑难名在这里钉死。
    """
    path = os.path.join(CONFIG_DIR, "名称修正表.md")
    table = {}
    if not os.path.exists(path):
        return table
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line.startswith('|') or '<!--' in line:
                continue
            cells = [c.strip() for c in line.strip('|').split('|')]
            if len(cells) < 2:
                continue
            src, new = cells[0], cells[1]
            # 跳过表头 / 分隔行 / 空行
            if not src or src.startswith('-') or src in ('原文件名', '原文件名（含或不含.pdf）'):
                continue
            if not new or new.startswith('-') or new in ('正确新名', '正确新名（不含.pdf）'):
                continue
            key = re.sub(r'\.pdf$', '', src, flags=re.I)
            table[key] = re.sub(r'\.pdf$', '', new, flags=re.I)
    return table


# ── 名称重构 ───────────────────────────────────────────────────────
def build_suffix_pattern(words):
    return '|'.join(f'({re.escape(w)})' for w in sorted(words, key=len, reverse=True))


def extract_name(cell_text, suffix_pattern):
    """PDF 第2页列2单元格 -> 完整纳税人名称。

    单元格里英文被换行切碎（LUCALIZ / E / MANAGEM / ENT ...），先去换行合并，
    清标点瑕疵，再在已知业务词前插空格还原词边界。没匹配到业务词就原样返回
    （如 WordPowerS.r.l. —— 这是正确的，不算可疑）。
    """
    s = cell_text.replace('\n', '')
    s = s.replace('’', "'").replace('‘', "'").replace('“', '"').replace('”', '"')
    s = re.sub(r',\.', '.', s)
    s = re.sub(r'\.{2,}', '.', s)
    s = re.sub(r'\.,', '.', s)
    s = s.strip().rstrip('.')
    matches = list(re.finditer(suffix_pattern, s))
    if not matches:
        return s
    parts, prev = [], 0
    for m in matches:
        pre = s[prev:m.start()]
        if pre:
            parts.append(pre)
        parts.append(m.group())
        prev = m.end()
    rest = s[prev:]
    if rest:
        parts.append(rest)
    name = ' '.join(parts)
    name = re.sub(r'\s+', ' ', name).strip().rstrip('.')
    name = re.sub(r'(CO\.) (L\.L\.C)', r'\1\2', name)
    return name


def find_name(pdf, suffix_pattern):
    """在第2页表格里找纳税人名称：列索引>=2、含3+英文字母、非纯数字识别号。"""
    if len(pdf.pages) < 2:
        return None
    for table in pdf.pages[1].extract_tables():
        for row in table:
            for ci, cell in enumerate(row):
                if ci < 2 or not cell:
                    continue
                text = str(cell).replace('\n', '')
                if not re.search(r'[A-Za-z]{3,}', text):
                    continue
                if re.match(r'^[0-9]{6,}$', text):  # 识别号那列跳过
                    continue
                return extract_name(str(cell), suffix_pattern)
    return None


def get_totals(pdf):
    """第1页合计行 -> (计税依据合计, 实缴税额合计)。找不到返回 (None, None)。"""
    if not pdf.pages:
        return None, None
    for table in pdf.pages[0].extract_tables():
        for row in table:
            if len(row) < 11:
                continue
            first = str(row[0]) if row[0] else ''
            if '合' in first or '总' in first:
                basis = _num(row[5])
                tax = _num(row[10])
                return basis, tax
    return None, None


def _num(cell):
    if cell is None:
        return None
    s = str(cell).replace(',', '').strip()
    if s in ('', '--', '-'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return None


def fmt_amount(total):
    """金额去尾零：1500.40 -> 1500.4，84.00 -> 84。"""
    return f"{total:.2f}".rstrip('0').rstrip('.')


def sanitize(name):
    for ch in ILLEGAL:
        name = name.replace(ch, '')
    return name.strip()


# ── 单文件 -> 计划 ─────────────────────────────────────────────────
def plan_one(path, suffix_pattern, overrides):
    """返回 dict：src/newbase/name/basis/tax/total/status/note。status: ok | manual。"""
    base = os.path.basename(path)
    key = re.sub(r'\.pdf$', '', base, flags=re.I)
    rec = {"src": path, "name": "", "basis": "", "tax": "", "total": "",
           "newbase": "", "status": "ok", "note": ""}

    if key in overrides:
        rec["newbase"] = sanitize(overrides[key])
        rec["name"] = "(名称修正表指定)"
        rec["note"] = "名称修正表命中"
        return rec

    try:
        with pdfplumber.open(path) as pdf:
            name = find_name(pdf, suffix_pattern)
            basis, tax = get_totals(pdf)
    except Exception as e:  # 坏 PDF 不连坐其它文件
        rec["status"] = "manual"
        rec["note"] = f"打开/解析失败: {e}"
        return rec

    rec["name"] = name or ""
    rec["basis"] = "" if basis is None else fmt_amount(basis)
    rec["tax"] = "" if tax is None else fmt_amount(tax)

    # 人在环：缺名 / 缺金额 / 名称可疑 -> 待人工，不瞎命名
    if not name:
        rec["status"], rec["note"] = "manual", "没抽到纳税人名称"
        return rec
    if basis is None or tax is None:
        rec["status"], rec["note"] = "manual", "没抽到合计金额（合计行缺失或金额异常）"
        return rec
    if '?' in name or len(name.strip()) < 2:
        rec["status"], rec["note"] = "manual", f"名称可疑：{name!r}"
        return rec

    total = round(basis + tax, 2)
    rec["total"] = fmt_amount(total)
    rec["newbase"] = sanitize(name + rec["total"])
    return rec


def dedup(plans):
    """同名冲突加 _1/_2 后缀（只对 ok 的）。"""
    used = {}
    for p in plans:
        if p["status"] != "ok":
            continue
        nb = p["newbase"]
        if nb in used:
            used[nb] += 1
            p["newbase"] = f"{nb}_{used[nb]}"
            p["note"] = (p["note"] + "；" if p["note"] else "") + "重名已加后缀"
        else:
            used[nb] = 0


# ── 主流程 ─────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="代扣代缴申报表 PDF 批量重命名")
    ap.add_argument("--input", required=True, help="待重命名 PDF 所在文件夹")
    ap.add_argument("--out-dir", help="输出夹（copy 模式默认放 input 同级的 重命名_<日期>/）")
    ap.add_argument("--mode", choices=["copy", "rename"], default="copy",
                    help="copy=改好名的副本写新夹、原件不动（默认）；rename=就地改名（先自动备份）")
    ap.add_argument("--dry-run", action="store_true", help="只算不写，打印计划")
    args = ap.parse_args()

    in_dir = os.path.abspath(args.input)
    if not os.path.isdir(in_dir):
        log(f"✗ 输入不是文件夹：{in_dir}")
        sys.exit(1)

    pdfs = sorted(os.path.join(in_dir, f) for f in os.listdir(in_dir)
                  if f.lower().endswith(".pdf"))
    if not pdfs:
        log(f"✗ {in_dir} 里没有 PDF")
        sys.exit(1)

    suffix_pattern = build_suffix_pattern(load_suffix_words())
    overrides = load_overrides()

    plans = [plan_one(p, suffix_pattern, overrides) for p in pdfs]
    dedup(plans)

    ok = [p for p in plans if p["status"] == "ok"]
    manual = [p for p in plans if p["status"] != "ok"]

    log(f"共 {len(plans)} 个 PDF：可重命名 {len(ok)}，待人工确认 {len(manual)}")
    for p in plans:
        tag = "✓" if p["status"] == "ok" else "⚠ 待人工"
        log(f"  {tag}  {os.path.basename(p['src'])}")
        log(f"        -> {p['newbase'] + '.pdf' if p['status']=='ok' else '（' + p['note'] + '）'}")

    date = datetime.date.today().strftime("%Y%m%d")

    if args.dry_run:
        log("\n[dry-run] 没动任何文件。")
        return

    # 输出夹
    if args.out_dir:
        out_dir = os.path.abspath(args.out_dir)
    else:
        out_dir = os.path.join(os.path.dirname(in_dir), f"重命名_{date}")

    if args.mode == "rename":
        # 先整批备份，再就地改名
        backup = os.path.join(os.path.dirname(in_dir), f"申报表_备份_{date}")
        os.makedirs(backup, exist_ok=True)
        for p in pdfs:
            shutil.copy2(p, os.path.join(backup, os.path.basename(p)))
        log(f"\n已备份原件 -> {backup}")
        for p in ok:
            dst = os.path.join(in_dir, p["newbase"] + ".pdf")
            os.rename(p["src"], dst)
        report_dir = in_dir
        log(f"就地改名完成：{len(ok)} 个（待人工的 {len(manual)} 个原样保留）。")
    else:
        os.makedirs(out_dir, exist_ok=True)
        for p in ok:
            shutil.copy2(p["src"], os.path.join(out_dir, p["newbase"] + ".pdf"))
        report_dir = out_dir
        log(f"\n副本已写入 -> {out_dir}（{len(ok)} 个；原件未动）。")

    # 对照表 CSV（含待人工，便于复核 / 回滚）
    csv_path = os.path.join(report_dir, "对照表.csv")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["原文件名", "新文件名", "纳税人名称", "计税依据", "实缴税额", "合计金额", "状态", "备注"])
        for p in plans:
            w.writerow([
                os.path.basename(p["src"]),
                (p["newbase"] + ".pdf") if p["status"] == "ok" else "",
                p["name"], p["basis"], p["tax"], p["total"],
                "可重命名" if p["status"] == "ok" else "待人工",
                p["note"],
            ])
    log(f"对照表 -> {csv_path}")

    if manual:
        log(f"\n⚠ 有 {len(manual)} 个没改名（待人工确认），见对照表「待人工」行：")
        for p in manual:
            log(f"    {os.path.basename(p['src'])} —— {p['note']}")


if __name__ == "__main__":
    main()
