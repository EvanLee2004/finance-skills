#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
无损写单元格：**只补丁目标 sheet 的那几个格，工作簿其余部件逐字节原样保留**。

为什么不用 openpyxl 保存：实测拿明妹的真表（`盈亏核算表2026全年`）用 openpyxl 载入再保存，
74 个部件变 59 —— **5 个 drawing、1 张内嵌图片、若干 rels 全部丢失**。
她一打开发现图没了，"你们把我的表搞坏了"，这个项目就别想推下去了。

做法（与 `xlsx` 技能的 unpack/pack 同思路，OOXML 外科手术）：
  1. 把 xlsx 当 zip 打开，只取出目标 sheet 的 XML
  2. 在 XML 里就地改那几个 `<c>` 单元格
  3. 重新打包：**除该 sheet 外的每个部件都按原始字节写回**

样式沿用：空格子没有样式时，从**同一列另一个有值的行**借样式索引，
这样日期/金额显示格式跟她表里其余行一致，不会出现"一列里有的显示日期有的显示数字"。
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

# Excel 日期序列号的零点。1900 历法有个著名的假闰日（1900-02-29），
# 所以 1900-03-01 之后要 +1；用 1899-12-30 当基准正好把这个偏差吸收掉。
_EXCEL_EPOCH = dt.date(1899, 12, 30)

_CELL_RE_TMPL = r'<c r="{ref}"(?P<attrs>[^>]*?)(?:/>|>(?P<inner>.*?)</c>)'
_ROW_RE_TMPL = r'<row[^>]*\br="{row}"[^>]*(?:/>|>.*?</row>)'


def col_letter(idx1: int) -> str:
    """1 → A，27 → AA。"""
    s = ""
    while idx1 > 0:
        idx1, r = divmod(idx1 - 1, 26)
        s = chr(65 + r) + s
    return s


def to_serial(d) -> float:
    """日期 → Excel 序列号。她的表要能按时间排序筛选，所以必须写成真日期值。"""
    if isinstance(d, dt.datetime):
        d = d.date()
    return float((d - _EXCEL_EPOCH).days)


def sheet_path_for(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """按 sheet 名找到它在 zip 里的 XML 路径（走 workbook.xml + rels，不猜文件名）。"""
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    rid = None
    # 属性顺序在不同产出工具里不一样（Excel 与 openpyxl 就不同），逐个 <sheet> 解析，别假设顺序
    for m in re.finditer(r"<sheet\b([^>]*)/?>", wb_xml):
        attrs = m.group(1)
        nm = re.search(r'\bname="([^"]*)"', attrs)
        rid_m = re.search(r'\br:id="([^"]+)"', attrs) or re.search(r'\bid="([^"]+)"', attrs)
        if nm and nm.group(1) == sheet_name and rid_m:
            rid = rid_m.group(1)
            break
    if rid is None:
        raise ValueError(f"workbook.xml 里没有名为「{sheet_name}」的 sheet")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", rels):
        attrs = m.group(1)
        i = re.search(r'\bId="([^"]+)"', attrs)
        tg = re.search(r'\bTarget="([^"]+)"', attrs)
        if i and tg and i.group(1) == rid:
            target = tg.group(1).lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    raise ValueError(f"找不到 {rid} 对应的 sheet 文件")


# 内建日期格式 id：14=短日期、22=日期时间。**只用内建的，不自定义**——
# openpyxl 对自定义 numFmt 是按"位置"而非 id 索引的，随手写个 176 会让它读表时越界报错。
_BUILTIN_DATE_FMTS = {"14", "15", "16", "17", "22"}
_FALLBACK_DATE_FMT = "14"


def ensure_date_style(styles_xml: str) -> Tuple[str, str]:
    """
    保证工作簿里有一个"日期"单元格样式，返回 (新的 styles.xml, 该样式在 cellXfs 里的索引)。

    为什么要费这个劲：日期得写成真日期值（她要按收款时间排序筛选），
    而序列号没有日期格式就显示成 46211 这种数字。她表里那列本来有日期样式可借，
    空表/新列借不到——不能靠运气。追加不删除，"部件零丢失"的保证不受影响。
    """
    xfs_m = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", styles_xml, re.S)
    if not xfs_m:
        raise ValueError("styles.xml 里没有 cellXfs，工作簿结构异常")
    xf_list = re.findall(r"<xf\b[^>]*/?>", xfs_m.group(1))

    # 已有能当日期用的就直接复用
    date_ids = set(_BUILTIN_DATE_FMTS)
    for m in re.finditer(r'<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"', styles_xml):
        code = m.group(2).lower()
        if "y" in code and "d" in code and "h" not in code:
            date_ids.add(m.group(1))
    for i, xf in enumerate(xf_list):
        fid = re.search(r'numFmtId="(\d+)"', xf)
        if fid and fid.group(1) in date_ids:
            return styles_xml, str(i)

    new_idx = str(len(xf_list))
    new_xf = (
        f'<xf numFmtId="{_FALLBACK_DATE_FMT}" fontId="0" fillId="0" '
        f'borderId="0" applyNumberFormat="1"/>'
    )
    block = xfs_m.group(0)
    block = re.sub(r'count="\d+"', f'count="{len(xf_list) + 1}"', block, count=1)
    block = block.replace("</cellXfs>", new_xf + "</cellXfs>")
    out = styles_xml[: xfs_m.start()] + block + styles_xml[xfs_m.end() :]
    return out, new_idx


def _find_style_in_column(xml: str, col: str, skip_row: int) -> Optional[str]:
    """在同一列里找一个已有样式的单元格，借它的 s= 索引。"""
    for m in re.finditer(r'<c r="%s(\d+)"([^>]*?)(?:/>|>)' % col, xml):
        if int(m.group(1)) == skip_row:
            continue
        s = re.search(r'\bs="(\d+)"', m.group(2))
        if s:
            return s.group(1)
    return None


def _render_cell(ref: str, value, style: Optional[str]) -> str:
    """生成一个 <c> 元素。字符串走 inlineStr，避免动 sharedStrings（那会牵连全表）。"""
    s_attr = f' s="{style}"' if style else ""
    if value is None:
        return f'<c r="{ref}"{s_attr}/>'
    if isinstance(value, (dt.date, dt.datetime)):
        return f'<c r="{ref}"{s_attr}><v>{to_serial(value):g}</v></c>'
    if isinstance(value, bool):
        return f'<c r="{ref}"{s_attr} t="b"><v>{int(value)}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s_attr}><v>{value}</v></c>'
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def _patch_row(xml: str, row: int, cells: Dict[str, object], date_style: Optional[str] = None) -> str:
    """把一行里的若干单元格改掉；行内没有该单元格就按列序插进去。"""
    rm = re.search(_ROW_RE_TMPL.format(row=row), xml, re.S)
    if not rm:
        raise ValueError(f"sheet 里找不到第 {row} 行")
    row_xml = rm.group(0)
    new_row = row_xml
    for col, value in cells.items():
        ref = f"{col}{row}"
        style = None
        cm = re.search(_CELL_RE_TMPL.format(ref=re.escape(ref)), new_row, re.S)
        if cm:
            s = re.search(r'\bs="(\d+)"', cm.group("attrs") or "")
            style = s.group(1) if s else None
        if style is None:
            style = _find_style_in_column(xml, col, row)
        if style is None and isinstance(value, (dt.date, dt.datetime)):
            style = date_style  # 借不到就用刚追加的日期样式，别让日期显示成 46211
        cell_xml = _render_cell(ref, value, style)
        if cm:
            new_row = new_row[: cm.start()] + cell_xml + new_row[cm.end() :]
        else:
            # 插到列序正确的位置，Excel 对 <c> 的顺序是敏感的
            inserted = False
            for m in re.finditer(r'<c r="([A-Z]+)(\d+)"', new_row):
                if _col_index(m.group(1)) > _col_index(col):
                    new_row = new_row[: m.start()] + cell_xml + new_row[m.start() :]
                    inserted = True
                    break
            if not inserted:
                if new_row.endswith("/>"):  # 空行 <row .../>
                    new_row = new_row[:-2] + ">" + cell_xml + "</row>"
                else:
                    new_row = new_row[: new_row.rfind("</row>")] + cell_xml + "</row>"
    return xml[: rm.start()] + new_row + xml[rm.end() :]


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def patch_cells(
    src: Path,
    out: Path,
    sheet_name: str,
    edits: List[Tuple[int, int, object]],
) -> int:
    """
    edits = [(行号1基, 列号1基, 值)]。返回改动的格数。
    src 不动；out 是新文件，除目标 sheet 外**每个部件按原始字节复制**。
    """
    src, out = Path(src), Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as zin:
        target = sheet_path_for(zin, sheet_name)
        xml = zin.read(target).decode("utf-8", "replace")
        infos = zin.infolist()
        payload = {i.filename: zin.read(i.filename) for i in infos}
        by_row: Dict[int, Dict[str, object]] = {}
        for r, c, v in edits:
            by_row.setdefault(r, {})[col_letter(c)] = v

        # 日期值若在同列借不到日期样式，就往 styles.xml 追加一个（不删任何东西）
        date_style = None
        needs_date = [
            (r, col) for r, cells in by_row.items()
            for col, v in cells.items() if isinstance(v, (dt.date, dt.datetime))
        ]
        if needs_date and any(_find_style_in_column(xml, col, r) is None for r, col in needs_date):
            styles_name = "xl/styles.xml"
            if styles_name in payload:
                new_styles, date_style = ensure_date_style(
                    payload[styles_name].decode("utf-8", "replace")
                )
                payload[styles_name] = new_styles.encode("utf-8")

        for r, cells in sorted(by_row.items()):
            xml = _patch_row(xml, r, cells, date_style=date_style)
    payload[target] = xml.encode("utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in infos:
            zo.writestr(i, payload[i.filename])
    return sum(len(c) for c in by_row.values())


def parts_diff(a: Path, b: Path) -> List[str]:
    """列出 b 相对 a 少掉的部件——用来证明"没搞坏她的表"。"""
    def names(p):
        with zipfile.ZipFile(p) as z:
            return {n for n in z.namelist() if not n.endswith("/")}
    return sorted(names(a) - names(b))
