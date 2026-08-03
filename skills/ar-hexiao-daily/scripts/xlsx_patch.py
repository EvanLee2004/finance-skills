#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
无损写单元格：**只补丁目标 sheet 的那几个格，工作簿其余业务部件逐字节原样保留**。

为什么不用 openpyxl 保存：实测拿明妹的真表（`盈亏核算表2026全年`）用 openpyxl 载入再保存，
74 个部件变 59 —— **5 个 drawing、1 张内嵌图片、若干 rels 全部丢失**。
她一打开发现图没了，"你们把我的表搞坏了"，这个项目就别想推下去了。

做法（与 `xlsx` 技能的 unpack/pack 同思路，OOXML 外科手术）：
  1. 把 xlsx 当 zip 打开，只取出目标 sheet 的 XML
  2. 在 XML 里就地改那几个 `<c>` 单元格
  3. 重新打包：**除该 sheet 与可重建的 calcChain 外，每个部件都按原始字节写回**

只要公式或行坐标发生变化，旧 `calcChain.xml` 就可能继续指向被覆盖或已下移的单元格。
保留这种过期计算链会让 Excel 打开时进入“修复并删除公式”流程。因此写入时完整移除
calcChain 部件、关系和 Content Type，并令 Excel 在打开时全量重建计算链；工作表公式本身
和缓存值仍由本模块保留。

样式沿用：空格子没有样式时，从**同一列另一个有值的行**借样式索引，
这样日期/金额显示格式跟她表里其余行一致，不会出现"一列里有的显示日期有的显示数字"。
"""

from __future__ import annotations

import datetime as dt
import math
import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape, unescape

from openpyxl.formula.translate import Translator

# Excel 日期序列号的零点。1900 历法有个著名的假闰日（1900-02-29），
# 所以 1900-03-01 之后要 +1；用 1899-12-30 当基准正好把这个偏差吸收掉。
_EXCEL_EPOCH = dt.date(1899, 12, 30)

_CELL_RE_TMPL = r'<c r="{ref}"(?P<attrs>[^>]*?)(?:/>|>(?P<inner>.*?)</c>)'
# 自闭合元素必须作为独立分支优先匹配；若把自闭合与普通结束标签放在
# 同一个贪婪尾部分支里，匹配会吞掉斜杠并跨到下一个结束标签。
_CELL_ELEMENT_RE = re.compile(r'<c\b[^>]*/>|<c\b[^>]*>.*?</c>', re.S)
_ROW_ELEMENT_RE = re.compile(r'<row\b[^>]*/>|<row\b[^>]*>.*?</row>', re.S)
_FORMULA_ELEMENT_RE = re.compile(r'<f\b[^>]*/>|<f\b[^>]*>.*?</f>', re.S)
_ROW_RE_TMPL = (
    r'<row\b(?=[^>]*\br="{row}")[^>]*/>'
    r'|<row\b(?=[^>]*\br="{row}")[^>]*>.*?</row>'
)


@dataclass(frozen=True)
class FormulaValue:
    """需要同时保留 Excel 公式与缓存显示值的单元格。"""

    formula: str
    cached: float


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


def _formula_cache_text(value) -> str:
    """把公式缓存写成最短无损浮点文本，保留财务金额的分位精度。"""
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"公式缓存值必须是有限数字：{value!r}")
    # 不可使用 ``:g``：其默认只有 6 位有效数字，会把 15479.75 写成 15479.8。
    return repr(number)


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
    if isinstance(value, FormulaValue):
        formula = value.formula.lstrip("=")
        return (
            f'<c r="{ref}"{s_attr}><f>{escape(formula)}</f>'
            f'<v>{_formula_cache_text(value.cached)}</v></c>'
        )
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


def _translate_formula_between_cells(formula_xml: str, origin: str, target: str) -> str:
    """按 Excel 复制语义完整平移 A1 公式；绝对行列引用由 Translator 自动保持。"""
    formula = unescape(formula_xml, {"&quot;": '"', "&apos;": "'"})
    try:
        translated = Translator(f"={formula}", origin=origin).translate_formula(target)
    except Exception as exc:
        raise ValueError(
            f"公式平移失败：{origin}->{target}，formula={formula!r}"
        ) from exc
    return escape(translated.lstrip("="))


def _renumber_row_xml(row_xml: str, old_row: int, new_row: int) -> str:
    """把整行下移，并按每个单元格的原/新坐标完整平移其中的普通公式。"""

    def shift_cell(match) -> str:
        cell_xml = match.group(0)
        ref_match = re.search(r'\br="([A-Z]{1,3})(\d+)"', cell_xml)
        if not ref_match:
            return cell_xml
        col, row_text = ref_match.groups()
        if int(row_text) != old_row:
            return cell_xml
        origin = f"{col}{old_row}"
        target = f"{col}{new_row}"
        shifted = re.sub(
            rf'(\br="){re.escape(origin)}(")',
            rf"\g<1>{target}\2",
            cell_xml,
            count=1,
        )
        return re.sub(
            r'(<f\b[^>]*>)(.*?)(</f>)',
            lambda formula_match: formula_match.group(1)
            + _translate_formula_between_cells(
                formula_match.group(2), origin, target
            )
            + formula_match.group(3),
            shifted,
            flags=re.S,
        )

    out = _CELL_ELEMENT_RE.sub(shift_cell, row_xml)
    return re.sub(
        rf'(<row\b[^>]*\br="){old_row}(")',
        rf'\g<1>{new_row}\2',
        out,
        count=1,
    )


def _shift_a1_token(token: str, after_row: int, include_inserted: bool = False) -> str:
    """把本 sheet 的 A1 / A1:B2 / 多段 sqref 按“在 after_row 后插一行”平移。"""
    cell_re = re.compile(r'^(?P<col>\$?[A-Z]{1,3})(?P<abs>\$?)(?P<row>\d+)$')

    def one_cell(cell: str, *, is_end: bool = False) -> str:
        m = cell_re.match(cell)
        if not m:
            return cell
        row = int(m.group("row"))
        should_shift = row > after_row or (include_inserted and is_end and row == after_row)
        return f'{m.group("col")}{m.group("abs")}{row + 1}' if should_shift else cell

    parts = []
    for piece in token.split():
        if ":" in piece:
            left, right = piece.split(":", 1)
            parts.append(f"{one_cell(left)}:{one_cell(right, is_end=True)}")
        else:
            parts.append(one_cell(piece))
    return " ".join(parts)


def _shift_sheet_ranges(xml: str, after_row: int) -> str:
    """同步 sheet 内常见范围，避免插行后筛选、验证、合并范围仍停在旧行号。"""
    tags = (
        "dimension", "autoFilter", "mergeCell", "conditionalFormatting",
        "dataValidation", "ignoredError", "hyperlink",
    )
    for tag in tags:
        xml = re.sub(
            rf'(<{tag}\b[^>]*\b(?:ref|sqref)=")([^"]+)(")',
            lambda m: m.group(1)
            + _shift_a1_token(m.group(2), after_row, include_inserted=True)
            + m.group(3),
            xml,
        )
    return xml


def _shift_shared_formula_refs(xml: str, after_row: int) -> str:
    """插行时同步 shared formula 主公式的 ref；跨过插入点的组扩展一行。"""
    def repl(m):
        tag = m.group(0)
        if not re.search(r'\bt="shared"', tag) or not re.search(r'\bref="', tag):
            return tag
        return re.sub(
            r'(\bref=")([^"]+)(")',
            lambda x: x.group(1)
            + _shift_shared_formula_ref(x.group(2), after_row)
            + x.group(3),
            tag,
            count=1,
        )

    return re.sub(r"<f\b[^>]*>", repl, xml)


def _shift_shared_formula_ref(ref: str, after_row: int) -> str:
    """扩展/平移 shared master ref；单格 master 被复制时也必须覆盖新 follower。"""
    if ":" not in ref and " " not in ref:
        match = re.fullmatch(r'(?P<col>\$?[A-Z]{1,3})(?P<abs>\$?)(?P<row>\d+)', ref)
        if match and int(match.group("row")) == after_row:
            end = f'{match.group("col")}{match.group("abs")}{after_row + 1}'
            return f"{ref}:{end}"
    return _shift_a1_token(ref, after_row, include_inserted=True)


def _cell_in_a1_range(cell: str, ref: str) -> bool:
    def coordinate(value: str):
        match = re.fullmatch(r'\$?([A-Z]{1,3})\$?(\d+)', value)
        if not match:
            return None
        return _col_index(match.group(1)), int(match.group(2))

    if " " in ref:
        return any(_cell_in_a1_range(cell, piece) for piece in ref.split())
    left, right = (ref.split(":", 1) + [ref])[:2] if ":" in ref else (ref, ref)
    point = coordinate(cell)
    start = coordinate(left)
    end = coordinate(right)
    if not point or not start or not end:
        return False
    return (
        min(start[0], end[0]) <= point[0] <= max(start[0], end[0])
        and min(start[1], end[1]) <= point[1] <= max(start[1], end[1])
    )


def _validate_shared_formula_integrity(xml: str) -> None:
    """拒绝 master 缺失/重复或 follower 落在 ref 之外的共享公式结构。"""
    groups: Dict[str, Dict[str, object]] = {}
    for cell_match in _CELL_ELEMENT_RE.finditer(xml):
        cell_xml = cell_match.group(0)
        cell_ref_match = re.search(r'\br="([A-Z]{1,3}\d+)"', cell_xml)
        if not cell_ref_match:
            continue
        formula_match = _FORMULA_ELEMENT_RE.search(cell_xml)
        if not formula_match:
            continue
        formula_xml = formula_match.group(0)
        tag_match = re.match(r'<f\b[^>]*?/?>', formula_xml)
        if not tag_match or not re.search(r'\bt="shared"', tag_match.group(0)):
            continue
        si_match = re.search(r'\bsi="([^"]+)"', tag_match.group(0))
        if not si_match:
            raise ValueError(f"共享公式缺少 si：{cell_ref_match.group(1)}")
        group = groups.setdefault(si_match.group(1), {"masters": [], "followers": []})
        ref_match = re.search(r'\bref="([^"]+)"', tag_match.group(0))
        if ref_match:
            group["masters"].append((cell_ref_match.group(1), ref_match.group(1)))
        else:
            group["followers"].append(cell_ref_match.group(1))

    problems: List[str] = []
    for si, group in groups.items():
        masters = group["masters"]
        followers = group["followers"]
        if len(masters) != 1:
            problems.append(f"si={si} master_count={len(masters)}")
            continue
        master_cell, master_ref = masters[0]
        if not _cell_in_a1_range(master_cell, master_ref):
            problems.append(f"si={si} master={master_cell} outside ref={master_ref}")
        for follower in followers:
            if not _cell_in_a1_range(follower, master_ref):
                problems.append(f"si={si} follower={follower} outside ref={master_ref}")
    if problems:
        raise ValueError("共享公式结构非法：" + "; ".join(problems[:10]))


def _drop_calc_chain(payload: Dict[str, bytes]) -> None:
    """公式/坐标变更后移除可重建的旧计算链，避免 Excel 打开时修复并删公式。"""
    payload.pop("xl/calcChain.xml", None)

    rels_name = "xl/_rels/workbook.xml.rels"
    if rels_name in payload:
        rels_xml = payload[rels_name].decode("utf-8", "replace")
        rels_xml = re.sub(
            r'<Relationship\b(?=[^>]*\bType="[^"]*/calcChain")(?=[^>]*\bTarget="calcChain\.xml")[^>]*/>',
            "",
            rels_xml,
        )
        payload[rels_name] = rels_xml.encode("utf-8")

    content_types_name = "[Content_Types].xml"
    if content_types_name in payload:
        content_xml = payload[content_types_name].decode("utf-8", "replace")
        content_xml = re.sub(
            r'<Override\b(?=[^>]*\bPartName="/xl/calcChain\.xml")[^>]*/>',
            "",
            content_xml,
        )
        payload[content_types_name] = content_xml.encode("utf-8")


def _demote_copied_shared_masters(row_xml: str) -> str:
    """复制到新行的 shared master 改成同组 follower，避免同一 si 出现两个主公式。"""
    def repl(m):
        formula_xml = m.group(0)
        tag_match = re.match(r'<f\b[^>]*>', formula_xml)
        if not tag_match:
            return formula_xml
        tag = tag_match.group(0)
        if not re.search(r'\bt="shared"', tag) or not re.search(r'\bref="', tag):
            return formula_xml
        si = re.search(r'\bsi="([^"]+)"', tag)
        if not si:
            return formula_xml
        return f'<f t="shared" si="{si.group(1)}"/>'

    return _FORMULA_ELEMENT_RE.sub(repl, row_xml)


def _insert_row_copy(xml: str, source_row: int) -> str:
    """在 source_row 正下方复制一行；共享公式保留同一 si，并扩展主公式 ref。"""
    source_match = re.search(_ROW_RE_TMPL.format(row=source_row), xml, re.S)
    if not source_match:
        raise ValueError(f"sheet 里找不到待拆分的第 {source_row} 行")
    def shift_row(m):
        row_xml = m.group(0)
        rm = re.search(r'<row\b[^>]*\br="(\d+)"', row_xml)
        if not rm:
            return row_xml
        old = int(rm.group(1))
        return _renumber_row_xml(row_xml, old, old + 1) if old > source_row else row_xml

    xml = _ROW_ELEMENT_RE.sub(shift_row, xml)
    source_match = re.search(_ROW_RE_TMPL.format(row=source_row), xml, re.S)
    source_xml = source_match.group(0)
    inserted = _renumber_row_xml(source_xml, source_row, source_row + 1)
    inserted = _demote_copied_shared_masters(inserted)
    xml = xml[: source_match.end()] + inserted + xml[source_match.end() :]
    xml = _shift_shared_formula_refs(xml, source_row)
    return _shift_sheet_ranges(xml, source_row)


def patch_cells(
    src: Path,
    out: Path,
    sheet_name: str,
    edits: List[Tuple[int, int, object]],
    insertions: Optional[List[Tuple[int, Dict[int, object]]]] = None,
) -> int:
    """
    edits = [(原始行号1基, 列号1基, 值)]。
    insertions = [(原始源行号, {列号1基: 新行覆盖值})]，表示在源行下复制一行。
    返回改动格数 + 插入行数。
    src 不动；out 是新文件，除目标 sheet 外**每个部件按原始字节复制**。
    """
    src, out = Path(src), Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as zin:
        target = sheet_path_for(zin, sheet_name)
        xml = zin.read(target).decode("utf-8", "replace")
        infos = zin.infolist()
        payload = {i.filename: zin.read(i.filename) for i in infos}
        insertion_specs = list(insertions or [])
        source_rows = [int(x[0]) for x in insertion_specs]
        if len(source_rows) != len(set(source_rows)):
            raise ValueError("同一原始行不能在一批计划里拆分两次")
        for source_row in sorted(source_rows, reverse=True):
            xml = _insert_row_copy(xml, source_row)

        def final_original_row(row: int) -> int:
            return int(row) + sum(1 for source in source_rows if source < int(row))

        by_row: Dict[int, Dict[str, object]] = {}
        for r, c, v in edits:
            by_row.setdefault(final_original_row(int(r)), {})[col_letter(c)] = v
        for source_row, overrides in insertion_specs:
            inserted_row = final_original_row(int(source_row)) + 1
            for c, v in overrides.items():
                by_row.setdefault(inserted_row, {})[col_letter(int(c))] = v

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
        _validate_shared_formula_integrity(xml)
        _drop_calc_chain(payload)
        workbook_name = "xl/workbook.xml"
        if workbook_name in payload:
            wb_xml = payload[workbook_name].decode("utf-8", "replace")
            attrs = 'calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1" calcId="0"'
            if re.search(r"<calcPr\b[^>]*/>", wb_xml):
                wb_xml = re.sub(r"<calcPr\b[^>]*/>", f"<calcPr {attrs}/>", wb_xml, count=1)
            elif re.search(r"<calcPr\b[^>]*>.*?</calcPr>", wb_xml, flags=re.S):
                wb_xml = re.sub(
                    r"<calcPr\b[^>]*>.*?</calcPr>",
                    f"<calcPr {attrs}/>",
                    wb_xml,
                    count=1,
                    flags=re.S,
                )
            else:
                wb_xml = wb_xml.replace("</workbook>", f"<calcPr {attrs}/></workbook>")
            payload[workbook_name] = wb_xml.encode("utf-8")
    payload[target] = xml.encode("utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in infos:
            if i.filename in payload:
                zo.writestr(i, payload[i.filename])
    return sum(len(c) for c in by_row.values()) + len(insertion_specs)


def parts_diff(a: Path, b: Path) -> List[str]:
    """列出 b 相对 a 意外少掉的部件；可重建的 calcChain 允许被主动失效。"""
    def names(p):
        with zipfile.ZipFile(p) as z:
            return {n for n in z.namelist() if not n.endswith("/")}
    return sorted((names(a) - names(b)) - {"xl/calcChain.xml"})
