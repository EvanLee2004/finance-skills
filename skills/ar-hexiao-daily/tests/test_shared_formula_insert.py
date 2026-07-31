# -*- coding: utf-8 -*-
"""共享公式行插入：保留 si、扩展 ref，且复制 master 时只保留一个主公式。"""
import re

import xlsx_patch as X


def _sheet(rows):
    return (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1:A3"/><sheetData>'
        + "".join(rows)
        + "</sheetData></worksheet>"
    )


def test_insert_shared_formula_follower_expands_group():
    xml = _sheet([
        '<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="7">B1*2</f><v>2</v></c></row>',
        '<row r="2"><c r="A2"><f t="shared" si="7"/><v>4</v></c></row>',
        '<row r="3"><c r="A3"><f t="shared" si="7"/><v>6</v></c></row>',
    ])
    got = X._insert_row_copy(xml, 2)
    assert 'ref="A1:A4"' in got
    assert len(re.findall(r'<f\b[^>]*si="7"', got)) == 4
    assert len(re.findall(r'<f\b[^>]*ref="A1:A4"', got)) == 1
    assert '<c r="A3"><f t="shared" si="7"/>' in got
    assert '<c r="A4"><f t="shared" si="7"/>' in got


def test_insert_shared_formula_master_demotes_copied_master():
    xml = _sheet([
        '<row r="1"><c r="A1"><f ref="A1:A3" si="9" t="shared">B1*2</f><v>2</v></c></row>',
        '<row r="2"><c r="A2"><f t="shared" si="9"/><v>4</v></c></row>',
        '<row r="3"><c r="A3"><f t="shared" si="9"/><v>6</v></c></row>',
    ])
    got = X._insert_row_copy(xml, 1)
    assert 'ref="A1:A4"' in got
    assert len(re.findall(r'<f\b[^>]*ref="A1:A4"', got)) == 1
    assert '<c r="A2"><f t="shared" si="9"/>' in got
    assert len(re.findall(r'<f\b[^>]*si="9"', got)) == 4

