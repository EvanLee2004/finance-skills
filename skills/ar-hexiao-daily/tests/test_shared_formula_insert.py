# -*- coding: utf-8 -*-
"""共享公式行插入：保留 si、扩展 ref，且所有 follower 必须落在 master ref 内。"""
import re
from xml.sax.saxutils import unescape

from openpyxl.formula.translate import Translator
import pytest
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


def test_insert_above_translates_current_and_next_row_references():
    xml = _sheet([
        '<row r="1"><c r="M1"><v>0</v></c></row>',
        '<row r="2"><c r="M2"><f>N2+N3</f><v>30</v></c>'
        '<c r="N2"><v>10</v></c></row>',
        '<row r="3"><c r="N3"><v>20</v></c></row>',
    ])
    got = X._insert_row_copy(xml, 1)
    assert '<c r="M3"><f>N3+N4</f><v>30</v></c>' in got
    assert "N3+N3" not in got


def test_row_translation_preserves_absolute_refs_and_moves_relative_refs():
    row = (
        '<row r="7"><c r="M7">'
        '<f>N7+N8+$N$2+N$3+$N7+SUM(N7:N9)</f><v>30</v>'
        '</c></row>'
    )
    got = X._renumber_row_xml(row, 7, 8)
    assert 'r="8"' in got
    assert 'r="M8"' in got
    assert (
        "<f>N8+N9+$N$2+N$3+$N8+SUM(N8:N10)</f>"
        in got
    )


def test_self_closing_cell_before_shared_master_does_not_swallow_master():
    xml = _sheet([
        '<row r="2742"><c r="A2742"><v>1</v></c></row>',
        '<row r="3109"><c r="P3109" s="1"/>'
        '<c r="V3109"><f t="shared" ref="V3109:V3180" si="2417">'
        'IF(Q3109=&quot;否&quot;,W3109,0)</f><v>0</v></c></row>',
        '<row r="3159"><c r="V3159"><f t="shared" si="2417"/>'
        '<v>0</v></c></row>',
    ])

    got = X._insert_row_copy(xml, 2742)

    assert '<c r="P3110" s="1"/>' in got
    master = re.search(
        r'<c r="V3110"><f t="shared" ref="V3110:V3181" si="2417">'
        r'(.*?)</f>',
        got,
    )
    assert master is not None
    assert unescape(master.group(1)) == 'IF(Q3110="否",W3110,0)'
    assert '<c r="V3160"><f t="shared" si="2417"/>' in got
    expanded = Translator(
        "=" + unescape(master.group(1)),
        origin="V3110",
    ).translate_formula("V3160")
    assert expanded == '=IF(Q3160="否",W3160,0)'
    assert 'IF(Q3161=&quot;否&quot;,W3161,0)' not in got


def test_self_closing_row_does_not_swallow_following_formula_row():
    xml = _sheet([
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        '<row r="2"/>',
        '<row r="3"><c r="M3"><f>N3+N4</f><v>30</v></c></row>',
    ])

    got = X._insert_row_copy(xml, 1)

    assert '<row r="3"/>' in got
    assert '<row r="4"><c r="M4"><f>N4+N5</f><v>30</v></c></row>' in got


def test_copied_shared_follower_does_not_swallow_later_master_demotion():
    xml = _sheet([
        '<row r="1">'
        '<c r="V1"><f t="shared" si="1"/><v>0</v></c>'
        '<c r="AL1"><f t="shared" ref="AL1" si="2">AI1-AK1</f><v>0</v></c>'
        '<c r="AM1"><f t="shared" ref="AM1" si="3">AL1+1</f><v>0</v></c>'
        '</row>',
        '<row r="2"><c r="A2"><v>2</v></c></row>',
    ])

    got = X._insert_row_copy(xml, 1)

    assert '<c r="V2"><f t="shared" si="1"/><v>0</v></c>' in got
    assert '<c r="AL2"><f t="shared" si="2"/><v>0</v></c>' in got
    assert '<c r="AM2"><f t="shared" si="3"/><v>0</v></c>' in got
    assert 'ref="AL1:AL2"' in got
    assert 'ref="AM1:AM2"' in got
    assert 'r="AL2"><f t="shared" ref=' not in got
    assert 'r="AM2"><f t="shared" ref=' not in got


def test_single_cell_shared_master_expands_to_copied_follower():
    xml = _sheet([
        '<row r="2742">'
        '<c r="AL2742"><f t="shared" ref="AL2742" si="2076">'
        'AI2742-AK2742</f><v>0</v></c>'
        '<c r="AM2742"><f t="shared" ref="AM2742" si="2077">'
        'AL2742+1</f><v>0</v></c>'
        '</row>',
        '<row r="2743"><c r="A2743"><v>1</v></c></row>',
    ])

    got = X._insert_row_copy(xml, 2742)

    assert '<f t="shared" ref="AL2742:AL2743" si="2076">' in got
    assert '<f t="shared" ref="AM2742:AM2743" si="2077">' in got
    assert '<c r="AL2743"><f t="shared" si="2076"/><v>0</v></c>' in got
    assert '<c r="AM2743"><f t="shared" si="2077"/><v>0</v></c>' in got
    X._validate_shared_formula_integrity(got)


def test_shared_formula_validator_rejects_follower_outside_master_ref():
    xml = _sheet([
        '<row r="1"><c r="AL1"><f t="shared" ref="AL1" si="9">'
        'AI1-AK1</f><v>0</v></c></row>',
        '<row r="2"><c r="AL2"><f t="shared" si="9"/><v>0</v></c></row>',
    ])

    with pytest.raises(ValueError, match="follower=AL2 outside ref=AL1"):
        X._validate_shared_formula_integrity(xml)
