from __future__ import annotations

from typing import Any, Literal

from openpyxl.formula.translate import Translator


FormulaRelation = Literal[
    "identical",
    "coordinate_shift_equivalent",
    "logic_difference",
    "formula_value_difference",
    "value_difference",
]


def is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def translate_formula(formula: str, origin: str, target: str) -> str:
    """按 Excel 复制语义把公式从 origin 平移到 target。"""
    if not is_formula(formula):
        raise ValueError(f"不是 Excel 公式：{formula!r}")
    return Translator(formula, origin=origin).translate_formula(target)


def classify_formula_relation(
    left: Any,
    right: Any,
    *,
    left_coordinate: str,
    right_coordinate: str,
) -> FormulaRelation:
    """
    比较两个已配对单元格的公式关系。

    不直接比较公式文本中的行号。先把 left 从其真实坐标按 Excel 复制语义
    平移到 right 的真实坐标，再与 right 比较。这样能区分：
    - 插行造成的正常坐标平移；
    - 公式逻辑真实改变；
    - 公式与常量之间的表示差异。
    """
    left_is_formula = is_formula(left)
    right_is_formula = is_formula(right)
    if not left_is_formula and not right_is_formula:
        return "identical" if left == right else "value_difference"
    if left_is_formula != right_is_formula:
        return "formula_value_difference"
    if left_coordinate == right_coordinate and left == right:
        return "identical"
    try:
        translated = translate_formula(left, left_coordinate, right_coordinate)
    except Exception:
        return "logic_difference"
    if translated == right:
        return "coordinate_shift_equivalent"
    return "logic_difference"
