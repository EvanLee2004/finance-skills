#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""金额比较的统一边界。

技术容差用于来源匹配、金额守恒、写前/写后校验；业务容差只用于明确的
结清尾差和离线业务差异判断。禁止用业务容差放宽技术校验。
"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any


TECHNICAL_EPSILON = Decimal("0.005")
CENT_TOLERANCE = Decimal("0.01")
BUSINESS_SETTLEMENT_TOLERANCE = Decimal("1.00")


def as_decimal(value: Any) -> Decimal:
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError(f"金额不是有效数字：{value!r}") from exc


def difference(left: Any, right: Any) -> Decimal:
    return as_decimal(left) - as_decimal(right)


def technical_equal(left: Any, right: Any) -> bool:
    return abs(difference(left, right)) <= TECHNICAL_EPSILON


def cent_equal(left: Any, right: Any) -> bool:
    return abs(difference(left, right)) <= CENT_TOLERANCE


def within_business_tolerance(left: Any, right: Any) -> bool:
    return abs(difference(left, right)) <= BUSINESS_SETTLEMENT_TOLERANCE


def business_comparison(left: Any, right: Any) -> dict:
    delta = difference(left, right)
    return {
        "delta": delta,
        "technical_equal": abs(delta) <= TECHNICAL_EPSILON,
        "business_equal": abs(delta) <= BUSINESS_SETTLEMENT_TOLERANCE,
        "business_tolerance": BUSINESS_SETTLEMENT_TOLERANCE,
    }
