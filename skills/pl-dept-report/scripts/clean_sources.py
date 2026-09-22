#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把引出洗成同一套：同文件补期间，再只留本次月份。对不上的记进运行报告，不填进本期。"""
from __future__ import annotations

from pathlib import Path

from common import prev_period


def inherit_file_periods(items: list[dict]) -> list[dict]:
    """同一文件里有且只有一个会计期间时，没写期间的 sheet 跟随它。"""
    copied = [dict(item) for item in items]
    groups: dict[str, list[dict]] = {}
    for item in copied:
        groups.setdefault(str(item.get("path") or ""), []).append(item)
    for group in groups.values():
        periods = {item.get("period") for item in group if item.get("period")}
        if len(periods) != 1:
            continue
        only = next(iter(periods))
        for item in group:
            if item.get("period"):
                continue
            if item.get("kind") not in {"account", "assist", "profit"}:
                continue
            item["period"] = only
            item["period_from"] = "同文件"
    return copied


def bind_sources(items: list[dict], period: str) -> tuple[list[dict], list[str]]:
    """科目余额/核算项目只留本期。利润表留本期和上月。其余种类原样通过。"""
    items = inherit_file_periods(items)
    prev = prev_period(period)
    kept: list[dict] = []
    notes: list[str] = []
    undated: list[str] = []
    foreign: list[str] = []
    inherited: list[str] = []
    for item in items:
        kind = item.get("kind")
        if kind not in {"account", "assist", "profit"}:
            kept.append(item)
            continue
        name = Path(str(item.get("path") or "")).name
        sheet = str(item.get("sheet") or "")
        label = f"{name}/{sheet}" if sheet else name
        who = item.get("entity") or name
        per = item.get("period")
        if item.get("period_from") == "同文件" and name not in inherited:
            inherited.append(name)
        if kind in {"account", "assist"}:
            if per == period:
                kept.append(item)
            elif not per:
                undated.append(label)
            else:
                notes.append(f"过期{kind}跳过={who}:{per}")
                foreign.append(f"{label}:{per}")
            continue
        if per in {period, prev}:
            kept.append(item)
        elif not per:
            undated.append(label)
        else:
            notes.append(f"过期profit跳过={who}:{per}")
            foreign.append(f"{label}:{per}")
    if inherited:
        notes.append("期间随同文件=" + "、".join(dict.fromkeys(inherited)))
    if undated:
        notes.append("期间未写明=" + "、".join(dict.fromkeys(undated)))
    if foreign:
        notes.append("期间不符=" + "、".join(dict.fromkeys(foreign)))
    return kept, notes
