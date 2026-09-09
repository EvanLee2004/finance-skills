#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""名称对档案：去有限公司、写两遍、唯一包含。2–5 候选不自动猜编码。"""
from __future__ import annotations

from dataclasses import dataclass, field

ORG_MARKERS = ("公司", "中心", "大学", "局", "学院", "医院", "银行", "集团", "厂", "所", "委员会", "办公室")
SUFFIXES = ("股份有限公司", "有限责任公司", "有限公司")
PERSON_CODE = "0582"
PERSON_NAME = "个人"
POLICE_CODE = "0386"
POLICE_NAME = "公安部"


def norm_name(name: str) -> str:
    s = (name or "").strip().replace("(", "（").replace(")", "）")
    return "".join(s.split())


def peel(name: str) -> str:
    n = norm_name(name)
    for suf in SUFFIXES:
        if n.endswith(suf) and len(n) > len(suf) + 1:
            return n[: -len(suf)]
    return n


def is_person_heading(name: str) -> bool:
    raw = (name or "").strip()
    if not raw or any(m in raw for m in ORG_MARKERS):
        return False
    n = norm_name(raw)
    return 2 <= len(n) <= 4


def is_haidian_police(name: str) -> bool:
    raw = name or ""
    return "公安" in raw and "海淀" in raw


def maps_to_police_ministry(name: str, applicant: str = "") -> bool:
    if "公安" not in (name or ""):
        return False
    if is_haidian_police(name) and (applicant or "").strip() == "陈霞":
        return False
    return True


def hang_employee(name: str, mapping: dict[str, str]) -> str:
    raw = (name or "").strip()
    if not raw:
        return raw
    if raw in mapping:
        return str(mapping[raw]).strip()
    folded = norm_name(raw)
    if folded in mapping:
        return str(mapping[folded]).strip()
    for key, dest in (mapping or {}).items():
        if str(key).startswith("_") or not dest:
            continue
        if norm_name(str(key)) == folded:
            return str(dest).strip()
    return raw


@dataclass
class NameMatch:
    hit: tuple[str, str] | None = None
    status: str = "none"  # ok / none / many
    candidates: list[tuple[str, str]] = field(default_factory=list)


def _unique(hits: list[tuple[str, str]]) -> list[tuple[str, str]]:
    out = []
    for item in hits:
        if item not in out:
            out.append(item)
    return out


def match_records(query: str, records: list[tuple[str, str]]) -> NameMatch:
    qn = norm_name(query)
    qp = peel(query)
    if not qn:
        return NameMatch(status="none")
    exact = _unique([(c, n) for c, n in records if norm_name(n) == qn])
    if len(exact) == 1:
        return NameMatch(hit=exact[0], status="ok", candidates=exact)
    if len(exact) > 1:
        return NameMatch(status="many", candidates=exact)
    doubled = _unique(
        [(c, n) for c, n in records if norm_name(n) in (qn + qn, qp + qp) and qn]
    )
    if len(doubled) == 1:
        return NameMatch(hit=doubled[0], status="ok", candidates=doubled)
    peeled = _unique([(c, n) for c, n in records if peel(n) == qp and qp])
    if len(peeled) == 1:
        return NameMatch(hit=peeled[0], status="ok", candidates=peeled)
    if len(peeled) > 1:
        return NameMatch(status="many", candidates=peeled)
    contain = []
    if len(qp) >= 4 or len(qn) >= 4:
        needle = qp if len(qp) >= 4 else qn
        for code, name in records:
            nn, npn = norm_name(name), peel(name)
            if needle in nn or needle in npn or nn in qn or npn in qp:
                contain.append((code, name))
    contain = _unique(contain)
    if len(contain) == 1:
        return NameMatch(hit=contain[0], status="ok", candidates=contain)
    if len(contain) >= 2:
        return NameMatch(status="many", candidates=contain[:8])
    return NameMatch(status="none")


def by_code(records: list[tuple[str, str]], code: str) -> tuple[str, str] | None:
    code = str(code or "").strip()
    hits = _unique([(c, n) for c, n in records if str(c) == code])
    if len(hits) == 1:
        return hits[0]
    return None
