"""Read-only provenance for decisions before and after ledger-dependent expansion."""
from __future__ import annotations

import hashlib
import json

import common


def payment_source_lineage(payment: dict, so: str) -> dict:
    # Capture before BOTH parent allocation and SOD expansion. A settled SO may
    # receive zero on reclassification; that is an outcome, not a new event.
    day = common.norm_date(payment.get("hexiao_date"))
    facts = {key: payment.get(key) or "" for key in ("ar", "currency")}
    facts.update(
        so=so,
        reconciliation_date=day.isoformat() if day else "",
        writeoff_sequence_key=(payment.get("_writeoff_sequence_key_by_so") or {}).get(so),
        parent_amount_orig=common.to_number(payment.get("amount_orig")),
        parent_amount_local=common.to_number(payment.get("amount_local")),
        order_writeoff_orig=common.to_number((payment.get("writeoffs") or {}).get(so)),
        order_writeoff_local=common.to_number((payment.get("writeoffs_local") or {}).get(so)),
    )
    encoded = json.dumps(facts, ensure_ascii=False, sort_keys=True)
    return {"schema_version": "ar-source-lineage-v1", "source_id": hashlib.sha256(encoded.encode("utf-8")).hexdigest(),
            "source": facts}


def indexed_decisions(payload: dict) -> dict[str, dict]:
    result = {}
    for bucket in ("auto", "hold", "exception"):
        for row in payload.get(bucket) or []:
            key = str(row.get("case_id") or "")
            if not key or key in result:
                raise ValueError("判定记录缺少唯一 case_id，不能追溯首次与最终结果")
            result[key] = {**row, "bucket": bucket}
    return result


def match_final_records(initial: dict[str, dict], final: dict[str, dict]) -> dict[str, list[str]]:
    """Map explicit source groups; never guess relationships from SO or amounts."""
    def groups(records):
        grouped = {}
        for key, row in records.items():
            lineage = row.get("source_lineage") or {}
            source_id = lineage.get("source_id")
            if not source_id or lineage.get("schema_version") != "ar-source-lineage-v1":
                raise ValueError("判定缺少展开前来源标识，不能猜测拆分或合并对应关系")
            encoded = json.dumps(lineage.get("source"), ensure_ascii=False, sort_keys=True)
            if hashlib.sha256(encoded.encode("utf-8")).hexdigest() != source_id:
                raise ValueError("判定来源标识与来源字段不一致")
            source = lineage.get("source") or {}
            if any((row.get(field) or "") != source.get(field) for field in ("ar", "so")):
                raise ValueError("判定 AR/SO 与展开前来源不一致")
            grouped.setdefault(source_id, []).append(key)
        return grouped

    before_groups, after_groups = groups(initial), groups(final)
    if before_groups.keys() != after_groups.keys():
        raise ValueError("首次与写后来源集合变化，存在未映射的来源事件，不能发布")
    links = {}
    for source_id, before_keys in before_groups.items():
        after_keys = after_groups[source_id]
        for key in before_keys:
            if len(before_keys) == 1 or len(after_keys) == 1:
                links[key] = list(after_keys)
            elif key in after_keys:
                links[key] = [key]
            else:
                raise ValueError("同一来源的多个 SOD 前后对应不唯一，需核对分配链，不能按 SO 合并")
        if {value for key in before_keys for value in links[key]} != set(after_keys):
            raise ValueError("写后存在没有首次记录对应的 SOD，不能漏报新增记录")
    return links


def checked_records(plan: dict, expected: set[str]) -> dict[str, dict[str, dict]]:
    result = {}
    seen = set()
    for bucket in ("write", "skip", "conflict"):
        result[bucket] = {}
        for row in plan.get(bucket) or []:
            key = str(row.get("case_id") or "")
            if not key or key in seen or key not in expected:
                raise ValueError("校验计划包含重复、缺失或无法对应自动判定的案例 ID")
            seen.add(key)
            result[bucket][key] = row
    if seen != expected:
        raise ValueError("自动判定未全部进入写入、跳过或冲突清单，不能生成最终报告")
    return result
