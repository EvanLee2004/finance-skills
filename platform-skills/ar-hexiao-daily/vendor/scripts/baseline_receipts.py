"""Preserve historical receivables while recording delivery-based receipts.

This module builds and checks explicit plans; it never modifies a workbook.
Receipt identities are published with the existing auxiliary-ledger bundle.
"""
from __future__ import annotations

import copy
import json
from decimal import Decimal, InvalidOperation

import common
import amount_policy

OPERATION = "preserve_baseline_receipts"
SETTLEMENT_CENTS = int(amount_policy.BUSINESS_SETTLEMENT_TOLERANCE * 100)
FIELDS = ("SO", "SOD", "应收金额", "计提", "回款明细", "是否结账", "收款时间", "收款方式", "差异")
NUMERIC = {"应收金额", "计提", "回款明细", "差异"}


def cents(value):
    if value is None or value == "":
        return None
    try:
        number = Decimal(str(value))
        if not number.is_finite():
            raise ValueError("金额必须是有限数值")
        return int((number * 100).quantize(Decimal("1")))
    except (InvalidOperation, TypeError) as exc:
        raise ValueError("金额格式无效") from exc


def normalized(row: dict) -> dict:
    result = {}
    for key in FIELDS:
        value = row.get(key)
        if key in NUMERIC:
            number = cents(value)
            result[key] = number / 100 if number is not None else None
        elif key == "收款时间":
            day = common.norm_date(value)
            result[key] = day.isoformat() if day else None
        else:
            result[key] = str(value or "").strip()
    return result


def ledger_rows(ledger, so: str, sod: str) -> dict:
    aliases = dict(zip(FIELDS, ("so", "sod", "yingshou", "jiti", "huikuan", "jiezhang",
                               "shoukuan_time", "shoukuan_way", "chayi")))
    return {str(ref): normalized({key: ledger.row_snapshot[ref].get(alias)
                                 for key, alias in aliases.items()})
            for ref in ledger.business_rows(so, sod)}


def group_key(so, sod):
    return json.dumps([so, sod], ensure_ascii=False, separators=(",", ":"))


def event_key(rec: dict) -> str:
    sequence = rec.get("writeoff_sequence_key") or []
    record_id = str(sequence[1]).strip() if len(sequence) >= 2 else ""
    # No-detail allocations already have a durable parent/SOD identity.
    if not record_id and rec.get("parent_allocation_audit"):
        record_id = "parent-allocation"
    if not all((rec.get("ar"), rec.get("so"), rec.get("sod"), record_id)):
        return ""
    return json.dumps([rec["ar"], rec["so"], rec["sod"], record_id], ensure_ascii=False)


def signature(row):
    return tuple(row.get(key) for key in ("回款明细", "收款时间", "收款方式"))


def candidate(rec: dict, result: dict, ledger) -> dict | None:
    """Select the business mode before any current-payment or SO skip."""
    if ledger is None or not rec.get("so") or not rec.get("sod"):
        return None
    before = ledger_rows(ledger, rec["so"], rec["sod"])
    if not before:
        return None
    values = list(before.values())
    nonblank = [(ref, row) for ref, row in before.items() if row["应收金额"] is not None]
    baseline = sum(cents(row["应收金额"]) or 0 for row in values)
    delivery = cents(rec.get("deliver_local"))
    amount = cents(rec.get("amount_local"))
    received = sum(cents(row["回款明细"]) or 0 for row in values)
    cumulative = cents(rec.get("cumulative_received_local"))
    key = group_key(rec["so"], rec["sod"])
    journal = getattr(ledger, "baseline_receipt_state", {}).get(key) or {}
    blank_receipts = [row for row in values if row["应收金额"] is None and (cents(row["回款明细"]) or 0) > 0]
    preserved = (bool(rec.get("receivable_group_scope")) and delivery != baseline) or bool(journal and not journal.get("scope_only")) or bool(len(nonblank) == 1 and blank_receipts)
    prospective = cumulative if cumulative is not None else received + (amount or 0)
    # D == B0 uses ordinary splitting unless historical structure pins this mode.
    changed = delivery is not None and delivery != baseline
    batch_total = cents(rec.get("_baseline_batch_cumulative"))
    first_total = cents(rec.get("_baseline_batch_first_cumulative"))
    select = preserved or (changed and delivery > baseline and
                          max(prospective, batch_total or 0) > baseline and
                          min(prospective, first_total if first_total is not None else prospective) < delivery)
    if not select:
        return None
    receipt = {
        "回款明细": amount / 100 if amount is not None else None,
        "收款时间": None, "收款方式": "",
    }
    paid_day, posting_day = common.norm_date(rec.get("shoukuan_date")), common.norm_date(rec.get("hexiao_date"))
    day = common.receipt_time(paid_day, posting_day)
    receipt["收款时间"] = day.isoformat() if day else None
    receipt["收款方式"] = common.pay_way(rec.get("status") or "", paid_day, posting_day)
    identity = event_key(rec)
    audit = {"mode": OPERATION, "baseline_receivable": baseline / 100,
             "latest_delivery": delivery / 100 if delivery is not None else None,
             "current_received": receipt["回款明细"], "event_key": identity,
             "historical_received": ((cumulative - amount) / 100
                                     if cumulative is not None and amount is not None else received / 100),
             "remaining": (delivery - prospective) / 100 if delivery is not None else None,
             "receivable_group_scope": copy.deepcopy(rec.get("receivable_group_scope") or {}),
             "before_rows": before, "group_key": key, "disposition": "write"}
    result = copy.deepcopy(result)
    result.update(bucket="auto", code="E5", ledger_row_ref=int(nonblank[0][0]) if nonblank else int(next(iter(before))),
                  baseline_receipt_audit=audit)

    def conflict(reason):
        audit.update(disposition="conflict", reason=reason)
        result.update(five_cols={}, derived_cols={}, reason=reason)
        result.pop("row_operation", None)
        return result

    if rec.get("forced_code") or rec.get("customer_archive_failed"):
        return None  # Preserve existing input and parent-payment hard gates.
    if delivery is None or amount is None or amount <= 0 or baseline <= 0:
        return conflict("交付额、原始应收或本次回款缺少有效正金额，不能生成保留应收方案")
    if len(nonblank) != 1 or any((cents(row[k]) or 0) < 0 for row in values for k in NUMERIC - {"差异"}):
        return conflict("原始应收行不能唯一识别或存在负金额，不能改写已有拆行结构")
    if not identity or not day or not posting_day or not receipt["收款方式"]:
        return conflict("本次回款缺少父 AR、核销记录标识、SOD 或收款信息，无法证明写入身份")
    if journal and cents(journal.get("baseline_receivable")) != baseline:
        return conflict("原始应收合计与已发布的历史基线不一致")
    previous_scope = journal.get("receivable_group_scope") or {}
    current_scope = rec.get("receivable_group_scope") or {}
    if previous_scope and (any(previous_scope.get(field) != current_scope.get(field)
                               for field in ("basis", "so", "ledger_sod")) or
                           not set(previous_scope.get("source_sods") or []).issubset(
                               current_scope.get("source_sods") or [])):
        return conflict("历史 SO 交付口径的应收组范围缺失或变化，禁止退回单个 SOD 金额")
    events = journal.get("events") or {}
    prior = events.get(identity)
    paid_rows = [(ref, row) for ref, row in before.items() if (cents(row["回款明细"]) or 0)>0]
    paid_rows.sort(key=lambda pair: int(pair[0]))
    # A slot is a legacy display ordinal. Uploaded materials can move a receipt
    # onto the main AR row, so bind each durable event to a unique signature.
    bindings = {}; occupied = set()
    for event_id,event in events.items():
        found=[(ref,row) for ref,row in paid_rows if signature(row)==signature(event) and row["是否结账"]=="是"]
        if len(found)!=1 or found[0][0] in occupied:
            return conflict("历史回款事件在当前 SO/SOD 内无法按金额、日期、方式唯一重定位")
        bindings[event_id]=found[0];occupied.add(found[0][0])
    if prior:
        if signature(prior) != signature(receipt):
            return conflict("本次回款身份已存在，但金额、日期或方式不一致；禁止重复写入或覆盖")
        ref, actual = bindings[identity]
        expected_settled = journal.get("settled", False)
        if expected_settled and any(row["是否结账"] != "是" for row in values):
            return conflict("回款已写但原始应收行尚未完成结账，需核对恢复条件")
        if not expected_settled and (nonblank[0][1]["是否结账"] == "是" or
                                     any(row["计提"] is not None or row["差异"] is not None for row in values)):
            return conflict("未收齐的已登记回款组出现提前结账、计提或差异，需核对恢复条件")
        if journal.get("accrual") is not None and sum(cents(row["计提"]) or 0 for row in values) != cents(journal["accrual"]):
            return conflict("已发布的计提与表内计提合计不一致，需核对恢复条件")
        audit.update(disposition="skip", historical_received=(cumulative - amount) / 100 if cumulative is not None else (received - amount) / 100,
                     remaining=(delivery - (cumulative if cumulative is not None else received)) / 100,
                     reason="本次父 AR 和核销记录已发布，实际回款行一致；跳过且不重复累计")
        result.update(code="OK_BASELINE_RECEIPT_APPLIED", ledger_row_ref=int(ref),
                      five_cols={k: actual[k] for k in ("计提", "回款明细", "是否结账", "收款时间", "收款方式")},
                      derived_cols={}, reason=audit["reason"])
        return result
    if any(signature(row) == signature(receipt) and ref not in occupied
           for ref,row in paid_rows):
        return conflict("表内存在相同回款，但缺少本次父 AR 和核销记录的归属证据；不能直接跳过或新增")
    anchor_ref, anchor = nonblank[0]
    if anchor["回款明细"] is not None or anchor["是否结账"] == "是":
        return conflict("原始应收行已有回款或已结账，不能继续作为保留应收的未结清行")
    if any(row["计提"] is not None or row["差异"] is not None for row in values):
        return conflict("尚未收齐的拆分组已有计提或差异，禁止覆盖历史值")
    if any(row["应收金额"] is None and (row["回款明细"] is None or row["是否结账"] != "是") for row in values):
        return conflict("已有空应收行未完整记录回款，需核对部分写入状态")
    before_amount = cumulative - amount if cumulative is not None else received
    after_amount = before_amount + amount
    remaining = delivery - after_amount
    if before_amount < 0 or remaining < -SETTLEMENT_CENTS:
        return conflict("累计回款口径不成立或超过最新交付额，禁止写入")
    settled = abs(remaining) <= SETTLEMENT_CENTS
    five = {**receipt, "计提": delivery / 100 if settled else None,
            "是否结账": "是", "实收SOD": rec["sod"]}
    derived = {"差异": (baseline - delivery) / 100} if settled else {}
    step = {"event_key": identity, "case_id": result["case_id"], "ar": rec["ar"],
            "so": rec["so"], "sod": rec["sod"], "order": list(rec.get("writeoff_sequence_key") or [str(posting_day), rec["ar"]]),
            "historical_received": before_amount / 100, "current_received": amount / 100,
            "cumulative_received": after_amount / 100, "remaining": remaining / 100,
            "settled": settled, "five_cols": five, "derived_cols": derived}
    audit.update(historical_received=before_amount / 100, remaining=remaining / 100,
                 reason="保留原始应收；新增空应收回款行" + ("，收齐后全部结账并按交付额计提" if settled else "，原始行保持未结账"))
    if audit["receivable_group_scope"]:
        audit["reason"] = f"按 SO 最新交付额 {delivery / 100:.2f} 核销既有应收组；" + audit["reason"]
    result.update(five_cols=five, derived_cols=derived, reason=audit["reason"], row_operation={
        "type": OPERATION, "schema_version": 1, "group_key": key,
        "baseline_receivable": baseline / 100, "latest_delivery": delivery / 100,
        "before_rows": before, "anchor_row": int(anchor_ref), "insert_after": max(map(int, before)),
        "steps": [step], "settled": settled,
    })
    return result


def combine(results: list[dict]) -> None:
    groups = {}
    for item in results:
        if (item.get("row_operation") or {}).get("type") == OPERATION:
            groups.setdefault(item["row_operation"]["group_key"], []).append(item)
    for members in groups.values():
        members.sort(key=lambda item: tuple(map(str, item["row_operation"]["steps"][0]["order"])))
        op = copy.deepcopy(members[0]["row_operation"])
        steps = [copy.deepcopy(item["row_operation"]["steps"][0]) for item in members]
        prior = sum(cents(row["回款明细"]) or 0 for row in op["before_rows"].values())
        identities = set()
        cases = set()
        error = ""
        for step in steps:
            if (step["event_key"] in identities or step["case_id"] in cases or
                    cents(step["historical_received"]) != prior):
                error = "本批回款身份重复或累计与历史及逐笔顺序不一致；禁止跳笔、漏算或重复累计"
                break
            identities.add(step["event_key"])
            cases.add(step["case_id"])
            prior += cents(step["current_received"])
        if any(item["row_operation"]["before_rows"] != op["before_rows"] or
               item["row_operation"]["latest_delivery"] != op["latest_delivery"] or
               item["baseline_receipt_audit"].get("receivable_group_scope") !=
               members[0]["baseline_receipt_audit"].get("receivable_group_scope") for item in members):
            error = "同一 SOD 的本批材料快照或交付额不一致"
        if any((item.get("baseline_receipt_audit") or {}).get("group_key") == op["group_key"] and
               item["baseline_receipt_audit"]["disposition"] == "conflict" for item in results):
            error = "同一 SOD 本批存在身份或金额冲突，不能执行不完整的回款链"
        if error:
            for item in members:
                item["baseline_receipt_audit"].update(disposition="conflict", reason=error)
                item.update(five_cols={}, derived_cols={}, reason=error)
                item.pop("row_operation", None)
            continue
        op.update(steps=steps, settled=steps[-1]["settled"])
        # The final event owns settlement, even if an earlier event is already
        # within the business tail tolerance. No event amount is rounded away.
        for step in steps[:-1]:
            step["settled"] = False
            step["five_cols"]["计提"] = None
            step["derived_cols"] = {}
        for index, item in enumerate(members):
            item.update(row_operation=op, split_chain_group_id=f"{OPERATION}|{op['group_key']}",
                        split_chain_index=index, split_chain_count=len(steps),
                        five_cols=steps[index]["five_cols"], derived_cols=steps[index]["derived_cols"])


def check_scope(item: dict, rows: dict) -> dict | None:
    """Validate the SO delivery basis for both ordinary and preserved splits."""
    source = item.get("split_payment_source") or {}
    scope = source.get("receivable_group_scope") or {}
    if not scope:
        return None
    try:
        all_so_rows = [row for row in rows.values() if row.get("SO") == item.get("so")]
        if (scope.get("basis") != "so_latest_delivery" or scope.get("so") != item.get("so") or
                scope.get("ledger_sod") != item.get("sod") or not all_so_rows or
                any(row.get("SOD") != item.get("sod") for row in all_so_rows) or
                not scope.get("source_sods") or
                set(scope["source_sods"]) != set(source.get("all_sods") or []) or
                item["sod"] not in scope["source_sods"] or
                sum(cents(row.get("应收金额")) or 0 for row in all_so_rows) != cents(scope.get("baseline_receivable")) or
                (cents(source.get("so_delivery_local")) or 0) <= 0 or
                cents(source.get("delivery_local")) != cents(source.get("so_delivery_local"))):
            raise ValueError("SO 交付额、来源 SOD 范围或原始应收组在写前复核中不一致")
    except (ValueError, KeyError, TypeError) as exc:
        return {"verdict": "conflict", "reason": str(exc)}
    return None


def check(item: dict, rows: dict) -> dict:
    audit = item["baseline_receipt_audit"]
    verdict = audit["disposition"]
    error = audit.get("reason") or ""
    actual = {str(ref): normalized(row) for ref, row in rows.items()
              if row.get("SO") == item.get("so") and row.get("SOD") == item.get("sod")}
    if actual != audit["before_rows"]:
        verdict, error = "conflict", "同一 SO/SOD 的行身份、金额或状态在生成方案后变化，需重新判定"
    op = item.get("row_operation") or {}
    source = item.get("split_payment_source") or {}
    scope = source.get("receivable_group_scope") or {}
    if scope != (audit.get("receivable_group_scope") or {}):
        return {"verdict": "conflict", "reason": "SO 应收组范围与判定审计不一致"}
    scope_error = check_scope(item, rows)
    if scope_error:
        return scope_error
    if verdict == "write":
        if op.get("type") != OPERATION or op.get("before_rows") != actual:
            return {"verdict": "conflict", "reason": "保留应收方案缺少完整的写前行快照"}
        try:
            if (op.get("group_key") != group_key(item["so"], item["sod"]) or
                    audit.get("group_key") != op["group_key"]):
                raise ValueError("拆行组身份与当前 SO/SOD 不一致")
            baseline = sum(cents(row["应收金额"]) or 0 for row in actual.values())
            running = sum(cents(row["回款明细"]) or 0 for row in actual.values())
            delivery = cents(op["latest_delivery"])
            if baseline != cents(op["baseline_receivable"]):
                raise ValueError("原始应收合计不守恒")
            anchors = [int(ref) for ref, row in actual.items() if row["应收金额"] is not None]
            if (anchors != [op["anchor_row"]] or op["insert_after"] != max(map(int, actual))
                    or int(item["ledger_row_ref"]) != op["anchor_row"]
                    or actual[str(op["anchor_row"])]["回款明细"] is not None
                    or actual[str(op["anchor_row"])]["是否结账"] == "是"
                    or any(row["计提"] is not None or row["差异"] is not None for row in actual.values())):
                raise ValueError("原始应收行、插入位置或写前计提状态不符合保留应收规则")
            seen = set()
            seen_cases = set()
            for index, step in enumerate(op["steps"]):
                amount = cents(step["current_received"])
                if (not step["event_key"] or step["event_key"] in seen or step["case_id"] in seen_cases or
                        amount <= 0 or cents(step["historical_received"]) != running):
                    raise ValueError("回款身份或运行累计不一致")
                seen.add(step["event_key"])
                seen_cases.add(step["case_id"])
                running += amount
                remaining = delivery - running
                settled = abs(remaining) <= SETTLEMENT_CENTS and index == len(op["steps"]) - 1
                five = step["five_cols"]
                if (remaining < -SETTLEMENT_CENTS or remaining != cents(step["remaining"]) or running != cents(step["cumulative_received"])
                        or settled != step["settled"] or (settled and index != len(op["steps"]) - 1)
                        or cents(five["回款明细"]) != amount or five["是否结账"] != "是"
                        or five["实收SOD"] != item["sod"] or not five["收款时间"] or not five["收款方式"]):
                    raise ValueError("回款、剩余未收或结清计划不一致")
                accrual = cents(five.get("计提"))
                gate = item.get("so_accrual_audit") or {}
                deferred = (gate.get("rule") == "all_sods_under_so_before_accrual" and
                            gate.get("all_settled") is False and len(gate.get("all_sods") or []) > 1 and
                            bool(gate.get("unsettled_sods") or gate.get("missing_delivery_sods")))
                if settled and accrual is None and not deferred:
                    raise ValueError("已结清且未触发整 SO 延后计提，必须按交付额计提")
                if deferred and accrual is not None:
                    raise ValueError("同 SO 尚有 SOD 未结清，当前不得提前计提")
                if accrual is not None and (not settled or accrual != delivery):
                    raise ValueError("计提必须在结清后按交付额填写一次")
                if accrual is not None and not rows[op["anchor_row"]].get("_差异列存在"):
                    raise ValueError("最终计提缺少差异列，不能生成完整结清方案")
                if step.get("derived_cols") != ({"差异": (baseline - delivery) / 100} if accrual is not None else {}):
                    raise ValueError("计提与差异计划不一致")
            selected = op["steps"][int(item.get("split_chain_index") or 0)]
            source = item.get("split_payment_source") or {}
            source_identity = event_key({**item, "writeoff_sequence_key": source.get("writeoff_sequence_key")})
            if (source_identity != selected["event_key"] or selected["ar"] != item["ar"] or
                    selected["so"] != item["so"] or selected["sod"] != item["sod"] or
                    cents(source.get("amount_local")) != cents(selected["current_received"]) or
                    cents(source.get("delivery_local")) != delivery or
                    (source.get("cumulative_local") is not None and
                     cents(source["cumulative_local"]) != cents(selected["cumulative_received"]))):
                raise ValueError("本笔金额、交付额、累计或回款身份与固定来源不一致")
            if (selected["event_key"] != audit["event_key"] or selected["case_id"] != item["case_id"]
                    or selected["five_cols"] != item["five_cols"] or selected["derived_cols"] != item.get("derived_cols", {})
                    or op["settled"] != op["steps"][-1]["settled"]):
                raise ValueError("本笔回款与完整拆行计划不一致")
        except (ValueError, KeyError, TypeError, IndexError) as exc:
            verdict, error = "conflict", str(exc)
    return {"verdict": verdict, "reason": error}


def merge_journal(existing: dict, checked: dict) -> dict:
    """Called only after successful write/readback, or in isolated review copies."""
    import receipt_history
    updated = receipt_history.merge(existing, checked)
    writes = checked.get("write") or []
    handled = set()
    for item in writes:
        op = item.get("row_operation") or {}
        scope = (item.get("split_payment_source") or {}).get("receivable_group_scope")
        if scope and op.get("type") != OPERATION:
            key = group_key(item["so"], item["sod"])
            entry = updated.setdefault(key, {"baseline_receivable": scope["baseline_receivable"],
                                             "events": {}, "scope_only": True})
            if not entry.get("scope_only") or entry["baseline_receivable"] != scope["baseline_receivable"]:
                raise ValueError("普通拆分不能覆盖保留应收历史或改变原始应收基线")
            entry["receivable_group_scope"] = copy.deepcopy(scope)
            continue
        if op.get("type") != OPERATION or op["group_key"] in handled:
            continue
        handled.add(op["group_key"])
        members = [other for other in writes if (other.get("row_operation") or {}).get("group_key") == op["group_key"]]
        if {other["case_id"] for other in members} != {step["case_id"] for step in op["steps"]}:
            raise ValueError("正式回款身份登记缺少完整拆行组")
        entry = updated.setdefault(op["group_key"], {"baseline_receivable": op["baseline_receivable"], "events": {}})
        if entry["baseline_receivable"] != op["baseline_receivable"]:
            raise ValueError("正式回款身份登记不能改变原始应收基线")
        entry.pop("scope_only", None)
        events = entry["events"]
        offset = sum(row["应收金额"] is None and row["回款明细"] is not None for row in op["before_rows"].values())
        for index, step in enumerate(op["steps"]):
            event = {key: step["five_cols"][key] for key in ("回款明细", "收款时间", "收款方式")}
            event.update(slot=offset + index, ar=step["ar"], so=step["so"], sod=step["sod"],
                         case_id=step["case_id"], event_key=step["event_key"])
            old = events.get(step["event_key"])
            if old is not None and old != event:
                raise ValueError("本次回款身份已登记且与计划不一致，禁止覆盖")
            events[step["event_key"]] = event
        entry.update(settled=op["settled"], latest_delivery=op["latest_delivery"],
                     accrual=op["steps"][-1]["five_cols"].get("计提"))
        scope = (item.get("split_payment_source") or {}).get("receivable_group_scope")
        if scope:
            entry["receivable_group_scope"] = copy.deepcopy(scope)
    for item in writes:
        for backfill in item.get("so_accrual_backfills") or []:
            entry = updated.get(group_key(backfill.get("so"), backfill.get("sod")))
            if entry and (backfill.get("_check") or {}).get("verdict") in {"write", "skip"}:
                if not entry.get("settled"):
                    raise ValueError("尚未结清的保留应收组不能登记历史计提补填")
                entry["accrual"] = backfill["accrual"]
    return updated


def validate_journal(groups: dict) -> None:
    if not isinstance(groups, dict):
        raise ValueError("回款身份台账结构无效")
    for key, group in groups.items():
        if not isinstance(group, dict) or not isinstance(group.get("events"), dict) or (cents(group.get("baseline_receivable")) or 0) <= 0:
            raise ValueError("回款身份台账缺少历史应收或事件集合")
        scope = group.get("receivable_group_scope")
        if group.get("scope_only") and ((not scope and not group.get("ordinary_events")) or group["events"]):
            raise ValueError("普通应收组口径登记不能包含保留应收回款事件")
        if scope and (not isinstance(scope, dict) or scope.get("basis") != "so_latest_delivery" or
                      key != group_key(scope.get("so"), scope.get("ledger_sod")) or
                      cents(scope.get("baseline_receivable")) != cents(group.get("baseline_receivable")) or
                      not isinstance(scope.get("source_sods"), list) or not scope["source_sods"] or
                      scope.get("ledger_sod") not in scope["source_sods"]):
            raise ValueError("回款身份台账的 SO 交付额范围与原始应收组不一致")
        ordinary = group.get("ordinary_events") or {}
        if not isinstance(ordinary, dict):
            raise ValueError("普通回款身份登记结构无效")
        ordinary_signatures = set()
        for identity, event in ordinary.items():
            try:
                parts = json.loads(identity)
                sig = event['signature']
                if (len(parts) != 4 or group_key(parts[1], parts[2]) != key or not all(parts)
                        or len(sig) != 3 or (cents(sig[0]) or 0) <= 0
                        or not common.norm_date(sig[1]) or not sig[2]
                        or tuple(sig) in ordinary_signatures):
                    raise ValueError("普通回款身份或金额无效")
                ordinary_signatures.add(tuple(sig))
            except (ValueError, TypeError, KeyError, IndexError) as exc:
                raise ValueError("普通回款身份无法核实") from exc
        slots = set()
        for identity, event in group["events"].items():
            if (not isinstance(event, dict) or identity != event.get("event_key")
                    or key != group_key(event.get("so"), event.get("sod"))
                    or not isinstance(event.get("slot"), int) or event["slot"] < 0 or event["slot"] in slots
                    or (cents(event.get("回款明细")) or 0) <= 0 or not common.norm_date(event.get("收款时间"))
                    or not event.get("收款方式")):
                raise ValueError("回款身份、金额或已写行序号无效")
            try:
                parts = json.loads(identity)
                if len(parts) != 4 or parts[:3] != [event["ar"], event["so"], event["sod"]] or not parts[3]:
                    raise ValueError("回款身份与父 AR、SO/SOD 不一致")
            except (ValueError, TypeError, KeyError) as exc:
                raise ValueError("回款身份键无法核实") from exc
            slots.add(event["slot"])
