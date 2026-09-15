"""Identify existing receipt events, preserving AR and avoiding duplicate money."""
import copy
import common
import baseline_receipts as BR

MODE = 'unchanged-delivery-history-v1'
FIVE = ('计提','回款明细','是否结账','收款时间','收款方式')

def plan(rec, result, rows, journal):
    amount, delivery, cumulative = (BR.cents(rec.get(k)) for k in ('amount_local','deliver_local','cumulative_received_local'))
    arrival, posting = common.norm_date(rec.get('shoukuan_date')), common.norm_date(rec.get('hexiao_date'))
    identity=BR.event_key(rec)
    if not identity or not arrival or not posting or not rows or amount is None or delivery is None or cumulative is None:return None
    if not (0 < amount <= cumulative <= delivery + BR.SETTLEMENT_CENTS):return None
    baseline = sum(BR.cents(r['应收金额']) or 0 for r in rows.values())
    if baseline <= 0:return None
    if journal and not journal.get('scope_only'):return None
    if journal and BR.cents(journal.get('baseline_receivable'))!=baseline:return None
    if any((BR.cents(r[k]) or 0)<0 for r in rows.values() for k in ('应收金额','回款明细')):return None
    paid={ref:r for ref,r in rows.items() if (BR.cents(r['回款明细']) or 0)>0}
    if not paid:return None
    received=sum(BR.cents(r['回款明细']) for r in paid.values())
    dirty = any(not common.norm_date(r['收款时间']) or not r['收款方式'] or r['是否结账']!='是' for r in paid.values())
    day=common.receipt_time(arrival,posting).isoformat()
    way=common.pay_way(rec.get('status') or '',arrival,posting)
    events=journal.get('ordinary_events') or {}
    owned=events.get(identity)
    matches=[ref for ref,r in paid.items() if BR.cents(r['回款明细'])==amount and r['收款时间'] in (arrival.isoformat(),day,posting.isoformat())]
    if owned:
        old_matches=[ref for ref,r in paid.items() if list(BR.signature(r))==owned['signature']]
        if old_matches:
            matches=old_matches
        elif (BR.cents(owned['signature'][0])!=amount or len(matches)!=1
              or sum(BR.cents(r['回款明细']) for r in paid.values() if common.norm_date(r['收款时间']) and common.norm_date(r['收款时间'])<=posting)!=cumulative):
            return None
        # Keep the durable identity and the old journal in the write audit.
    if len(matches)>1:return None
    # A single conserved receipt can have a mistyped date. Amount alone is
    # insufficient: require the complete current cumulative and ordinary AR split.
    if not matches and not owned and len(paid)==1 and cumulative==amount:
        ref, row = next(iter(paid.items()))
        same_date = row['收款时间'] in (arrival.isoformat(),day,posting.isoformat())
        exact_conserved = (BR.cents(row['回款明细'])==amount and baseline==delivery
                           and all(BR.cents(r['应收金额'])==BR.cents(r['回款明细']) for r in paid.values()))
        small_correction = (same_date and len(rows)==1
                            and abs(BR.cents(row['回款明细'])-amount)<=BR.SETTLEMENT_CENTS)
        if not dirty and (exact_conserved or small_correction):matches=[ref]
    target=None;op=None;kind='existing'
    if matches:
        ref=matches[0];target=rows[ref]
        if BR.cents(target['回款明细'])!=amount and not (len(paid)==1 and cumulative==amount and abs(BR.cents(target['回款明细'])-amount)<=BR.SETTLEMENT_CENTS):return None
        if any(key!=identity and event['signature']==list(BR.signature(target)) for key,event in events.items()):return None
        five={k:target[k] for k in FIVE}
        five.update(回款明细=amount/100,收款时间=day,收款方式=way,是否结账='是')
        remaining=delivery-(received-BR.cents(target['回款明细'])+amount)
        if len(rows)==1 and not dirty and abs(remaining)<=BR.SETTLEMENT_CENTS:
            five['计提']=delivery/100
    else:
        if baseline!=delivery or dirty or received>delivery:return None
        if owned:return None
        # A same-size unidentified receipt is ambiguous, never silently counted twice.
        if any(BR.cents(r['回款明细'])==amount for r in paid.values()):return None
        earlier=sum(BR.cents(r['回款明细']) for r in paid.values() if common.norm_date(r['收款时间'])<=posting)
        if earlier != cumulative-amount:return None
        # Only a fully conserved ordinary split can absorb a missing earlier receipt.
        anchors=[ref for ref,r in rows.items() if r['回款明细'] is None and r['是否结账']=='否' and BR.cents(r['应收金额'])==delivery-received]
        if len(anchors)!=1 or received+amount>=delivery:return None
        if any(BR.cents(r['应收金额'])!=BR.cents(r['回款明细']) for r in paid.values()):return None
        ref=anchors[0];target=rows[ref];kind='missing'
        if target['计提'] is not None or target['差异'] is not None:return None
        remaining=delivery-received-amount
        five={'计提':None,'回款明细':amount/100,'是否结账':'是','收款时间':day,'收款方式':way}
        op={'type':'split_below','source_receivable':target['应收金额'],'paid_receivable':amount/100,'unpaid_receivable':remaining/100,'latest_delivery':delivery/100,'cumulative_received':(received+amount)/100,'existing_received':received/100,'current_received':amount/100,'baseline_receivable':delivery/100,'paid_side_receivable_total':(received+amount)/100,'business_rows':list(map(int,rows)), 'inserted_five_cols':{'计提':None,'回款明细':None,'是否结账':'否','收款时间':None,'收款方式':None,'实收SOD':rec['sod']}}
    five['实收SOD']=rec['sod']
    answer=copy.deepcopy(result)
    answer.update(bucket='auto',code='E5',ledger_row_ref=int(ref),five_cols=five,derived_cols={'差异':target['差异']} if target['差异'] is not None else {})
    if kind=='existing' and len(rows)==1 and not dirty and abs(remaining)<=BR.SETTLEMENT_CENTS:
        # Ordinary first-time settlement leaves an absent zero difference blank.
        # Preserve that representation when reviewing the same receipt; real
        # differences and existing explicit values still require verification.
        if baseline != delivery or target['差异'] is not None:
            answer['derived_cols']={'差异':(baseline-delivery)/100}
    if dirty:
        answer.setdefault('warning_codes',[]).append('W_UNRELATED_RECEIPT_INCOMPLETE')
    answer.pop('row_operation',None)
    if op:answer['row_operation']=op
    after=BR.normalized({**target,**five,**answer.get('derived_cols',{})})
    if op:after['应收金额']=op['paid_receivable']
    import receipt_correction as R
    answer['receipt_correction']={'policy':R.POLICY,'history_mode':MODE,'kind':kind,'event_key':identity,'before':target,'before_rows':copy.deepcopy(rows),'after':after,'amount':amount/100,'delivery':delivery/100,'cumulative':cumulative/100,'baseline_receivable':baseline/100,'current_ledger_received':received/100,'current_remaining':remaining/100,'rec':copy.deepcopy(rec),'journal':copy.deepcopy(journal)}
    answer['receipt_correction']['planned']=R.planned(answer)
    answer['reason']=('已定位同笔已有回款，仅更新核销字段，不重复扣款' if kind=='existing' else '补写遗漏回款，按当前全部已登记回款计算余额，保留后续记录')
    return answer

def candidate(rec,result,ledger):
    if ledger is None:return None
    try:
        return plan(rec,result,BR.ledger_rows(ledger,rec['so'],rec['sod']),getattr(ledger,'baseline_receipt_state',{}).get(BR.group_key(rec['so'],rec['sod'])) or {})
    except (KeyError,TypeError,ValueError):return None

def check(item,rows):
    import receipt_correction as R
    import validate_plan as V
    audit=item['receipt_correction']
    def bad(reason):return {'verdict':'conflict','reason':reason}
    try:
        rec=audit['rec'];source=item['split_payment_source']
        if any(rec.get(k)!=item.get(k) for k in ('ar','so','sod')):return bad('历史回款身份与当前计划不同')
        if BR.event_key(rec)!=audit['event_key'] or source.get('writeoff_sequence_key')!=rec.get('writeoff_sequence_key'):return bad('历史回款事件标识发生变化')
        for field,key in [('amount_local','amount_local'),('delivery_local','deliver_local'),('cumulative_local','cumulative_received_local')]:
            if BR.cents(source.get(field))!=BR.cents(rec.get(key)):return bad('历史回款来源金额发生变化')
        expected=plan(rec,item,audit['before_rows'],audit['journal'])
        if expected is None or R.planned(expected)!=R.planned(item):return bad('历史回款方案不能由原值和取数依据重建')
        clean=copy.deepcopy(item);clean.pop('receipt_correction')
        current={str(ref):BR.normalized(r) for ref,r in rows.items() if r.get('SO')==item['so'] and r.get('SOD')==item['sod']}
        checked=V.check_one(clean,rows)
        cohort=rec.get('existing_sod_receipt_audit')
        if cohort:
            now={str(ref):BR.normalized(r) for ref,r in rows.items() if r.get('SO')==item['so']}
            before_so=cohort['before_so_rows'];after_so=cohort.get('after_so_rows') or {}
            if set(now)!=set(before_so) or any(row not in (before_so[ref],after_so.get(ref)) for ref,row in now.items()):
                return bad('已有多 SOD 回款组在计划后出现未经计划的修改')
        if checked['verdict']=='skip':return checked
        if current!=audit['before_rows']:return bad('历史回款组在计划后被修改')
        if audit['kind']=='missing':return checked
        shadow=copy.deepcopy(rows)
        shadow[int(item['ledger_row_ref'])].update({k:None for k in FIVE})
        shadow[int(item['ledger_row_ref'])]['是否结账']='否'
        checked=V.check_one(clean,shadow)
        if checked['verdict']!='write':return checked
        return {'verdict':'write','reason':'同笔回款原值及核销依据已核实，仅更新字段，不拆行或重复扣减'}
    except (KeyError,TypeError,ValueError):return bad('历史回款核验依据不完整')

def merge(existing,checked):
    updated=copy.deepcopy(existing)
    for item in (checked.get('write') or [])+(checked.get('skip') or []):
        audit=item.get('receipt_correction') or {}
        if audit.get('history_mode')!=MODE:continue
        key=BR.group_key(item['so'],item['sod'])
        scope=(item.get('split_payment_source') or {}).get('receivable_group_scope') or {}
        entry=updated.setdefault(key,{'baseline_receivable':audit.get('baseline_receivable',audit['delivery']),'events':{},'scope_only':True})
        if not entry.get('scope_only'):raise ValueError('普通回款不能覆盖保留应收台账')
        if scope:entry['receivable_group_scope']=copy.deepcopy(scope)
        events=entry.setdefault('ordinary_events',{})
        signature=list(BR.signature(audit['after']))
        if any(k!=audit['event_key'] and v['signature']==signature for k,v in events.items()):raise ValueError('不同核销事件不能占用同一回款记录')
        events[audit['event_key']]={'signature':signature}
    return updated


ZERO_CODE = 'OK_ZERO_DELIVERY_NO_RECEIPT'

def zero_rows(rows):
    return bool(rows) and all(BR.cents(row.get('应收金额')) == 0
                             and row.get('回款明细') in (None, 0, 0.0)
                             and row.get('计提') in (None, 0, 0.0)
                             and row.get('差异') in (None, 0, 0.0)
                             and row.get('是否结账') == '是' for row in rows.values())

def zero_candidate(rec, result, ledger):
    if ledger is None or rec.get('forced_code'):return None
    if any(BR.cents(rec.get(key)) != 0 for key in ('amount_local','deliver_local')):return None
    rows=BR.ledger_rows(ledger,rec.get('so'),rec.get('sod'))
    if not zero_rows(rows):return None
    answer=copy.deepcopy(result)
    answer.update(bucket='auto',code=ZERO_CODE,ledger_row_ref=min(map(int,rows)),five_cols={},derived_cols={},
                  zero_delivery_audit={'before_rows':rows},reason='来源交付额及本次核销均为0，零应收子单已结账，无需登记回款')
    return answer

def check_zero(item, rows):
    if item.get('code') != ZERO_CODE:return None
    source=item.get('split_payment_source') or {}
    before=(item.get('zero_delivery_audit') or {}).get('before_rows')
    current={str(ref):BR.normalized(row) for ref,row in rows.items() if row.get('SO')==item.get('so') and row.get('SOD')==item.get('sod')}
    valid=(all(BR.cents(source.get(key))==0 for key in ('amount_local','delivery_local'))
           and not any(item.get(key) for key in ('five_cols','derived_cols','row_operation','so_accrual_backfills'))
           and current==before and zero_rows(current))
    return {'verdict':'skip' if valid else 'conflict','reason':'零金额子单复核通过，无回款写入' if valid else '零金额子单的来源、原值或写入指令不一致'}

def existing_sod_slices(rec, ledger):
    """Reuse a unique dated receipt cohort before calculating unpaid capacity."""
    if rec.get('forced_code') not in (None, '', 'E5'):return None
    arrival=common.norm_date(rec.get('shoukuan_date'));posting=common.norm_date(rec.get('hexiao_date'))
    total=BR.cents(rec.get('default_amount_local'))
    cumulative=BR.cents(rec.get('default_cumulative_received_local'))
    original=BR.cents(rec.get('default_amount_orig'))
    if not arrival or not posting or total is None or total<=0 or cumulative is None or original is None:return None
    sequence=rec.get('writeoff_sequence_key') or []
    if len(sequence)<2 or not sequence[1]:return None
    so=rec['so']; lines={line['sod']:line for line in rec.get('default_sod_lines') or [] if line.get('sod')}
    rows={}
    for sod in {str(ledger.row_snapshot[ref].get('sod') or '') for ref in ledger.so_index.get(so,[])}:
        rows.update(BR.ledger_rows(ledger,so,sod))
    paid={ref:row for ref,row in rows.items() if (BR.cents(row['回款明细']) or 0)>0}
    if not paid or any(not common.norm_date(row['收款时间']) or not row['收款方式'] or row['是否结账']!='是' for row in paid.values()):return None
    days={arrival,posting,common.receipt_time(arrival,posting)}
    cohort={ref:row for ref,row in paid.items() if common.norm_date(row['收款时间']) in days}
    if sum(BR.cents(row['回款明细']) for row in cohort.values())!=total:return None
    if len({row['SOD'] for row in cohort.values()})!=len(cohort):return None
    if sum(BR.cents(row['回款明细']) for row in paid.values() if common.norm_date(row['收款时间'])<=posting)!=cumulative:return None
    resolved=[]; original_used=0
    cohort_audit={'before_so_rows':rows,'after_so_rows':copy.deepcopy(rows),'total_local':total/100,'cumulative_local':cumulative/100,'matched_rows':list(cohort)}
    for index,(ref,row) in enumerate(cohort.items()):
        sod=row['SOD'];line=lines.get(sod) or {};delivery=BR.cents(line.get('deliver_local'))
        value=BR.cents(row['回款明细'])
        if delivery is None or value>delivery:return None
        part_orig=original-original_used if index==len(cohort)-1 else round(original*value/total)
        original_used+=part_orig
        part={key:copy.deepcopy(value) for key,value in rec.items() if not key.startswith('default_') and key not in ('forced_code','forced_reason')}
        part.update(sod=sod,amount_local=value/100,amount_orig=part_orig/100,deliver_local=delivery/100,
                    cumulative_received_local=sum(BR.cents(r['回款明细']) for r in paid.values() if r['SOD']==sod and common.norm_date(r['收款时间'])<=posting)/100,
                    existing_sod_receipt_audit=cohort_audit,
                    match_basis='同 SO 本次核销金额、日期组及累计回款共同确认的已有 SOD 回款')
        # A recognized cohort must also pass the ordinary existing-receipt proof.
        planned_part=candidate(part,{},ledger)
        if planned_part is None:return None
        cohort_audit['after_so_rows'][ref]=planned_part['receipt_correction']['after']
        resolved.append(part)
    return resolved
