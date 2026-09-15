"""Re-plan an unambiguous current receipt using the ordinary financial rules."""
import copy
import math
import common
import baseline_receipts as BR

POLICY = 'verified-current-receipt-v1'
CLEAR = {'jiti':None,'huikuan':None,'jiezhang':'否','shoukuan_time':None,'shoukuan_way':None,'chayi':None}


def planned(item):
    return {key:copy.deepcopy(item.get(key)) for key in ('ledger_row_ref','five_cols','derived_cols','row_operation')}


def candidate(rec, result, ledger, rates=None, thr=0.01, year_now=None):
    if ledger is None or rec.get('_receipt_correction_shadow') or rec.get('forced_code') or rec.get('customer_archive_failed'):
        return None
    import receipt_history
    history=receipt_history.candidate(rec,result,ledger)
    if history is not None:return history
    if rec.get('receivable_group_scope'):return None
    so,sod=rec.get('so'),rec.get('sod')
    if not so or not sod:return None
    rows=BR.ledger_rows(ledger,so,sod)
    if not rows:return None
    if getattr(ledger,'baseline_receipt_state',{}).get(BR.group_key(so,sod)):
        return None  # Published multi-row journals retain their own event-aware planner.
    try:
        amount,delivery,cumulative=(float(rec[k]) for k in ('amount_local','deliver_local','cumulative_received_local'))
        if not all(math.isfinite(v) and v>0 for v in (amount,delivery,cumulative)):return None
        if amount>cumulative+0.011 or cumulative>delivery+1.001:return None
        prior=cumulative-amount
        possible=[]
        for ref,row in rows.items():
            others=[r for key,r in rows.items() if key!=ref]
            paid=sum(float(r['回款明细'] or 0) for r in others)
            if row['应收金额'] is not None and row['应收金额']>0 and abs(paid-prior)<=0.011:
                possible.append(ref)
        if len(possible)!=1:return None
        ref=possible[0]
    except (KeyError,TypeError,ValueError):return None
    identity=BR.event_key(rec)
    arrival,posting=common.norm_date(rec.get('shoukuan_date')),common.norm_date(rec.get('hexiao_date'))
    if not identity or arrival is None or posting is None:return None
    for key,row in rows.items():
        if key==ref or not row['回款明细']:continue
        day=common.norm_date(row['收款时间'])
        if row['回款明细']<0 or day is None or day>posting or not row['收款方式']:return None
    from classify_hexiao import classify_one
    shadow=copy.deepcopy(ledger)
    shadow.row_snapshot[int(ref)].update(CLEAR)
    answer=classify_one({**rec,'_receipt_correction_shadow':True},shadow,rates or {},thr,year_now or posting.year)
    op=answer.get('row_operation') or {}
    if answer.get('bucket')!='auto' or answer.get('ledger_row_ref')!=int(ref) or op.get('type') not in (None,'split_below'):
        return None
    if not answer.get('five_cols') or answer.get('baseline_receipt_audit'):return None
    expected=BR.normalized({**rows[ref],**answer['five_cols'],**answer.get('derived_cols',{})})
    if op.get('type')=='split_below':expected['应收金额']=op['paid_receivable']
    if not op and expected==rows[ref]:return None
    answer['receipt_correction']={'policy':POLICY,'event_key':identity,'before':rows[ref],
                                  'before_rows':rows,'after':expected,'planned':planned(answer),
                                  'amount':amount,'delivery':delivery,'cumulative':cumulative}
    answer['reason']='按已核实的本次核销结果补写或纠正；'+str(answer.get('reason') or '')
    answer.setdefault('warning_codes',[]).append('W_VERIFIED_RECEIPT_CORRECTION')
    return answer


def check(item, rows):
    audit=item.get('receipt_correction')
    if not audit:return None
    if audit.get('history_mode'):
        import receipt_history
        return receipt_history.check(item,rows)
    def conflict(reason):return {'verdict':'conflict','reason':reason}
    if audit.get('policy')!=POLICY or item.get('baseline_receipt_audit') or (item.get('row_operation') or {}).get('type') not in (None,'split_below','split_payment_chain','same_so_multi_sod_aggregate','settlement_tail_aggregate','preserve_aggregate_tail_tolerance'):
        return conflict('本次核销覆盖计划类型不一致')
    source=item.get('split_payment_source') or {}
    try:
        if not all(math.isfinite(float(audit[k])) and float(audit[k])>0 for k in ('amount','delivery','cumulative')):return conflict('核销覆盖金额无效')
        if not math.isfinite(float(source['amount_local'])) or abs(float(source['amount_local'])-audit['amount'])>0.011:return conflict('核销覆盖金额与取数依据不同')
        event=BR.event_key({'ar':item.get('ar'),'so':item.get('so'),'sod':item.get('sod'),
                            'writeoff_sequence_key':source.get('writeoff_sequence_key'),
                            'parent_allocation_audit':item.get('parent_allocation_audit')})
        if not event or event!=audit['event_key']:return conflict('核销覆盖的事件身份无法核实')
        if planned(item)!=audit['planned']:return conflict('核销覆盖计划在判定后被改动')
        clean=copy.deepcopy(item);clean.pop('receipt_correction')
        import validate_plan
        actual_result=validate_plan.check_one(clean,rows)
        if actual_result['verdict']=='skip':return actual_result
        current={str(ref):BR.normalized(row) for ref,row in rows.items() if row.get('SO')==item['so'] and row.get('SOD')==item['sod']}
        if current!=audit['before_rows']:return conflict('核销覆盖目标在计划后发生变化')
        if actual_result['verdict']=='write':
            return {'verdict':'write','reason':'普通核销方案及原值复核通过，按本次结果覆盖并保留变更记录'}
        ref=int(item['ledger_row_ref']);shadow=copy.deepcopy(rows)
        shadow[ref].update({'计提':None,'回款明细':None,'是否结账':'否','收款时间':None,'收款方式':None,'差异':None})
        checked=validate_plan.check_one(clean,shadow)
        if checked['verdict']!='write':return checked
        return {'verdict':'write','reason':'按已核实的本次核销结果覆盖；修改前后值已记录'}
    except (KeyError,TypeError,ValueError):return conflict('核销覆盖证据不完整')


def finalize_plans(records):
    # Ordinary batch grouping/accrual calculations finalize the financial plan.
    for item in records:
        audit=item.get('receipt_correction')
        if audit and item.get('bucket')=='auto' and not audit.get('history_mode'):
            audit['planned']=planned(item)
            audit['after']=BR.normalized({**audit['before'],**item.get('five_cols',{}),**item.get('derived_cols',{})})
            op=item.get('row_operation') or {}
            if op.get('type')=='split_below':audit['after']['应收金额']=op['paid_receivable']
