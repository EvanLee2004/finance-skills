import copy
import datetime as dt
import unittest
import test_receipt_history as fixture
import classify_hexiao as C
import baseline_receipts as BR
import validate_plan as V
import writeoff_duplicate_audit as WDA


def ledger_rows(ledger):
    result={}
    for sod in ledger.sod_index:
        result.update({int(k):{**v,'_差异列存在':True} for k,v in BR.ledger_rows(ledger,'SO_TEST',sod).items()})
    return result


def parent(amount=70, explicit=False):
    orders = [{'so':'SO_A','deliver':40.,'deliver_local':40.,'currency':'CNY','delivery_date':dt.date(2026,6,1)},
              {'so':'SO_B','deliver':60.,'deliver_local':60.,'currency':'CNY','delivery_date':dt.date(2026,6,1)}]
    if explicit:
        for order in orders:
            order.update(written_off=order['deliver'],written_off_local=order['deliver_local'],written_off_present=True,written_off_local_present=True)
    return {'ar':'AR_TEST','amount_orig':amount,'amount_local':amount,'total_amount_orig':amount,'total_amount_local':amount,'currency':'CNY','hexiao_date':dt.date(2026,8,5),'shoukuan_date':dt.date(2026,8,4),'status':'手工核销','huikuan_type':'整笔回款','orders':orders,'_source_meta':{'historical_detail_rows':0},'sod_lines':{}}


class RecognitionTests(unittest.TestCase):
    def test_logical_identity_reaches_expansion(self):
        payment=parent(100,True)
        C.reconcile_writeoff_details([payment],{'AR_TEST':payment},[],dt.date(2026,8,5))
        self.assertTrue(payment['_writeoff_sequence_key_by_so'])
        self.assertTrue(all(r.get('writeoff_sequence_key') for r in C.expand_payments([payment],{})))

    def test_missing_detail_uses_parent_allocation(self):
        payment=parent()
        C.reconcile_writeoff_details([payment],{'AR_TEST':payment},[],dt.date(2026,8,5))
        records=C.expand_payments([payment],{})
        self.assertFalse(payment.get('_parent_audit_unresolved'))
        self.assertEqual(sum(r.get('amount_local') or 0 for r in records),70.)
        self.assertEqual([r['amount_local'] for r in records],[40.,30.])
        self.assertTrue(all(r.get('parent_allocation_audit') for r in records))
        self.assertEqual(V._whole_parent_gate_error(payment['duplicate_writeoff_audit']),'')

    def test_explicit_amounts_exceeding_parent_remain_blocked(self):
        payment=parent(70,True)
        C.reconcile_writeoff_details([payment],{'AR_TEST':payment},[],dt.date(2026,8,5))
        self.assertTrue(payment.get('_parent_audit_unresolved'))

    def existing(self, **kwargs):
        return fixture.HistoryTest().make(**kwargs)

    def checked(self, rec, ledger):
        plan=C.classify_records([rec],ledger,{})
        self.assertEqual(len(plan['auto']),1,plan)
        item=plan['auto'][0]
        self.assertIsNotNone(item.get('receipt_correction'),item)
        self.assertFalse(item.get('row_operation'),item)
        verdict=V.check_one(item,ledger_rows(ledger))
        self.assertIn(verdict['verdict'],['write','skip'],verdict)
        checked=V.validate(plan,ledger_rows(ledger))
        self.assertEqual(len(checked[verdict['verdict']]),1,checked)
        return item

    def test_existing_receipt_with_changed_delivery_is_not_resplit(self):
        rec,ledger=self.existing()
        rec.update(deliver_local=120,so_delivery_local=120,sod_delivery_local={'SOD_TEST':120})
        item=self.checked(rec,ledger)
        self.assertEqual(item['ledger_row_ref'],3)
        self.assertEqual(item['five_cols']['回款明细'],20)

    def test_unique_same_event_with_wrong_old_date_is_corrected(self):
        rec,ledger=self.existing(arrival='2026-06-10',paid_date='2026-08-04')
        item=self.checked(rec,ledger)
        self.assertEqual(item['five_cols']['收款时间'],'2026-08-05')

    def test_unrelated_incomplete_receipt_does_not_block_date_update(self):
        rec,ledger=self.existing()
        ledger.row_snapshot[2]['yingshou']=79
        ledger.row_snapshot[4]={'so':'SO_TEST','sod':'SOD_TEST','yingshou':1,'huikuan':1,'jiezhang':'否'}
        ledger.so_index['SO_TEST'].append(4);ledger.sod_index['SOD_TEST'].append(4)
        item=self.checked(rec,ledger)
        self.assertIsNone(item['five_cols']['计提'])
        self.assertEqual(item['five_cols']['回款明细'],20)

    def test_unique_current_amount_can_be_corrected_without_resplitting(self):
        rec,ledger=self.existing(amount=19.7,prior=20,arrival='2026-08-04',paid_date='2026-08-04')
        rec.update(deliver_local=20,so_delivery_local=20,sod_delivery_local={'SOD_TEST':20})
        ledger.row_snapshot.pop(2);ledger.so_index['SO_TEST']=[3];ledger.sod_index['SOD_TEST']=[3]
        ledger.row_snapshot[3]['yingshou']=50
        rec['receivable_group_scope']['baseline_receivable']=50
        item=self.checked(rec,ledger)
        self.assertEqual(item['five_cols']['回款明细'],19.7)
        self.assertEqual(item['five_cols']['计提'],20)
        self.assertEqual(item['derived_cols']['差异'],30)

    def test_existing_multi_sod_event_is_recognized_before_capacity(self):
        rec,ledger=self.existing(amount=40,prior=40,arrival='2026-08-04',paid_date='2026-08-04')
        rec.pop('receivable_group_scope')
        ledger.row_snapshot[2]['yingshou']=10
        ledger.row_snapshot[4]={'so':'SO_TEST','sod':'SOD_OTHER','yingshou':20,'huikuan':None,'jiezhang':'否'}
        ledger.so_index['SO_TEST'].append(4);ledger.sod_index['SOD_OTHER']=[4]
        rec.update(forced_code='E5',default_first_sod=True,default_amount_local=40,default_amount_orig=40,default_cumulative_received_local=40,default_sod_lines=[{'sod':'SOD_TEST','deliver_local':50},{'sod':'SOD_OTHER','deliver_local':20}],all_sods=['SOD_TEST','SOD_OTHER'],sod_delivery_local={'SOD_TEST':50,'SOD_OTHER':20},so_delivery_local=70)
        plan=C.classify_records([rec],ledger,{})
        self.assertEqual(len(plan['auto']),1,plan)
        item=plan['auto'][0]
        self.assertEqual(item['ledger_row_ref'],3)
        self.assertTrue(item['receipt_correction']['rec'].get('existing_sod_receipt_audit'))
        self.assertEqual(V.check_one(item,ledger_rows(ledger))['verdict'],'skip')
        changed=ledger_rows(ledger);changed[4]['应收金额']=19
        self.assertEqual(V.check_one(item,changed)['verdict'],'conflict')
        import receipt_history as H
        for code in ('E_PARENT_WRITEOFF_MISMATCH','E_PARENT_ALLOCATION_HISTORY_MISSING'):
            self.assertIsNone(H.existing_sod_slices({**rec,'forced_code':code},ledger))

    def test_zero_delivery_line_has_no_receipt_obligation(self):
        rec,ledger=self.existing(amount=0,prior=0)
        ledger.row_snapshot={3:{'so':'SO_TEST','sod':'SOD_TEST','yingshou':0,'huikuan':None,'jiezhang':'是','shoukuan_time':'2026-08-01','shoukuan_way':'汇'}}
        ledger.so_index={'SO_TEST':[3]};ledger.sod_index={'SOD_TEST':[3]}
        rec.pop('receivable_group_scope')
        rec.update(deliver_local=0,so_delivery_local=0,sod_delivery_local={'SOD_TEST':0})
        plan=C.classify_records([rec],ledger,{})
        self.assertEqual(len(plan['auto']),1,plan)
        self.assertEqual(V.check_one(plan['auto'][0],ledger_rows(ledger))['verdict'],'skip')

    def test_flow_ambiguity_reason_reports_actual_matching_basis(self):
        rec,ledger=self.existing()
        rec.update(flow_hits=2,flow_matched_by='日期+金额(名字不符)')
        item=C.classify_one(rec,ledger,{},.01,2026)
        self.assertEqual(item['bucket'],'auto')
        import build_flow_plan
        flow=build_flow_plan.build_plan({'auto':[item],'hexiao_date':rec['hexiao_date']})['items'][0]
        self.assertEqual(flow['verdict'],'hand')
        self.assertIn('名字不符',flow['reason'])
        self.assertNotIn('同名',flow['reason'])


    def test_export_supplement_reuses_unique_order_date_and_rejects_conflict(self):
        import tempfile
        from pathlib import Path
        import openpyxl
        import common
        aliases=common.load_aliases()
        day=dt.date(2026,8,5)
        for conflict in (False,True):
            with self.subTest(conflict=conflict),tempfile.TemporaryDirectory() as temp:
                work=Path(temp);out=work/'01_智云导出';out.mkdir()
                specs=[('回款记录','回款记录',['AR','核销日期','到账日期','到账金额原币','到账金额本币','原币币种'],
                        [['AR_A',day,day,10,10,'CNY'],['AR_B',day,day,20,20,'CNY']]),
                       ('订单交付','订单交付',['AR','SO','交付额原币','项目交付日期'],
                        [['AR_A','SO_TEST',100,dt.date(2026,6,1)]]),
                       ('核销明细','核销明细',['核销记录NUM','回款记录NUM','核销日期','SO','本次核销金额','本次核销金额本币'],
                        [['HX_A','AR_A',day,'SO_TEST',10,10],['HX_B','AR_B',day,'SO_TEST',20,20]]),
                       ('订单明细','订单明细',['SO','SOD','交付额原币'],[['SO_TEST','SOD_TEST',100]])]
                if conflict:specs[1][3].append(['AR_B','SO_TEST',100,dt.date(2025,6,1)])
                for filename,role,keys,body in specs:
                    b=openpyxl.Workbook();w=b.active;w.append([aliases[role][key][0] for key in keys])
                    for row in body:w.append(row)
                    b.save(out/(filename+'_20260805.xlsx'));b.close()
                payments=C.load_exports(work,target_date=day)
                orders=[p['orders'][0] for p in payments]
                if conflict:self.assertTrue(all(o['delivery_date'] is None and '冲突' in o['delivery_date_issue'] for o in orders))
                else:
                    self.assertTrue(all(o['delivery_date']==dt.date(2026,6,1) for o in orders))
                    self.assertTrue(all(r['delivery_date']==dt.date(2026,6,1) for r in C.expand_payments(payments,{})))

    def test_changed_history_writes_and_repeats_without_duplicate_money(self):
        import tempfile
        from pathlib import Path
        import openpyxl
        import apply_to_copy as writer
        for kind in ('changed_delivery','wrong_date','amount_correction'):
            with self.subTest(kind=kind),tempfile.TemporaryDirectory() as temp:
                rec,ledger=self.existing()
                if kind=='changed_delivery':rec.update(deliver_local=120,so_delivery_local=120,sod_delivery_local={'SOD_TEST':120})
                if kind=='wrong_date':ledger.row_snapshot[3]['shoukuan_time']='2026-08-04'
                if kind=='amount_correction':
                    rec.update(amount_orig=19.7,amount_local=19.7,cumulative_received_local=19.7,deliver_local=20,so_delivery_local=20,sod_delivery_local={'SOD_TEST':20})
                    ledger.row_snapshot.pop(2);ledger.so_index['SO_TEST']=[3];ledger.sod_index['SOD_TEST']=[3]
                    ledger.row_snapshot[3]['yingshou']=50;rec['receivable_group_scope']['baseline_receivable']=50
                src=Path(temp)/'before.xlsx';out=Path(temp)/'after.xlsx'
                b=openpyxl.Workbook();w=b.active;w.title='明细'
                w.append(['新智云单号','实收金额','应收金额','计提','回款明细','是否结账','收款时间','收款方式(支/汇/现)','差异'])
                for row in ledger.row_snapshot.values():w.append([row.get(k) for k in ('so','sod','yingshou','jiti','huikuan','jiezhang','shoukuan_time','shoukuan_way','chayi')])
                b.save(src);b.close();original=src.read_bytes()
                ledger=C.LedgerIndex(src)
                rec['so_receipt_source']={'amount_orig':rec['amount_orig'],'amount_local':rec['amount_local'],'delivery_local':rec['deliver_local'],'cumulative_local':rec['cumulative_received_local'],'all_sods':rec['all_sods'],'sod_delivery_local':rec['sod_delivery_local'],'currency':'CNY','writeoff_sequence_key':rec['writeoff_sequence_key']}
                plan=C.classify_records([rec],ledger,{})
                checked=V.validate(plan,writer.read_ledger_rows(src))
                self.assertEqual(len(checked['write']),1,checked)
                item=checked['write'][0]
                writer.write_plan(src,out,[item]);self.assertEqual(writer.verify_written(out,[item]),[])
                rows=writer.read_ledger_rows(out)
                self.assertEqual(sum(r['应收金额'] or 0 for r in rows.values()),50 if kind=='amount_correction' else 100)
                self.assertAlmostEqual(sum(r['回款明细'] or 0 for r in rows.values()),19.7 if kind=='amount_correction' else 20)
                self.assertEqual(V.validate(plan,rows)['counts'],{'write':0,'skip':1,'conflict':0})
                self.assertEqual(src.read_bytes(),original)
                journal=BR.merge_journal({},checked);BR.validate_journal(journal)
                after=C.LedgerIndex(out);after.baseline_receipt_state=journal
                second=C.classify_records([rec],after,{})
                self.assertEqual(V.validate(second,rows)['counts'],{'write':0,'skip':1,'conflict':0})


    def test_single_order_uses_actual_parent_including_cents(self):
        payment=parent(39.8);payment['orders']=payment['orders'][:1]
        C.reconcile_writeoff_details([payment],{'AR_TEST':payment},[],dt.date(2026,8,5))
        records=C.expand_payments([payment],{})
        self.assertEqual(sum(r['amount_local'] for r in records),39.8)
        self.assertTrue(records[0]['writeoff_sequence_key'])
        self.assertEqual(V._whole_parent_gate_error(payment['duplicate_writeoff_audit']),'')

    def test_two_existing_partial_sods_do_not_create_accrual(self):
        rec,ledger=self.existing(amount=20,prior=20,arrival='2026-08-04',paid_date='2026-08-05')
        rec.pop('receivable_group_scope')
        ledger.row_snapshot[2]['yingshou']=30
        ledger.row_snapshot[4]={'so':'SO_TEST','sod':'SOD_OTHER','yingshou':30,'huikuan':None,'jiezhang':'否'}
        ledger.row_snapshot[5]={'so':'SO_TEST','sod':'SOD_OTHER','yingshou':20,'huikuan':20,'jiezhang':'是','shoukuan_time':'2026-08-05','shoukuan_way':'汇'}
        ledger.so_index['SO_TEST'] += [4,5];ledger.sod_index['SOD_OTHER']=[4,5]
        rec.update(forced_code='E5',default_first_sod=True,default_amount_local=40,default_amount_orig=40,default_cumulative_received_local=40,default_sod_lines=[{'sod':'SOD_TEST','deliver_local':50},{'sod':'SOD_OTHER','deliver_local':50}],all_sods=['SOD_TEST','SOD_OTHER'],sod_delivery_local={'SOD_TEST':50,'SOD_OTHER':50},so_delivery_local=100)
        plan=C.classify_records([rec],ledger,{})
        self.assertEqual(len(plan['auto']),2,plan)
        self.assertFalse(any(r.get('so_accrual_backfills') or r['five_cols'].get('计提') for r in plan['auto']))
        checked=V.validate(plan,ledger_rows(ledger));self.assertEqual(checked['counts'],{'write':2,'skip':0,'conflict':0},checked)
        after=ledger_rows(ledger)
        for item in plan['auto']:after[item['ledger_row_ref']].update(item['five_cols'])
        self.assertEqual(V.validate(plan,after)['counts'],{'write':0,'skip':2,'conflict':0})

    def test_missing_detail_parent_gate_rejects_tampered_capacity(self):
        payment=parent()
        C.reconcile_writeoff_details([payment],{'AR_TEST':payment},[],dt.date(2026,8,5))
        audit=copy.deepcopy(payment['duplicate_writeoff_audit']);audit['order_records'][0]['capacity']=999
        self.assertTrue(V._whole_parent_gate_error(audit))

    def test_whole_parent_with_real_details_still_checks_actual_overage(self):
        payment=parent(70)
        raw=[{'record_id':'HX_A','rowid':'ROW_A','ar':'AR_TEST','date':dt.date(2026,8,5),'so':'SO_A','amount':90,'amount_local':90,'currency':'CNY'}]
        C.reconcile_writeoff_details([payment],{'AR_TEST':payment},raw,dt.date(2026,8,5))
        self.assertTrue(payment.get('_parent_audit_unresolved'))


    def test_zero_receipt_proof_cannot_hide_positive_delivery_or_writes(self):
        import receipt_history as H
        row={'SO':'SO_TEST','SOD':'SOD_TEST','应收金额':0,'计提':None,'回款明细':None,'差异':None,'是否结账':'是','收款时间':None,'收款方式':''}
        item={'code':H.ZERO_CODE,'so':'SO_TEST','sod':'SOD_TEST','split_payment_source':{'amount_local':0,'delivery_local':0},'zero_delivery_audit':{'before_rows':{'2':row}}}
        self.assertEqual(V.check_one(item,{2:row})['verdict'],'skip')
        bad=copy.deepcopy(item);bad['split_payment_source']['delivery_local']=10
        self.assertEqual(V.check_one(bad,{2:row})['verdict'],'conflict')
        bad=copy.deepcopy(item);bad['so_accrual_backfills']=[{'accrual':10}]
        self.assertEqual(V.check_one(bad,{2:row})['verdict'],'conflict')

    def test_invalid_parent_capacity_is_a_conflict_not_an_exception(self):
        payment=parent()
        C.reconcile_writeoff_details([payment],{'AR_TEST':payment},[],dt.date(2026,8,5))
        audit=copy.deepcopy(payment['duplicate_writeoff_audit']);audit['order_records'][0]['capacity']='not a number'
        self.assertTrue(V._whole_parent_gate_error(audit))
