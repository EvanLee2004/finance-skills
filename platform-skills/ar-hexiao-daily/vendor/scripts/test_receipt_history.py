import unittest
import copy
from classify_hexiao import LedgerIndex, classify_one
import baseline_receipts as BR
import validate_plan as V

class HistoryTest(unittest.TestCase):
    def make(self, amount=20, prior=20, arrival='2026-07-10', paid_date='2026-07-10'):
        rows={2:{'so':'SO_TEST','sod':'SOD_TEST','yingshou':100-prior,'jiti':None,'huikuan':None,'jiezhang':'否'},
              3:{'so':'SO_TEST','sod':'SOD_TEST','yingshou':prior,'jiti':None,'huikuan':prior,'jiezhang':'是','shoukuan_time':paid_date,'shoukuan_way':'汇'}}
        ledger=LedgerIndex(synthetic={'so':{'SO_TEST':[2,3]},'sod':{'SOD_TEST':[2,3]},'rows':rows})
        rec={'ar':'AR_TEST','so':'SO_TEST','sod':'SOD_TEST','amount_local':amount,'amount_orig':amount,'deliver_local':100,'so_delivery_local':100,'sod_delivery_local':{'SOD_TEST':100},'deliver_orig':100,'currency':'CNY','cumulative_received_local':amount,'shoukuan_date':arrival,'hexiao_date':'2026-08-05','writeoff_sequence_key':['2026-08-05','HX_TEST','DETAIL_TEST'],'status':'正常','all_sods':['SOD_TEST'],'receivable_group_scope':{'basis':'so_latest_delivery','so':'SO_TEST','ledger_sod':'SOD_TEST','source_sods':['SOD_TEST'],'baseline_receivable':100}}
        return rec,ledger
    def test_existing_receipt_updates_without_splitting(self):
        rec,ledger=self.make()
        item=classify_one(rec,ledger,{},.01,2026)
        self.assertEqual(item['bucket'],'auto')
        self.assertFalse(item.get('row_operation'))
        self.assertEqual(item['ledger_row_ref'],3)
        self.assertEqual(item['five_cols']['收款时间'],'2026-08-05')
        rows={int(k):v for k,v in BR.ledger_rows(ledger,'SO_TEST','SOD_TEST').items()}
        self.assertEqual(V.check_one(item,rows)['verdict'],'write')
    def test_earlier_missing_receipt_preserves_future_receipt(self):
        rec,ledger=self.make(amount=10,prior=20,arrival='2026-08-04',paid_date='2026-08-28')
        item=classify_one(rec,ledger,{},.01,2026)
        self.assertEqual(item['bucket'],'auto')
        op=item['row_operation']
        self.assertEqual((op['source_receivable'],op['paid_receivable'],op['unpaid_receivable']),(80,10,70))
        self.assertEqual(item['split_payment_source']['cumulative_local'],10)
        rows={int(k):v for k,v in BR.ledger_rows(ledger,'SO_TEST','SOD_TEST').items()}
        self.assertEqual(V.check_one(item,rows)['verdict'],'write')

    def test_real_write_readback_and_repeat_preserve_other_receipts(self):
        import tempfile
        from pathlib import Path
        import openpyxl
        import apply_to_copy as writer
        for missing in (False,True):
            with self.subTest(missing=missing), tempfile.TemporaryDirectory() as temp:
                rec,ledger=self.make(**({'amount':10,'prior':20,'arrival':'2026-08-04','paid_date':'2026-08-28'} if missing else {}))
                src=Path(temp)/'before.xlsx';out=Path(temp)/'after.xlsx'
                b=openpyxl.Workbook();w=b.active;w.title='明细'
                w.append(['新智云单号','实收金额','应收金额','计提','回款明细','是否结账','收款时间','收款方式(支/汇/现)','差异'])
                for row in ledger.row_snapshot.values():w.append([row.get(k) for k in ('so','sod','yingshou','jiti','huikuan','jiezhang','shoukuan_time','shoukuan_way','chayi')])
                b.create_sheet('保留')['A1']='unchanged';b.save(src);b.close();original=src.read_bytes()
                item=classify_one(rec,ledger,{},.01,2026)
                self.assertEqual(V.check_one(item,writer.read_ledger_rows(src))['verdict'],'write')
                writer.write_plan(src,out,[item]);self.assertEqual(writer.verify_written(out,[item]),[])
                rows=writer.read_ledger_rows(out)
                self.assertEqual(V.check_one(item,rows)['verdict'],'skip')
                self.assertEqual(sum(r['应收金额'] or 0 for r in rows.values()),100)
                self.assertEqual(sum(r['回款明细'] or 0 for r in rows.values()),30 if missing else 20)
                if missing:
                    future=[r for r in rows.values() if r['回款明细']==20][0]
                    self.assertEqual(str(future['收款时间'])[:10],'2026-08-28')
                self.assertEqual(src.read_bytes(),original)
                journal=BR.merge_journal({}, {'write':[item]});BR.validate_journal(journal)
                aliases=dict(zip(BR.FIELDS,('so','sod','yingshou','jiti','huikuan','jiezhang','shoukuan_time','shoukuan_way','chayi')))
                raw={ref:{aliases[k]:r.get(k) for k in aliases} for ref,r in rows.items()}
                again=LedgerIndex(synthetic={'so':{'SO_TEST':list(raw)},'sod':{'SOD_TEST':list(raw)},'rows':raw});again.baseline_receipt_state=journal
                second=classify_one(rec,again,{},.01,2026)
                self.assertEqual(V.check_one(second,rows)['verdict'],'skip',second)

    def test_ambiguous_receipts_and_changed_source_are_rejected(self):
        rec,ledger=self.make()
        item=classify_one(rec,ledger,{},.01,2026)
        rows={int(k):v for k,v in BR.ledger_rows(ledger,'SO_TEST','SOD_TEST').items()}
        changed=copy.deepcopy(item);changed['split_payment_source']['amount_local']=21
        self.assertEqual(V.check_one(changed,rows)['verdict'],'conflict')
        rows[2]['应收金额']=79
        self.assertEqual(V.check_one(item,rows)['verdict'],'conflict')
        import receipt_history as H
        before=BR.ledger_rows(ledger,'SO_TEST','SOD_TEST');before['4']=copy.deepcopy(before['3']);before['2']['应收金额']=60
        self.assertIsNone(H.plan(rec,item,before,{}))

    def test_registered_receipt_cannot_be_claimed_by_another_event(self):
        rec,ledger=self.make();item=classify_one(rec,ledger,{},.01,2026)
        journal=BR.merge_journal({}, {'write':[item]})
        ledger.row_snapshot[3].update(shoukuan_time='2026-08-05',shoukuan_way='冲预收');ledger.baseline_receipt_state=journal
        import receipt_history as H
        other={**rec,'ar':'AR_OTHER','writeoff_sequence_key':['2026-08-05','HX_OTHER','DETAIL_OTHER']}
        self.assertIsNone(H.candidate(other,item,ledger))

    def test_batch_classifier_keeps_existing_update_and_missing_split(self):
        from classify_hexiao import classify_records
        for kwargs in ({},{'amount':10,'prior':20,'arrival':'2026-08-04','paid_date':'2026-08-28'}):
            rec,ledger=self.make(**kwargs)
            classified=classify_records([rec],ledger,{})
            self.assertEqual(len(classified['auto']),1,classified)
            item=classified['auto'][0]
            rows={int(k):v for k,v in BR.ledger_rows(ledger,'SO_TEST','SOD_TEST').items()}
            self.assertEqual(V.check_one(item,rows)['verdict'],'write',item)
