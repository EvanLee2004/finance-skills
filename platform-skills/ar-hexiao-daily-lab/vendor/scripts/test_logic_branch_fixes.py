import copy
import unittest
import baseline_receipts as BR
import receipt_history as H
import flow_monthly as M
import validate_plan as V
from classify_hexiao import classify_one
import test_receipt_history as history_fixtures
import test_flow_monthly as flow_fixtures

class IdentityCorrections(unittest.TestCase):

    def test_flow_ambiguity_does_not_block_verified_receipt(self):
        rec, ledger = history_fixtures.HistoryTest().make()
        rec['flow_hits'] = 2
        item = classify_one(rec, ledger, {}, 0.01, 2026)
        self.assertEqual(item['bucket'], 'auto', item)
        rows = {int(k): r for k, r in BR.ledger_rows(ledger, rec['so'], rec['sod']).items()}
        self.assertEqual(V.check_one(item, rows)['verdict'], 'write')

    def test_owned_event_restores_signature_after_uploaded_table(self):
        rec, ledger = history_fixtures.HistoryTest().make()
        identity = BR.event_key(rec)
        ledger.baseline_receipt_state = {BR.group_key(rec['so'], rec['sod']): {'baseline_receivable': 100, 'scope_only': True, 'events': {}, 'ordinary_events': {identity: {'signature': [20, '2026-08-05', '冲预收']}}}}
        item = classify_one(rec, ledger, {}, 0.01, 2026)
        self.assertEqual(item.get('receipt_correction', {}).get('kind'), 'existing', item)
        self.assertEqual(item['ledger_row_ref'], 3)
        rows = {int(k): r for k, r in BR.ledger_rows(ledger, rec['so'], rec['sod']).items()}
        self.assertEqual(V.check_one(item, rows)['verdict'], 'write')

    def test_excel_float_tail_only(self):
        self.assertEqual(M.money(321.96000000000026), M.money('321.96'))
        with self.assertRaises(ValueError):
            M.money('321.961')
        with self.assertRaises(ValueError):
            M.money(321.961)

    def test_registered_receipt_rebind_keeps_settlement_guards(self):
        from classify_hexiao import LedgerIndex
        rec, _ = history_fixtures.HistoryTest().make(amount=30)
        rec.update(deliver_local=50,so_delivery_local=50,cumulative_received_local=50,sod_delivery_local={'SOD_TEST':50})
        rec.pop('receivable_group_scope')
        rows={2:{'so':'SO_TEST','sod':'SOD_TEST','yingshou':100,'huikuan':20,'jiezhang':'是','shoukuan_time':'2026-07-10','shoukuan_way':'汇'},
              3:{'so':'SO_TEST','sod':'SOD_TEST','yingshou':None,'huikuan':30,'jiti':50,'jiezhang':'是','shoukuan_time':'2026-08-05','shoukuan_way':'冲预收'}}
        ledger=LedgerIndex(synthetic={'so':{'SO_TEST':[2,3]},'sod':{'SOD_TEST':[2,3]},'rows':rows})
        old={**rec,'ar':'AR_OLD','writeoff_sequence_key':['2026-07-10','HX_OLD','detail']}
        events={}
        for r,slot,amount,date,way in [(old,8,20,'2026-07-10','汇'),(rec,9,30,'2026-08-05','冲预收')]:
            identity=BR.event_key(r)
            events[identity]={'event_key':identity,'ar':r['ar'],'so':r['so'],'sod':r['sod'],'case_id':r['ar'],'slot':slot,'回款明细':amount,'收款时间':date,'收款方式':way}
        key=BR.group_key('SO_TEST','SOD_TEST')
        ledger.baseline_receipt_state={key:{'baseline_receivable':100,'events':events,'settled':True,'accrual':50}}
        for repeat in range(2):
            item=classify_one(rec,ledger,{},.01,2026)
            self.assertEqual(item['code'],'OK_BASELINE_RECEIPT_APPLIED',item)
            actual={int(k):v for k,v in BR.ledger_rows(ledger,'SO_TEST','SOD_TEST').items()}
            self.assertEqual(V.check_one(item,actual)['verdict'],'skip')
            ledger.baseline_receipt_state=BR.merge_journal(ledger.baseline_receipt_state,{'skip':[item]})
            BR.validate_journal(ledger.baseline_receipt_state)
        for col,value in [('jiti',999),('jiezhang','否')]:
            changed=copy.deepcopy(ledger);changed.row_snapshot[3][col]=value
            result=classify_one(rec,changed,{},.01,2026)
            actual={int(k):v for k,v in BR.ledger_rows(changed,'SO_TEST','SOD_TEST').items()}
            self.assertEqual(V.check_one(result,actual)['verdict'],'conflict')
        unowned=copy.deepcopy(ledger);unowned.baseline_receipt_state[key]['events'].pop(BR.event_key(rec))
        result=classify_one(rec,unowned,{},.01,2026)
        self.assertEqual(result['baseline_receipt_audit']['disposition'],'conflict')


    def test_zero_difference_representation(self):
        rec,ledger=history_fixtures.HistoryTest().make(amount=100,prior=100,arrival='2026-08-05',paid_date='2026-08-05')
        rows=BR.ledger_rows(ledger,'SO_TEST','SOD_TEST');rows.pop('2')
        rows['3'].update(计提=100,收款方式='汇',差异=None)
        base={'ar':rec['ar'],'so':rec['so'],'sod':rec['sod']}
        self.assertNotIn('差异',H.plan(rec,base,rows,{})['derived_cols'])
        for value in [0,2]:
            rows['3']['差异']=value
            self.assertEqual(H.plan(rec,base,rows,{})['derived_cols'],{'差异':0.0})

    def test_zero_allocation(self):
        audit={'reused_successful_allocation':True,'applied_sos':['SO_PAID'],'applied_cases':{'p':{'so':'SO_PAID','amount_local':100}},'allocations':[{'so':'SO_PAID','allocated':100,'allocated_local':100,'historical_received_local':0},{'so':'SO_ZERO','allocated':0,'allocated_local':0,'historical_received_local':0}]}
        plan={'auto':[{'ar':'AR_TEST','code':'E5'}],'parent_fallback_allocations':{'AR_TEST':audit}}
        rows={2:{'SO':'SO_PAID','回款明细':100},3:{'SO':'SO_ZERO','回款明细':50}}
        self.assertEqual(V.parent_allocation_history_errors(plan,rows),{})
        audit['allocations'][1].update(allocated=10,allocated_local=10)
        self.assertIn('AR_TEST',V.parent_allocation_history_errors(plan,rows))

class FlowCorrections(unittest.TestCase):
    def setUp(self):
        self.f=flow_fixtures.MonthlySafetyTest();self.f.setUp()
    def tearDown(self):self.f.tearDown()

    def test_trailing_newline(self):
        f=self.f;f.modify(lambda ws:setattr(ws['E2'],'value','WX\n'))
        changes,errors=f.run_items('2026-08-01',[f.entry(amount=1000)])
        self.assertEqual(errors,[]);self.assertEqual(len(changes),1)
        self.assertEqual(f.read('E2'),'WX转8月')
        before=f.path.read_bytes()
        self.assertEqual(f.run_items('2026-08-01',[f.entry(amount=1000)]),([],[]))
        self.assertEqual(f.path.read_bytes(),before)

    def test_valid_and_invalid_equations(self):
        import openpyxl
        f=self.f
        for raw in ['还剩：9,000-1000=8,000-3000=5,000转','9000-1000=8000-3000=5000']:
            f.modify(lambda ws:setattr(ws['F2'],'value',raw))
            wb=openpyxl.load_workbook(f.path)
            try:month=M.legacy_month(wb['流水'],2,M.columns(wb['流水']))
            finally:wb.close()
            self.assertEqual([e['amount'] for e in month['entries']],['1000','3000'])
            self.assertEqual(month['remaining'],'5000')
        f.modify(lambda ws:setattr(ws['F2'],'value','预付：5000'))
        self.assertEqual(f.run_items('2026-07-28',[f.entry(amount=500)])[1],[])
        self.assertEqual(f.read('F2'),'9000-4000-500=4500')
        for raw in ['9000-1000=7000-3000=4000','9,00-100=800','9000-1000','9000-10000=-1000']:
            f.modify(lambda ws:setattr(ws['F2'],'value',raw));before=f.path.read_bytes()
            self.assertTrue(f.run_items('2026-07-28',[f.entry('SO2',500)])[1])
            self.assertEqual(f.path.read_bytes(),before)

    def test_unaffected_and_affected_formula(self):
        f=self.f
        def edit(ws):
            ws['H1']='=VLOOKUP(A1,$J$1:$K$2,2,0)'
            for row in range(3,8):ws.cell(row,1).value=''
        f.modify(edit);M.insertion_guard(f.path,'流水',7)
        f.modify(lambda ws:setattr(ws['H1'],'value','=SUM(C1:C9)'))
        with self.assertRaises(ValueError):M.insertion_guard(f.path,'流水',7)

    def test_prefill_is_not_counted(self):
        import execution_flow_stage as E,json
        f=self.f;flow,plan=f.plans('2026-07-28',[f.entry()])
        path=f.root/'04_产出'/'checked.json';path.write_text(json.dumps(plan),encoding='utf-8')
        (path.parent/'流转写入计划_校验后.json').write_text(json.dumps(flow),encoding='utf-8')
        phase=E.run(f.root,path)['phases']['prefill']
        self.assertIs(phase.get('applicable'),False)
        self.assertEqual(phase['eligible_count'],0)
        self.assertEqual(phase['unchanged_count'],0)

    def test_shared_formula_followers_are_not_exempted(self):
        import zipfile,re
        f=self.f
        def edit(ws):
            for r in range(1,13):ws.cell(r,8).value='=C'+str(r+1)
        f.modify(edit)
        with zipfile.ZipFile(f.path) as z:payload={name:z.read(name) for name in z.namelist()}
        name='xl/worksheets/sheet1.xml';xml=payload[name].decode()
        for r in range(1,13):
            formula='<f t="shared" si="0" ref="H1:H12">C2</f>' if r==1 else '<f t="shared" si="0"/>'
            xml=re.sub(r'(<c\b[^>]*r="H'+str(r)+r'"[^>]*>).*?(</c>)',lambda m:m[1]+formula+'<v>1</v>'+m[2],xml)
        payload[name]=xml.encode()
        with zipfile.ZipFile(f.path,'w',zipfile.ZIP_DEFLATED) as z:
            for name,data in payload.items():z.writestr(name,data)
        before=f.path.read_bytes()
        with self.assertRaises(ValueError):M.insertion_guard(f.path,'流水',7)
        self.assertEqual(f.path.read_bytes(),before)
