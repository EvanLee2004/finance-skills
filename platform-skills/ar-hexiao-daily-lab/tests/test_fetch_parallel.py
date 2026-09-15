import sys,threading,time,unittest
from unittest.mock import patch
import importlib.util
from pathlib import Path
spec = importlib.util.spec_from_file_location(
    'parallel_fetch_under_test',
    Path(__file__).resolve().parents[1] / 'vendor/scripts/fetch_zhiyun.py',
)
F = importlib.util.module_from_spec(spec)
spec.loader.exec_module(F)

class Tests(unittest.TestCase):
 def test_bounded_isolated_sessions_and_order(self):
  client=F.ZhiyunClient('http://example.invalid','fixture');active=0;peak=0;lock=threading.Lock();sessions=set();barrier=threading.Barrier(F.FETCH_READ_CONCURRENCY)
  def search(child,ws,key):
   nonlocal active,peak
   with lock:
    active+=1;peak=max(peak,active);sessions.add(id(child.session))
   assert child is not client and child.headers==client.headers and child.headers is not client.headers
   barrier.wait(timeout=3);time.sleep(0.01 if key=='a' else 0.005)
   with lock:active-=1
   return [{'so':key}]
  with patch.object(F.ZhiyunClient,'search_rows',search):client.prefetch_search_rows('ws',['a','b','a','c','d'])
  self.assertEqual(peak,F.FETCH_READ_CONCURRENCY);self.assertEqual(len(sessions),F.FETCH_READ_CONCURRENCY)
  self.assertEqual([client.search_rows('ws',k)[0]['so'] for k in ['a','b','c','d']],['a','b','c','d']);self.assertEqual(client._prefetched_search,{})
  client.session.close()
 def test_error_is_preserved_until_original_caller(self):
  c=F.ZhiyunClient('http://example.invalid','fixture')
  def search(child,ws,key):
   if key=='bad':raise F.FetchError('synthetic failure')
   return []
  with patch.object(F.ZhiyunClient,'search_rows',search):c.prefetch_search_rows('ws',['ok','bad'])
  self.assertEqual(c.search_rows('ws','ok'),[])
  with self.assertRaises(F.FetchError):c.search_rows('ws','bad')
  c.session.close()
 def test_pagination_remains_ordered(self):
  c=F.ZhiyunClient('http://example.invalid','fixture');seen={};lock=threading.Lock()
  def post(child,path,body,timeout=90):
   key=body['keyWords'];page=body['pageIndex']
   with lock:seen.setdefault(key,[]).append(page)
   return {'data':[{'id':str(i)} for i in range(200)] if page==1 else [{'id':'last'}]}
  with patch.object(F.ZhiyunClient,'post',post):c.prefetch_search_rows('ws',['a','b'])
  for k in ['a','b']:
   self.assertEqual(seen[k],[1,2]);self.assertEqual(len(c.search_rows('ws',k)),201)
  c.session.close()
 def test_single_query_and_serial_mode_keep_original_path(self):
  c=F.ZhiyunClient('http://example.invalid','fixture')
  with patch.object(F.ZhiyunClient,'search_rows',side_effect=AssertionError('unexpected')):
   c.prefetch_search_rows('ws',['a'])
   with patch.object(F,'FETCH_READ_CONCURRENCY',1):c.prefetch_search_rows('ws',['a','b'])
  self.assertEqual(c._prefetched_search,{});c.session.close()


class OutputTests(unittest.TestCase):
 def test_full_fetch_serial_parallel_equivalence(self):
  import copy,tempfile
  def controls(names):return [{'controlId':n,'controlName':n,'dataSource':n} for n in names]
  tables={
   F.WS_HUIKUAN:controls([F.REL_XIADAN,F.REL_JIESUAN,F.REL_HEXIAO_MINGXI,F.REL_SODLINE]),
   F.REL_XIADAN:controls(F.XIADAN_COLS),
   F.REL_HEXIAO_MINGXI:controls(F.MINGXI_COLS),
   F.REL_SODLINE:controls(F.SODLINE_COLS),
  }
  def post(client,path,b,timeout=90):
   ws=b['worksheetId']
   if path=='worksheet/getWorksheetInfo':return {'controls':tables[ws]}
   if path=='worksheet/getRowRelationRows':
    name=b['controlId'];rows=[]
    if name==F.REL_XIADAN:rows=[{'SO':'SO26000001','交付额/本币':'10'},{'SO':'SO26000002','交付额/本币':'20'}]
    return {'data':rows,'worksheet':{'controls':tables[name]}}
   assert path=='worksheet/getFilterRows'
   so=b['keyWords']
   if ws==F.WS_HUIKUAN:return {'data':[{'rowid':'ar',F.F_HK['ar']:'AR_TEST_001',F.F_HK['hexiao_date']:'2026-09-03'}],'count':1}
   if ws==F.REL_XIADAN:return {'data':[{'SO':so,'项目交付日期':'2026-09-01'}]}
   if ws==F.REL_SODLINE:return {'data':[{'SO':so,'SOD':so.replace('SO','SOD')+'01','交付额/原币':'10'}]}
   assert ws==F.REL_HEXIAO_MINGXI
   return {'data':[]}
  results=[]
  for n in [1,2,4]:
   c=F.ZhiyunClient('http://example.invalid','fixture');captured={}
   def publish(out,tag,datasets,summary):captured['datasets']=copy.deepcopy(datasets);return summary
   with tempfile.TemporaryDirectory() as out,patch.object(F,'FETCH_READ_CONCURRENCY',n),patch.object(F.ZhiyunClient,'post',post),patch.object(F,'publish_day_exports',publish):
    summary=F.fetch_day(c,'2026-09-03',__import__('pathlib').Path(out));results.append((captured,summary))
   c.session.close()
  self.assertEqual(results[0],results[1]);self.assertEqual(results[0],results[2])
  self.assertEqual(results[0][1]['涉及SO数'],2)

if __name__ == '__main__':
 unittest.main()
