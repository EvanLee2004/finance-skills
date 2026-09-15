"""Verify the report toolbar refresh and the data returned by Kingdee."""
import json
from urllib.parse import urlsplit, parse_qs
from playwright.sync_api import expect, TimeoutError
from engine import SourceError, amount, label_norm, line_id, worksheets

REFRESH = '#toolbarap #fresh[data-btn-key="fresh"]'
TIMEOUT = 45000

def request_matches(response, page_id, action, period=None):
    request = response.request
    url = urlsplit(response.url)
    query = parse_qs(url.query)
    if (url.hostname != "tf.jdy.com" or url.path != "/ierp/form/batchInvokeAction.do"
            or query.get("ac") != [action] or request.method != "POST"):
        return False
    fields = parse_qs(request.post_data or "")
    if fields.get("pageId") != [page_id]:
        return False
    try:
        calls = json.loads(fields.get("params", [""])[0])
    except (ValueError, TypeError):
        return False
    if not isinstance(calls, list):
        return False
    for call in calls:
        if not isinstance(call, dict):
            continue
        if action == "fresh" and call.get("key") == "toolbarap" and call.get("methodName") == "itemClick" and call.get("args", [])[:1] == ["fresh"]:
            return True
        if action == "updateValue" and call.get("methodName") == "updateValue":
            post = call.get("postData", [])
            if len(post) == 2 and isinstance(post[1], list):
                if any(isinstance(v,dict) and v.get("k") == "period" and v.get("v") == period for v in post[1]):
                    return True
    return False

def response_actions(response):
    if response.status != 200:
        raise SourceError("报表请求失败，停止导出")
    response.finished()
    try:
        actions = response.json()
    except (ValueError, TypeError):
        raise SourceError("报表响应无法核验，停止导出") from None
    if not isinstance(actions, list) or any(not isinstance(a,dict) or a.get("a") not in {"setEnable", "u", "InvokeControlMethod", "setValue", "setVisible", "setFocus"} for a in actions):
        raise SourceError("报表响应包含未确认状态，停止导出")
    return actions

def refresh_control(page):
    button = page.locator(REFRESH).filter(visible=True)
    expect(button).to_have_count(1)
    token = button.get_attribute("data-page-id") or ""
    if not token.endswith("_fresh"):
        raise SourceError("报表刷新按钮身份无法核验")
    return button, token[:-6]

def select_period(page, selected, target, period):
    expected = period[:4] + "年" + period[4:] + "期"
    if selected.input_value() == expected:
        return
    _, page_id = refresh_control(page)
    try:
        with page.expect_response(lambda r: request_matches(r,page_id,"updateValue",period), timeout=TIMEOUT) as pending:
            selected.click()
            target.click()
        actions = response_actions(pending.value)
        if not any(isinstance(p,dict) and p.get("k") == "period" and p.get("v") == period
                   for a in actions for p in a.get("p",[])):
            raise SourceError("服务端未确认目标会计期间，停止导出")
        expect(selected).to_have_value(expected)
    except TimeoutError:
        raise SourceError("会计期间尚未同步完成，停止导出") from None

def balance_rows(actions):
    grids = []
    fields = ("assetname","assetindex","assetend","assetyear","debitname","debitindex","debitend","debityear")
    for action in actions:
        for part in action.get("p", []):
            if not isinstance(part,dict):
                continue
            data = part.get("data", {})
            if not isinstance(data,dict):
                continue
            index = data.get("dataindex", {})
            if isinstance(index,dict) and all(k in index for k in fields) and isinstance(data.get("rows"),list):
                try:
                    grids.append([[row[index[k]] for k in fields] for row in data["rows"]])
                except (IndexError,KeyError,TypeError):
                    raise SourceError("资产负债表刷新数据结构异常") from None
    if len(grids) != 1 or not grids[0]:
        raise SourceError("未收到完整资产负债表刷新数据，停止导出")
    return grids[0]

def balance_values(rows):
    values = {}
    for row in rows:
        if len(row) < 8:
            continue
        for offset in (0,4):
            if line_id(row[offset+1]) is None:
                continue
            key = (label_norm(row[offset]), line_id(row[offset+1]))
            if key in values:
                raise SourceError("资产负债表刷新行重复")
            values[key] = (amount(row[offset+2])[0], amount(row[offset+3])[0])
    if not values:
        raise SourceError("资产负债表刷新数据为空")
    return values

def verify_visible_balance(page, rows):
    expected = balance_values(rows)
    # Kingdee virtualizes rows; compare every currently rendered data row.
    cells = page.locator("table:visible tbody tr").evaluate_all(
        "(rs)=>rs.map(r=>Array.from(r.querySelectorAll('td')).map(c=>c.innerText.trim()))")
    actual = balance_values(cells)
    if any(key not in expected or values != expected[key] for key,values in actual.items()):
        raise SourceError("页面金额与本次刷新响应不一致，停止导出")

def refresh_report(page, selected, period, kind):
    expected = period[:4] + "年" + period[4:] + "期"
    expect(selected).to_have_value(expected)
    button,page_id = refresh_control(page)
    try:
        with page.expect_response(lambda r: request_matches(r,page_id,"fresh"),timeout=TIMEOUT) as pending:
            button.click()
        actions = response_actions(pending.value)
        expect(selected).to_have_value(expected)
        # Completion is the matching toolbar response, never network idle alone.
        page.wait_for_load_state("networkidle", timeout=TIMEOUT)
        expect(selected).to_have_value(expected)
    except TimeoutError:
        raise SourceError("报表刷新未确认完成，停止导出") from None
    rows = balance_rows(actions) if kind == "bs" else None
    if rows is not None:
        verify_visible_balance(page,rows)
    return rows

def verify_balance_export(path, rows):
    sheets = list(worksheets(path))
    if len(sheets) != 1 or balance_values(sheets[0][1]) != balance_values(rows):
        raise SourceError("导出金额与本次刷新数据不一致，停止纳入底稿")

def select_profit_period(page, selected, target, period):
    """Profit report filters are client-side until the Query button is pressed."""
    expected = period[:4] + "年" + period[4:] + "期"
    changed = selected.input_value() != expected
    if changed:
        selected.click()
        target.click()
        expect(selected).to_have_value(expected)
    return changed

def profit_request_matches(response,page_id,period,changed):
    request=response.request
    url=urlsplit(response.url)
    if (url.hostname!="tf.jdy.com" or url.path!="/ierp/form/batchInvokeAction.do"
            or parse_qs(url.query).get("ac")!=["click"] or request.method!="POST"):
        return False
    data=parse_qs(request.post_data or "")
    if data.get("pageId") != [page_id]:return False
    try:calls=json.loads(data.get("params",[""])[0])
    except (ValueError,TypeError):return False
    for call in calls if isinstance(calls,list) else []:
        if not isinstance(call,dict) or call.get("key")!="query" or call.get("methodName")!="click":continue
        post=call.get("postData",[])
        updates=post[1] if len(post)==2 and isinstance(post[1],list) else []
        periods=[v.get("v") for v in updates if isinstance(v,dict) and v.get("k")=="period"]
        if periods and periods!=[period]:return False
        return not changed or periods==[period]
    return False

def profit_rows(actions):
    fields=("name","rownumber","currentamount1","yearamount")
    grids=[]
    for action in actions:
        for part in action.get("p",[]):
            data=part.get("data",{}) if isinstance(part,dict) else {}
            if not isinstance(data,dict):continue
            index=data.get("dataindex",{})
            if all(k in index for k in fields) and isinstance(data.get("rows"),list):
                try:grids.append([[row[index[k]] for k in fields] for row in data["rows"]])
                except (IndexError,KeyError,TypeError):
                    raise SourceError("利润表查询数据结构异常") from None
    if len(grids)!=1 or not grids[0]:raise SourceError("未收到完整利润表查询数据，停止导出")
    return grids[0]

def profit_values(rows,columns=(0,1,2,3)):
    result={}
    for row in rows:
        if len(row)<=max(columns):continue
        name,number,current,ytd=(row[c] for c in columns)
        if line_id(number) is None:continue
        key=(label_norm(name),line_id(number))
        if key in result:raise SourceError("利润表查询行重复")
        result[key]=(amount(current)[0],amount(ytd)[0])
    if not result:raise SourceError("利润表查询数据为空")
    return result

def verify_visible_profit(page,rows):
    headers=page.locator("table:visible th").evaluate_all("(es)=>es.map(e=>e.getAttribute('data-code'))")
    fields=("name","rownumber","currentamount1","yearamount")
    if any(headers.count(f)!=1 for f in fields):raise SourceError("利润表本月及本年累计列无法确认")
    columns=tuple(headers.index(f) for f in fields)
    cells=page.locator("table:visible tbody tr").evaluate_all("(rs)=>rs.map(r=>Array.from(r.querySelectorAll('td')).map(c=>c.innerText.trim()))")
    actual=profit_values(cells,columns);expected=profit_values(rows)
    if any(k not in expected or expected[k]!=v for k,v in actual.items()):
        raise SourceError("利润表页面金额与查询响应不一致，停止导出")

def query_profit(page,selected,period,changed):
    expected=period[:4]+"年"+period[4:]+"期"
    expect(selected).to_have_value(expected)
    button=page.locator('[data-btn-key="query"]').filter(visible=True)
    expect(button).to_have_count(1)
    token=button.get_attribute("data-page-id") or ""
    if not token.endswith("_query"):raise SourceError("利润表查询按钮身份无法核验")
    try:
        with page.expect_response(lambda r:profit_request_matches(r,token[:-6],period,changed),timeout=TIMEOUT) as pending:
            button.click()
        actions=response_actions(pending.value)
        periods=[p.get("v") for a in actions for p in a.get("p",[]) if isinstance(p,dict) and p.get("k")=="period"]
        if (changed and periods!=[period]) or (periods and periods!=[period]):
            raise SourceError("利润表查询未确认目标会计期间")
        expect(selected).to_have_value(expected)
    except TimeoutError:
        raise SourceError("利润表查询未确认完成，停止导出") from None
    rows=profit_rows(actions)
    import time
    deadline=time.monotonic()+TIMEOUT/1000
    while True:
        try:
            verify_visible_profit(page,rows)
            break
        except SourceError:
            if time.monotonic()>=deadline:raise
            page.wait_for_timeout(100)
    return rows

def verify_profit_export(path,rows):
    sheets=list(worksheets(path))
    if len(sheets)!=1:raise SourceError("利润表导出工作表数量不符")
    source=sheets[0][1]
    candidates=[]
    for row in source[:10]:
        labels=[label_norm(v) for v in row]
        if all(labels.count(k)==1 for k in ("项目","行次","本月金额","本年累计金额")):
            candidates.append(tuple(labels.index(k) for k in ("项目","行次","本月金额","本年累计金额")))
    if len(candidates)!=1 or profit_values(source,candidates[0])!=profit_values(rows):
        raise SourceError("利润表导出金额与本次查询数据不一致，停止纳入底稿")
