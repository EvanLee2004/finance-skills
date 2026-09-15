"""Read-only Kingdee collection through the platform's explicit network policy."""
from pathlib import Path
import json
from urllib.parse import urlsplit, urlunsplit
import os
import re
from playwright.sync_api import sync_playwright, TimeoutError, expect
from engine import COMPANIES, KINDS, SourceError, parse_file
from report_refresh import select_period, refresh_report, verify_balance_export, verify_visible_balance, select_profit_period, query_profit, verify_visible_profit, verify_profit_export

LEDGERS = ["HEAD", "CULTURE", "SHANGHAI", "HUNAN_BRANCH", "HUNAN_TECH"]

def set_classification(page, *, verify_only=False):
    for label,wanted in [("重分类",True),("应交税费重分类",False)]:
        text=page.get_by_text(label,exact=True).filter(visible=True)
        text.wait_for(state="visible")
        container=text.locator("xpath=ancestor::*[.//input[@type='checkbox']][1]")
        box=container.locator("input[type=checkbox]")
        if box.count()!=1:raise SourceError("重分类控件不能唯一识别")
        if box.is_checked()!=wanted and not verify_only:
            # The custom checkbox is controlled by its visible enclosing component.
            container.click()
        if box.is_checked()!=wanted:raise SourceError("重分类口径设置未生效")

def query_classified_balance(page, query):
    """A changed option requires a fresh query; two unstable queries stop export."""
    for attempt in range(2):
        set_classification(page)
        rows = query()
        try:
            set_classification(page, verify_only=True)
            return rows
        except SourceError:
            if attempt:
                raise SourceError("查询后重分类口径不稳定，停止导出") from None


def collect(period, output, credentials, existing, emit):
    output=Path(output);output.mkdir(parents=True,exist_ok=True)
    files=[];issues=[]
    if not credentials.get("account") or not credentials.get("password"):
        return [],[{"message":"金蝶账号尚未配置，请安全保存账号后重试或上传报表","status":"needs_auth"}]
    allowed={urlsplit(v).hostname for v in os.environ.get("FINANCIAL_NETWORK_TARGETS","").split(",") if v}
    proxy=os.environ.get("HTTPS_PROXY")
    if not proxy or not allowed:
        return [],[{"message":"平台采集网络尚未配置","status":"network_unavailable"}]
    def visible_matches(page,text):
        loc=page.get_by_text(text,exact=True).filter(visible=True)
        return [loc.nth(i) for i in range(loc.count())]
    def click_visible(page,text,last=False):
        loc=page.get_by_text(text,exact=True).filter(visible=True)
        if last:
            loc.last.wait_for(state="visible")
            loc.last.click()
        else:
            expect(loc).to_have_count(1)
            loc.click()
    def current_company(page):
        for code in LEDGERS:
            if len(visible_matches(page,COMPANIES[code][1]))==1:return code
        raise SourceError("当前账套无法唯一识别")
    def period_input(page):
        loc=page.locator("input[readonly]")
        choices=[loc.nth(i) for i in range(loc.count()) if re.fullmatch(r"20\d{2}年\d{1,2}期",loc.nth(i).input_value()) and loc.nth(i).is_visible()]
        if len(choices)!=1:raise SourceError("期间选择器无法唯一识别")
        return choices[0]
    def dismiss_cash_notice(page):
        notice=page.get_by_text("现金流量表平衡检查",exact=True).filter(visible=True)
        if notice.count():
            page.get_by_text("我知道了",exact=True).filter(visible=True).click()
            notice.wait_for(state="hidden")

    def select_cumulative(page):
        headers=page.locator(".kd-table-header-title")
        if "本年累计金额" in headers.all_text_contents():
            return
        dropdown=page.locator('[title*="；"]').filter(visible=True)
        if dropdown.count()!=1:
            raise SourceError("数据类型选择器无法唯一识别")
        dropdown.click()
        option=page.get_by_text("本年累计金额",exact=True).filter(visible=True)
        option.last.wait_for(state="visible")
        option.last.click()
        page.get_by_text("会计期间",exact=True).filter(visible=True).click()

    with sync_playwright() as p:
        browser=p.chromium.launch(headless=True,proxy={"server":proxy})
        try:
            ctx=browser.new_context(accept_downloads=True,locale="zh-CN",viewport={"width":1440,"height":1000})
            def restrict_request(route):
                requested=urlsplit(route.request.url)
                # Kingdee's account switch can return an HTTP index redirect.
                # Upgrade that exact read-only navigation before any network request.
                if (requested.scheme=="http" and requested.hostname=="tf.jdy.com"
                    and requested.path=="/ierp/index.html" and route.request.method=="GET"
                    and route.request.is_navigation_request()):
                    secure=urlunsplit(("https","tf.jdy.com",requested.path,requested.query,requested.fragment))
                    route.fulfill(status=307,headers={"Location":secure})
                elif requested.scheme=="https" and requested.hostname in allowed:
                    route.continue_()
                else:route.abort()
            ctx.route("**/*",restrict_request)
            page=ctx.new_page();page.set_default_timeout(15000)
            try:
                page.goto("https://www.jdy.com/login/",wait_until="domcontentloaded",timeout=45000)
                page.get_by_role("textbox",name="账号",exact=True).fill(credentials.pop("account"))
                page.get_by_role("textbox",name="请输入密码",exact=True).fill(credentials.pop("password"))
                credentials.clear()
                page.get_by_role("button",name="登录",exact=True).click()
                try:page.wait_for_url("**/workbench/web/**",timeout=8000)
                except TimeoutError:
                    consent=page.get_by_text("同意",exact=True)
                    if consent.count()==1 and consent.is_visible():consent.click()
                    page.wait_for_url("**/workbench/web/**",timeout=25000)
                page.get_by_role("button",name="进入使用",exact=True).click()
                page.wait_for_url("https://tf.jdy.com/**",timeout=30000)
                page.get_by_text("财务报表",exact=True).wait_for(state="visible",timeout=30000)
            except Exception:
                return [],[{"message":"金蝶登录未完成，可能需重新授权或人工验证；可改用上传报表","status":"needs_auth"}]
            first_report=True
            for code in LEDGERS:
                for kind,title in KINDS.items():
                    if (code,kind) in existing:continue
                    emit("正在获取"+COMPANIES[code][0]+title,45)
                    for attempt in range(2):
                        stage="选择账套"
                        try:
                            if not first_report:
                                page.goto("https://tf.jdy.com/ierp/index.html",wait_until="domcontentloaded",timeout=45000)
                                page.get_by_text("财务报表",exact=True).wait_for(state="visible",timeout=30000)
                            first_report=False
                            current=current_company(page)
                            if current!=code:
                                click_visible(page,COMPANIES[current][1])
                                with page.expect_navigation(wait_until="domcontentloaded",timeout=45000):
                                    click_visible(page,COMPANIES[code][1])
                                destination=urlsplit(page.url)
                                if (destination.scheme=="http" and destination.hostname=="tf.jdy.com"
                                    and destination.path=="/ierp/index.html"):
                                    page.goto(urlunsplit(("https","tf.jdy.com",destination.path,destination.query,destination.fragment)),wait_until="domcontentloaded",timeout=45000)

                                expect(page.get_by_text(COMPANIES[current][1],exact=True).filter(visible=True)).to_have_count(0,timeout=30000)
                                expect(page.get_by_text(COMPANIES[code][1],exact=True).filter(visible=True)).to_have_count(1,timeout=60000)
                                page.get_by_text("财务报表",exact=True).wait_for(state="visible",timeout=30000)
                            if current_company(page)!=code:raise SourceError("账套不匹配")
                            stage="打开报表"
                            announcement=page.get_by_text("系统更新公告",exact=True).filter(visible=True)
                            if announcement.count():
                                page.keyboard.press("Escape")
                                if announcement.is_visible():
                                    raise SourceError("金蝶系统公告遮挡页面，请在金蝶关闭公告后重新取数")

                            click_visible(page,"财务报表")
                            click_visible(page,title,last=True)
                            page.locator("input[readonly]:visible").first.wait_for(state="visible")
                            try:page.wait_for_load_state("networkidle",timeout=5000)
                            except TimeoutError:pass
                            if kind=="cf":dismiss_cash_notice(page)
                            selected=period_input(page)
                            stage="选择会计期间"
                            if not selected.input_value().startswith(period[:4]+"年"):
                                raise SourceError("会计年需人工选择")
                            target=page.get_by_text(str(int(period[4:]))+"期",exact=True).filter(visible=True)
                            if kind=="is":
                                profit_period_changed=select_profit_period(page,selected,target,period)
                            else:
                                select_period(page, selected, target, period)
                            expected=period[:4]+"年"+period[4:]+"期"
                            try:expect(selected).to_have_value(expected,timeout=3000)
                            except AssertionError:
                                raise SourceError(str(int(period[4:]))+"期无法选择，已跳过；可上传本期报表") from None
                            if kind == "is":
                                stage="选择本年累计金额"
                                select_cumulative(page)
                            stage="查询并导出"
                            def query_report():
                                return refresh_report(page, selected, period, kind)
                            if kind=="bs":
                                stage="查询后复核重分类口径"
                                refreshed_rows = query_classified_balance(page, query_report)
                            elif kind=="is":
                                refreshed_rows=query_profit(page,selected,period,profit_period_changed)
                            else:
                                query_report()
                            if kind=="cf":dismiss_cash_notice(page)
                            if kind == "is":
                                expect(page.locator(".kd-table-header-title").get_by_text("本年累计金额",exact=True)).to_be_visible(timeout=30000)
                            if kind=="bs":
                                stage="导出前复核重分类口径"
                                set_classification(page, verify_only=True)
                            if kind=="bs":verify_visible_balance(page,refreshed_rows)
                            if kind=="is":verify_visible_profit(page,refreshed_rows)
                            click_visible(page,"引出")
                            dialog=page.locator("#dialogShow").filter(visible=True)
                            dialog.get_by_text("引出报表",exact=True).wait_for(state="visible")
                            expect(dialog.locator("input[readonly]").filter(visible=True)).to_have_value(expected,timeout=30000)
                            periods=page.locator("input[readonly]:visible").evaluate_all("(es)=>es.map(e=>e.value).filter(v=>/^20\\d{2}年\\d{1,2}期$/.test(v))")
                            if not periods or any(v!=expected for v in periods):raise SourceError("导出期间不匹配")
                            if kind=="bs":
                                set_classification(page, verify_only=True)
                            with page.expect_download(timeout=40000) as pending:
                                dialog.get_by_text("引出",exact=True).click()
                            download=pending.value
                            name=download.suggested_filename
                            if name!=COMPANIES[code][1]+"_"+title+"__"+period+"期.xlsx":
                                download.delete();raise SourceError("下载文件公司或期间不符")
                            path=output/name
                            download.save_as(path);download.delete()
                            if kind in {"bs","is"}:
                                stage="核对导出与刷新数据"
                                try:
                                    (verify_balance_export if kind=="bs" else verify_profit_export)(path,refreshed_rows)
                                except SourceError:
                                    path.rename(path.with_stem(path.stem+"_核验失败"))
                                    raise
                            parsed,problems=parse_file(path,period,bs_reclassified=kind=="bs")
                            if len(parsed)!=1 or parsed[0].company!=code or parsed[0].kind!=kind or problems:
                                # Keep the original export as evidence, but never mark it normalized.
                                issues.append({"company":code,"kind":kind,"message":"已导出，但字段或金额列口径待核实","status":"needs_review"})
                            if kind in {"bs","is"}:
                                import hashlib
                                evidence=output.parent/"取数诊断"
                                evidence.mkdir(exist_ok=True)
                                proof={"company":code,"kind":kind,"period":period,"refresh_action":"query.click" if kind=="is" else "toolbarap.fresh","verification":"response_page_export_equal","reclassified":kind=="bs","tax_reclassified":False,"source_sha256":hashlib.sha256(path.read_bytes()).hexdigest()}
                                (evidence/(code+"_"+kind+"_verified.json")).write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding="utf-8")
                            files.append(path)
                            break
                        except Exception as exc:
                            reason=str(exc) if isinstance(exc,SourceError) else "页面操作未完成"
                            diagnostic={"company":code,"kind":kind,"stage":stage,"requested_period":period,"error_type":type(exc).__name__}
                            try:
                                diagnostic["selected_period"]=period_input(page).input_value()
                            except Exception:
                                diagnostic["selected_period"]=""
                            evidence=output.parent/"取数诊断"
                            evidence.mkdir(exist_ok=True)
                            prefix=code+"_"+kind+"_attempt"+str(attempt+1)
                            try:
                                page.screenshot(path=str(evidence/(prefix+".png")),mask=[page.locator("input"),page.get_by_text(COMPANIES[code][1],exact=True)],timeout=5000)
                            except Exception:
                                pass
                            (evidence/(prefix+".json")).write_text(json.dumps(diagnostic,ensure_ascii=False,indent=2),encoding="utf-8")
                            if attempt == 0 and isinstance(exc,(TimeoutError,AssertionError)):
                                emit(COMPANIES[code][0]+title+"页面等待失败，重新打开后再试一次",45)
                                continue
                            issues.append({"company":code,"kind":kind,"message":stage+"："+reason,"status":"missing","stage":stage,"diagnostic":prefix+".json"})
                            break
            return files,issues
        finally:
            credentials.clear()
            browser.close()
