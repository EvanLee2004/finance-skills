# -*- coding: utf-8 -*-
"""fetch_zhiyun 纯函数单测（不连网、不碰账密）。"""
import json
import sys
import threading
from pathlib import Path

import pytest
import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import fetch_zhiyun as F


class _FakeResponse:
    def __init__(self, status_code, payload=None, headers=None):
        self.status_code = status_code
        self._payload = payload or {}
        self.headers = headers or {}
        self.url = "<REDACTED_URL>"

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(
                f"{self.status_code} synthetic response",
                response=self,
            )

    def json(self):
        return self._payload


class _SequenceSession:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.calls = 0

    def post(self, *args, **kwargs):
        del args, kwargs
        self.calls += 1
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


def test_read_only_post_retries_transient_502(monkeypatch):
    sleeps = []
    client = F.ZhiyunClient("http://127.0.0.1:1", "redacted")
    client.session = _SequenceSession(
        [
            _FakeResponse(502),
            _FakeResponse(200, {"data": {"rows": [1]}}),
        ]
    )
    monkeypatch.setattr(F.time, "sleep", sleeps.append)

    assert client.post("read-only-query", {}) == {"rows": [1]}
    assert client.session.calls == 2
    assert sleeps == [1.0]


def test_read_only_post_retries_transport_timeout_with_a_bound(monkeypatch):
    sleeps = []
    client = F.ZhiyunClient("http://127.0.0.1:1", "redacted")
    client.session = _SequenceSession(
        [
            requests.Timeout("synthetic timeout"),
            requests.Timeout("synthetic timeout"),
            requests.Timeout("synthetic timeout"),
            _FakeResponse(200, {"data": {"unexpected": True}}),
        ]
    )
    monkeypatch.setattr(F.time, "sleep", sleeps.append)

    with pytest.raises(requests.Timeout):
        client.post("read-only-query", {})

    assert client.session.calls == 3
    assert sleeps == [1.0, 2.0]


def test_read_only_post_does_not_retry_non_transient_4xx(monkeypatch):
    sleeps = []
    client = F.ZhiyunClient("http://127.0.0.1:1", "redacted")
    client.session = _SequenceSession([_FakeResponse(403)])
    monkeypatch.setattr(F.time, "sleep", sleeps.append)

    with pytest.raises(requests.HTTPError):
        client.post("read-only-query", {})

    assert client.session.calls == 1
    assert sleeps == []


def test_read_only_post_retries_unlisted_5xx(monkeypatch):
    sleeps = []
    client = F.ZhiyunClient("http://127.0.0.1:1", "redacted")
    client.session = _SequenceSession(
        [_FakeResponse(507), _FakeResponse(200, {"data": {"ok": True}})]
    )
    monkeypatch.setattr(F.time, "sleep", sleeps.append)

    assert client.post("read-only-query", {}) == {"ok": True}
    assert client.session.calls == 2
    assert sleeps == [1.0]


def test_resolve_date_yesterday():
    d = F.resolve_date("yesterday")
    assert len(d) == 10 and d[4] == "-" and d[7] == "-"


def test_resolve_date_fixed():
    assert F.resolve_date("2026-07-21") == "2026-07-21"


def test_existing_exports_require_current_schema_version(tmp_path):
    day = "2026-07-27"
    stamp = "20260727"
    for role in ("回款记录", "订单交付", "核销明细", "订单明细"):
        (tmp_path / f"{role}_{stamp}.xlsx").write_bytes(b"fixture")

    # 旧四件套没有版本摘要时，默认禁止跳过重新取数。
    assert F.already_fetched(tmp_path, day) == []
    assert len(F.already_fetched(tmp_path, day, accept_unversioned=True)) == 4

    (tmp_path / f"取数摘要_{stamp}.json").write_text(
        '{"export_schema_version":"old"}',
        encoding="utf-8",
    )
    assert F.already_fetched(tmp_path, day) == []

    files = [f"{role}_{stamp}.xlsx" for role in ("回款记录", "订单交付", "核销明细", "订单明细")]
    (tmp_path / f"取数摘要_{stamp}.json").write_text(
        json.dumps(
            {
                "export_schema_version": F.EXPORT_SCHEMA_VERSION,
                "file_sha256": {
                    name: F._sha256(tmp_path / name)
                    for name in files
                },
            }
        ),
        encoding="utf-8",
    )
    assert len(F.already_fetched(tmp_path, day)) == 4

    (tmp_path / files[0]).write_bytes(b"changed")
    assert F.already_fetched(tmp_path, day) == []


def test_daily_bundle_failure_does_not_replace_previous_complete_files(tmp_path, monkeypatch):
    day_tag = "20260817"
    names = [f"{role}_{day_tag}.xlsx" for role in ("回款记录", "订单交付", "核销明细", "订单明细")]
    for name in names:
        (tmp_path / name).write_bytes(b"old-complete")
    old_summary = tmp_path / f"取数摘要_{day_tag}.json"
    old_summary.write_text("old-summary", encoding="utf-8")
    calls = 0

    def fail_before_bundle_is_complete(path, _headers, _rows):
        nonlocal calls
        calls += 1
        if calls == 4:
            raise OSError("synthetic write failure")
        path.write_bytes(b"new-staged")

    monkeypatch.setattr(F, "write_xlsx", fail_before_bundle_is_complete)

    with pytest.raises(OSError, match="synthetic write failure"):
        F.publish_day_exports(
            tmp_path,
            day_tag,
            [(name, ["h"], [[1]]) for name in names],
            {"export_schema_version": F.EXPORT_SCHEMA_VERSION},
        )

    assert all((tmp_path / name).read_bytes() == b"old-complete" for name in names)
    assert old_summary.read_text(encoding="utf-8") == "old-summary"


def test_interrupted_publish_invalidates_mixed_daily_bundle(tmp_path, monkeypatch):
    day = "2026-08-17"
    day_tag = day.replace("-", "")
    names = [f"{role}_{day_tag}.xlsx" for role in ("回款记录", "订单交付", "核销明细", "订单明细")]
    for name in names:
        (tmp_path / name).write_bytes(b"old-complete")
    old_hashes = {name: F._sha256(tmp_path / name) for name in names}
    summary_path = tmp_path / f"取数摘要_{day_tag}.json"
    summary_path.write_text(
        json.dumps(
            {
                "export_schema_version": F.EXPORT_SCHEMA_VERSION,
                "file_sha256": old_hashes,
            }
        ),
        encoding="utf-8",
    )
    real_replace = F.os.replace
    replace_calls = 0

    def write_new_file(path, _headers, _rows):
        path.write_bytes(b"new-staged")

    def interrupt_second_publish(source, destination):
        nonlocal replace_calls
        replace_calls += 1
        if replace_calls == 2:
            raise OSError("synthetic publish interruption")
        real_replace(source, destination)

    monkeypatch.setattr(F, "write_xlsx", write_new_file)
    monkeypatch.setattr(F.os, "replace", interrupt_second_publish)

    with pytest.raises(OSError, match="synthetic publish interruption"):
        F.publish_day_exports(
            tmp_path,
            day_tag,
            [(name, ["h"], [[1]]) for name in names],
            {"export_schema_version": F.EXPORT_SCHEMA_VERSION},
        )

    assert json.loads(summary_path.read_text(encoding="utf-8"))["file_sha256"] == old_hashes
    assert F.already_fetched(tmp_path, day) == []


def test_batch_publish_failure_restores_every_previous_daily_file(tmp_path, monkeypatch):
    staging = tmp_path / ".batch-staging"
    output = tmp_path / "exports"
    staging.mkdir()
    output.mkdir()
    summaries = {}
    expected = {}
    days = ["2026-08-17", "2026-08-18"]
    for day in days:
        tag = day.replace("-", "")
        names = [f"{role}_{tag}.xlsx" for role in ("回款记录", "订单交付", "核销明细", "订单明细")]
        summaries[day] = {"files": names}
        for name in [*names, f"取数摘要_{tag}.json"]:
            (staging / name).write_bytes(f"new-{name}".encode())
            old_content = f"old-{name}".encode()
            (output / name).write_bytes(old_content)
            expected[name] = old_content

    real_replace = F.os.replace

    def fail_during_second_day(source, target):
        source_path = Path(source)
        if source_path.parent == staging and source_path.name == "订单交付_20260818.xlsx":
            raise OSError("synthetic batch publish failure")
        real_replace(source, target)

    monkeypatch.setattr(F.os, "replace", fail_during_second_day)

    with pytest.raises(F.FetchError, match="原有文件已恢复"):
        F.publish_batch_exports(staging, output, days, summaries)

    assert {
        path.name: path.read_bytes()
        for path in output.iterdir()
        if path.is_file()
    } == expected


def test_plain_option_and_relation():
    opts = {"k1": "整笔回款"}
    assert F._plain('["k1"]', opts) == "整笔回款"
    assert F._plain('[{"name":"某某客户"}]') == "某某客户"
    assert F._plain(None) == ""


def test_settlement_relation_recovers_order_when_xiadan_is_empty():
    controls = [
        {"controlId": "order", "controlName": "结算订单"},
        {"controlId": "written", "controlName": "订单已核销金额"},
        {"controlId": "amount", "controlName": "交付额/原币"},
        {"controlId": "currency", "controlName": "结算币种"},
    ]
    rows = [{
        "order": '[{"name":"SO26000001"}]',
        "written": "120.47",
        "amount": "120.50",
        "currency": "人民币CNY",
    }]

    got = F.extract_related_orders(rows, controls, F.REL_JIESUAN)

    assert got == [{
        "so": "SO26000001",
        "written_off": "120.47",
        "written_off_local": "",
        "deliver": "120.50",
        "rate": "",
        "currency": "人民币CNY",
        "name": "",
        "delivery_date": "",
        "delivery_date_status": "",
        "source": "结算",
    }]


def test_related_order_reads_project_delivery_date_from_order_detail():
    controls = [
        {"controlId": "order", "controlName": "SO"},
        {"controlId": "date", "controlName": "项目交付日期"},
    ]
    rows = [{"order": "SO24100160", "date": "2025-08-13"}]

    got = F.extract_related_orders(rows, controls, F.REL_XIADAN)

    assert got[0]["delivery_date"] == "2025-08-13"
    assert got[0]["delivery_date_status"] == "关联下单明确值"


def test_lookup_order_delivery_date_requires_unique_explicit_date():
    controls = [
        {"controlId": "order", "controlName": "SO"},
        {"controlId": "date", "controlName": "项目交付日期"},
    ]

    class FakeClient:
        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        @staticmethod
        def search_rows(_worksheet_id, _so):
            return [
                {"order": "SO24100160", "date": "2025-08-13"},
                {"order": "SO24100160", "date": "2025-08-13"},
                {"order": "SO99999999", "date": "2024-01-01"},
            ]

    value, status = F.lookup_order_delivery_date(
        FakeClient(), "orders", controls, "SO24100160"
    )
    assert value == "2025-08-13"
    assert status == "订单详情明确值"


def test_no_credentials_in_source():
    src = Path(__file__).resolve().parents[1] / "scripts" / "fetch_zhiyun.py"
    text = src.read_text(encoding="utf-8")
    # 禁止真实账号/密码痕迹（允许文档里出现变量名 ZHIYUN_PASS）
    assert "sharon" not in text.lower()
    assert "sharon1234" not in text
    assert "getpass" not in text  # 核销任务不在取数中途交互询问
    assert "input(" not in text
    assert "自动任务不会在取数过程中弹出账号密码询问" in text
    # 禁止把真实密码字面量赋给环境示例
    assert "PASS='****'" not in text


def test_historical_writeoffs_for_sos_gets_cross_parent_history_only():
    names = [
        "核销记录NUM", "回款记录NUM", "订单NUM", "本次核销金额", "本次核销金额本币",
        "核销日期", "币种", "汇率", "订单名称", "是否已撤销",
    ]
    controls = [
        {"controlId": f"c{i}", "controlName": name}
        for i, name in enumerate(names)
    ]

    def row(**values):
        return {
            f"c{i}": values.get(name, "")
            for i, name in enumerate(names)
        }

    class FakeClient:
        def controls(self, _worksheet_id):
            return controls

        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        def search_rows(self, _worksheet_id, so):
            assert so == "SO26000001"
            return [
                row(
                    核销记录NUM="HX_OLD_001",
                    回款记录NUM='[{"name":"AR_OLD_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=30,
                    本次核销金额本币=30,
                    核销日期="2026-06-11",
                ),
                row(  # 目标日当前行由父回款关联子表负责，不在这里重复补。
                    核销记录NUM="HX_NOW_001",
                    回款记录NUM='[{"name":"AR_NOW_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=10,
                    本次核销金额本币=10,
                    核销日期="2026-07-24",
                ),
                row(  # 已撤销历史行不计。
                    核销记录NUM="HX_OLD_002",
                    回款记录NUM='[{"name":"AR_OLD_002"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=5,
                    本次核销金额本币=5,
                    核销日期="2026-05-01",
                    是否已撤销="是",
                ),
                row(  # 全文搜索误命中的其它 SO 必须精确排除。
                    核销记录NUM="HX_OTHER",
                    回款记录NUM='[{"name":"AR_OTHER"}]',
                    订单NUM='[{"name":"SO260000010"}]',
                    本次核销金额=99,
                    本次核销金额本币=99,
                    核销日期="2026-04-01",
                ),
            ]

    got = F.historical_writeoffs_for_sos(
        FakeClient(), "WS_MX", ["SO26000001"], "2026-07-24"
    )
    assert len(got) == 2  # 撤销记录也保留给分类器审计，但不参与金额。
    assert got[0][0] == "HX_OLD_001"
    assert got[0][2] == "AR_OLD_001"
    assert got[0][3] == "2026-06-11"
    assert got[0][8] == "SO26000001"
    assert got[1][10] == "是"


def test_historical_writeoffs_only_dedup_same_record_id_and_never_business_fields():
    names = [
        "核销记录NUM", "回款记录NUM", "订单NUM", "本次核销金额", "本次核销金额本币",
        "核销日期", "币种", "汇率", "订单名称", "是否已撤销",
    ]
    controls = [
        {"controlId": f"c{i}", "controlName": name}
        for i, name in enumerate(names)
    ]

    def row(record_id):
        values = {
            "核销记录NUM": record_id,
            "回款记录NUM": '[{"name":"AR26070001"}]',
            "订单NUM": '[{"name":"SO26000001"}]',
            "本次核销金额": 100,
            "本次核销金额本币": 100,
            "核销日期": "2026-07-01",
            "币种": "CNY",
        }
        return {
            **{f"c{i}": values.get(name, "") for i, name in enumerate(names)},
            "rowid": f"ROW-{record_id or 'NONE'}",
        }

    class FakeClient:
        def controls(self, _worksheet_id):
            return controls

        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        def search_rows(self, _worksheet_id, _so):
            return [row("HX1"), row("HX1"), row("HX2"), row(""), row("")]

    got = F.historical_writeoffs_for_sos(
        FakeClient(), "WS_MX", ["SO26000001"], "2026-07-31"
    )
    assert [item[0] for item in got] == ["HX1", "HX2", "", ""]


def test_supplement_identifiers_are_searched_exactly_and_reported_against_exports():
    class FakeClient:
        @staticmethod
        def controls(_worksheet_id):
            return [
                {"controlId": "order", "controlName": "SO"},
                {"controlId": "ar", "controlName": "回款记录NUM"},
            ]

        @staticmethod
        def id_by_name(_controls, _name):
            return "relation"

        @staticmethod
        def datasource_of(_worksheet_id, relation):
            return {
                F.REL_XIADAN: "orders",
                F.REL_HEXIAO_MINGXI: "writeoffs",
                F.REL_SODLINE: "details",
            }.get(relation, "")

        @staticmethod
        def search_rows(worksheet_id, identifier):
            if worksheet_id == F.WS_HUIKUAN and identifier == "AR26070140":
                return [
                    {F.F_HK["ar"]: "AR26070140"},
                    {F.F_HK["ar"]: "AR260701400"},
                ]
            if worksheet_id == "orders" and identifier == "SO26020320":
                return [{"order": '[{"name":"SO26020320"}]'}]
            return []

    searched = F.search_supplement_identifiers(
        FakeClient(),
        ["AR26070140", "AR26079999"],
        ["SO26020320", "SO26029999"],
    )
    assert searched == {
        "found_ar_ids": ["AR26070140"],
        "found_so_ids": ["SO26020320"],
    }

    result = F.build_supplement_result(
        ["AR26070140", "AR26079999"],
        ["SO26020320", "SO26029999"],
        before={"ar_ids": set(), "so_ids": {"SO26020320"}},
        after={"ar_ids": {"AR26070140"}, "so_ids": {"SO26020320"}},
        searched=searched,
    )
    assert result["added"]["ar_ids"] == ["AR26070140"]
    assert result["existing"]["so_ids"] == ["SO26020320"]
    assert result["unresolved"]["ar_ids"] == ["AR26079999"]
    assert result["unresolved"]["so_ids"] == ["SO26029999"]


def test_resolve_fetch_dates_keeps_workdays_in_chronological_order():
    assert F.resolve_fetch_dates(
        single_date="",
        date_from="2026-08-14",
        date_to="2026-08-18",
        include_weekends=False,
    ) == ["2026-08-14", "2026-08-17", "2026-08-18"]


def test_date_range_fetch_logs_in_once_and_writes_every_calendar_day_separately(
    tmp_path, monkeypatch
):
    login_calls = []
    client_calls = []
    filtered_days = []

    def fake_login(base_url, user, password, *, headless):
        login_calls.append((base_url, user, password, headless))
        return "session-cookie", "account-id"

    class FakeClient:
        def __init__(self, base_url, cookie, *, account_id=""):
            client_calls.append((base_url, cookie, account_id))

        @staticmethod
        def controls(_worksheet_id):
            return [
                {
                    "controlId": "xiadan",
                    "controlName": F.REL_XIADAN,
                    "dataSource": "orders",
                }
            ]

        @staticmethod
        def option_maps(_controls):
            return {}

        @staticmethod
        def id_by_name(controls, name):
            for control in controls:
                if control.get("controlName") == name:
                    return control.get("controlId", "")
            return ""

        @staticmethod
        def datasource_of(_worksheet_id, relation):
            return "orders" if relation == F.REL_XIADAN else ""

        @staticmethod
        def filter_rows_by_date(_worksheet_id, _control_id, day):
            filtered_days.append(day)
            return [], 0

    monkeypatch.delenv("MD_PSS_ID", raising=False)
    monkeypatch.setattr(F, "resolve_credentials", lambda _args: ("user", "secret"))
    monkeypatch.setattr(F, "login_with_password", fake_login)
    monkeypatch.setattr(F, "ZhiyunClient", FakeClient)

    assert F.main([
        "--date-from", "2026-08-14",
        "--date-to", "2026-08-17",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 0

    assert len(login_calls) == 1
    assert len(client_calls) == 4
    assert sorted(filtered_days) == ["2026-08-14", "2026-08-15", "2026-08-16", "2026-08-17"]
    export_dir = tmp_path / "01_智云导出"
    for tag in ("20260814", "20260815", "20260816", "20260817"):
        assert (export_dir / f"回款记录_{tag}.xlsx").is_file()
        assert (export_dir / f"订单交付_{tag}.xlsx").is_file()
        assert (export_dir / f"核销明细_{tag}.xlsx").is_file()
        assert (export_dir / f"订单明细_{tag}.xlsx").is_file()
        assert (export_dir / f"取数摘要_{tag}.json").is_file()


def test_date_range_fetches_at_most_eight_days_concurrently_with_independent_clients(
    tmp_path, monkeypatch
):
    login_calls = []
    clients = []
    report_days = []
    state_lock = threading.Lock()
    first_wave = threading.Barrier(8, timeout=3)
    started = 0
    active = 0
    max_active = 0

    def fake_login(base_url, user, password, *, headless):
        login_calls.append((base_url, user, password, headless))
        return "session-cookie", "account-id"

    class FakeClient:
        def __init__(self, _base_url, _cookie, *, account_id=""):
            del account_id
            self.closed = False
            clients.append(self)

        @staticmethod
        def controls(_worksheet_id):
            return []

        def close(self):
            self.closed = True

    def fake_fetch_day(client, day, _out_dir, _ar_ids=(), _so_ids=()):
        nonlocal started, active, max_active
        with state_lock:
            started += 1
            ordinal = started
            active += 1
            max_active = max(max_active, active)
        try:
            if ordinal <= 8:
                first_wave.wait()
            return {"day": day, "client_id": id(client)}
        finally:
            with state_lock:
                active -= 1

    monkeypatch.delenv("MD_PSS_ID", raising=False)
    monkeypatch.setattr(F, "resolve_credentials", lambda _args: ("user", "secret"))
    monkeypatch.setattr(F, "login_with_password", fake_login)
    monkeypatch.setattr(F, "ZhiyunClient", FakeClient)
    monkeypatch.setattr(F, "fetch_day", fake_fetch_day)
    monkeypatch.setattr(F, "publish_batch_exports", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        F,
        "report_fetched_day",
        lambda day, *_args, **_kwargs: report_days.append(day),
    )

    assert F.main([
        "--date-from", "2026-08-01",
        "--date-to", "2026-08-10",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 0

    assert F.MAX_BATCH_FETCH_CONCURRENCY == 8
    assert len(login_calls) == 1
    assert len(clients) == 10
    assert len({id(client) for client in clients}) == 10
    assert all(client.closed for client in clients)
    assert max_active == 8
    assert report_days == [f"2026-08-{day:02d}" for day in range(1, 11)]


def test_date_range_uses_current_daily_exports_without_logging_in(
    tmp_path, monkeypatch
):
    export_dir = tmp_path / "01_智云导出"
    export_dir.mkdir()
    for day in ("2026-08-14", "2026-08-15", "2026-08-16", "2026-08-17"):
        stamp = day.replace("-", "")
        names = []
        for role in ("回款记录", "订单交付", "核销明细", "订单明细"):
            name = f"{role}_{stamp}.xlsx"
            (export_dir / name).write_bytes(b"fixture")
            names.append(name)
        (export_dir / f"取数摘要_{stamp}.json").write_text(
            json.dumps(
                {
                    "export_schema_version": F.EXPORT_SCHEMA_VERSION,
                    "file_sha256": {
                        name: F._sha256(export_dir / name)
                        for name in names
                    },
                }
            ),
            encoding="utf-8",
        )

    def unexpected_login(*_args, **_kwargs):
        raise AssertionError("全部日期命中缓存时不应登录智云")

    monkeypatch.delenv("MD_PSS_ID", raising=False)
    monkeypatch.setattr(F, "resolve_credentials", unexpected_login)

    assert F.main([
        "--date-from", "2026-08-14",
        "--date-to", "2026-08-17",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 0


def test_date_range_failure_prevents_reporting_batch_success(tmp_path, monkeypatch):
    filtered_days = []
    reported_days = []
    clients = []

    class FakeClient:
        def __init__(self, *_args, **_kwargs):
            self.closed = False
            clients.append(self)

        def close(self):
            self.closed = True

        @staticmethod
        def controls(_worksheet_id):
            return [{
                "controlId": "xiadan",
                "controlName": F.REL_XIADAN,
                "dataSource": "orders",
            }]

        @staticmethod
        def option_maps(_controls):
            return {}

        @staticmethod
        def id_by_name(controls, name):
            for control in controls:
                if control.get("controlName") == name:
                    return control.get("controlId", "")
            return ""

        @staticmethod
        def datasource_of(_worksheet_id, relation):
            return "orders" if relation == F.REL_XIADAN else ""

        @staticmethod
        def filter_rows_by_date(_worksheet_id, _control_id, day):
            filtered_days.append(day)
            if day == "2026-08-17":
                raise RuntimeError("simulated daily fetch failure")
            return [], 0

    monkeypatch.setenv("MD_PSS_ID", "session-cookie")
    monkeypatch.setattr(F, "ZhiyunClient", FakeClient)
    monkeypatch.setattr(
        F,
        "report_fetched_day",
        lambda day, *_args, **_kwargs: reported_days.append(day),
    )

    assert F.main([
        "--date-from", "2026-08-14",
        "--date-to", "2026-08-18",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 2
    assert sorted(filtered_days) == [
        "2026-08-14",
        "2026-08-15",
        "2026-08-16",
        "2026-08-17",
        "2026-08-18",
    ]
    assert reported_days == []
    assert clients
    assert all(client.closed for client in clients)
    export_dir = tmp_path / "01_智云导出"
    assert list(export_dir.glob("取数摘要_*.json")) == []
    assert list(export_dir.glob("*.xlsx")) == []


def test_date_range_rejects_identifier_supplement(tmp_path):
    import pytest

    with pytest.raises(SystemExit):
        F.main([
            "--date-from", "2026-08-14",
            "--date-to", "2026-08-17",
            "--supplement-ar", "AR26080001",
            "--workspace", str(tmp_path),
            "--skip-gap-check",
        ])


def test_identifier_supplement_closes_its_independent_client(tmp_path, monkeypatch):
    clients = []

    class FakeClient:
        def __init__(self, *_args, **_kwargs):
            self.closed = False
            clients.append(self)

        def close(self):
            self.closed = True

    monkeypatch.setenv("MD_PSS_ID", "session-cookie")
    monkeypatch.setattr(F, "ZhiyunClient", FakeClient)
    monkeypatch.setattr(
        F,
        "exported_supplement_identifiers",
        lambda *_args, **_kwargs: {"ar_ids": set(), "so_ids": set()},
    )
    monkeypatch.setattr(
        F,
        "search_supplement_identifiers",
        lambda *_args, **_kwargs: {"found_ar_ids": [], "found_so_ids": []},
    )
    monkeypatch.setattr(F, "fetch_day", lambda *_args, **_kwargs: {"files": []})
    monkeypatch.setattr(
        F,
        "build_supplement_result",
        lambda *_args, **_kwargs: {"added": {}, "existing": {}, "unresolved": {}},
    )
    monkeypatch.setattr(F, "report_fetched_day", lambda *_args, **_kwargs: None)

    assert F.main([
        "--date", "2026-08-17",
        "--supplement-ar", "AR26080001",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 0
    assert len(clients) == 1
    assert clients[0].closed
