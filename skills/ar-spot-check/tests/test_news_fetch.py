import json
import sys
from datetime import date
from pathlib import Path
from urllib.parse import unquote

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import news_fetch  # noqa: E402


def rss(items: list[tuple[str, str, str, str]]) -> str:
    body = "".join(
        f"<item><title>{title}</title><link>{url}</link><description>{summary}</description><pubDate>{published}</pubDate></item>"
        for title, url, summary, published in items
    )
    return f"<?xml version='1.0' encoding='utf-8'?><rss version='2.0'><channel>{body}</channel></rss>"


PROBE = rss([("北京市", "https://example.com/probe", "探针", "周一, 29 9月 2026 08:00:00 GMT")])
EMPTY = "<?xml version='1.0' encoding='utf-8'?><rss version='2.0'><channel><title>空</title></channel></rss>"
HIT = rss(
    [
        ("甲旧闻", "https://example.com/old", "<b>亏损</b>扩大", "Mon, 01 Jan 2020 00:00:00 GMT"),
        ("别人的近闻", "https://example.com/other", "资金链紧张", "周一, 28 9月 2026 20:29:00 GMT"),
        ("甲近闻", "https://example.com/new", "<b>资金链</b>紧张", "周一, 28 9月 2026 20:29:00 GMT"),
    ]
)


def query_of(url: str) -> str:
    return unquote(url.split("q=", 1)[1])


def names_file(path: Path, names: list[str]) -> Path:
    target = path / "必搜.txt"
    target.write_text("\n".join(names) + ("\n" if names else ""), encoding="utf-8")
    return target


def run(tmp_path: Path, monkeypatch, names: list[str], fetch, batch: int = 20):
    monkeypatch.setattr(news_fetch, "fetch_rss", fetch)
    monkeypatch.setattr(news_fetch, "EMPTY_BATCH", batch)
    return news_fetch.main(
        ["--names", str(names_file(tmp_path, names)), "--out", str(tmp_path / "新闻"), "--today", "2026-09-29", "--workers", "2"]
    )


def test_name_forms_and_mentions():
    assert "甲乙丙丁" in news_fetch.name_forms("甲乙丙丁有限公司")
    forms = news_fetch.name_forms("甲乙丙丁（内里品牌）有限公司")
    assert "内里品牌" in forms
    assert "甲乙丙丁" in forms
    assert news_fetch.mentions("甲", "甲近闻", "无关")
    assert not news_fetch.mentions("甲", "近闻", "资金链紧张")
    assert news_fetch.search_query("甲“乙”") == '"甲乙"'


def test_snippet_without_the_customer_is_a_miss(tmp_path, monkeypatch):
    unrelated = rss([("近闻", "https://example.com/other", "资金链紧张", "周一, 28 9月 2026 20:29:00 GMT")])

    def fetch(url, timeout):
        query = query_of(url)
        if query == "北京":
            return PROBE
        if query == '"甲"':
            return unrelated
        raise AssertionError(query)

    code = run(tmp_path, monkeypatch, ["甲"], fetch, batch=1)
    assert code == 0
    missed = [json.loads(line) for line in (tmp_path / "新闻" / "没搜到.jsonl").read_text(encoding="utf-8").splitlines()]
    assert missed[0]["customer"] == "甲"
    assert missed[0]["summary"] == "未查到"
    assert not (tmp_path / "新闻" / "要判断.json").exists()


def test_chinese_and_english_dates():
    assert news_fetch.published_date("周一, 28 9月 2026 20:29:00 GMT") == date(2026, 9, 28)
    assert news_fetch.published_date("Mon, 01 Jan 2020 00:00:00 GMT") == date(2020, 1, 1)
    assert news_fetch.published_date("看不出来") is None


def test_empty_list_does_not_search(tmp_path, monkeypatch, capsys):
    def fetch(url, timeout):
        raise AssertionError(url)

    code = run(tmp_path, monkeypatch, [], fetch)
    assert code == 0
    assert "search=0" in capsys.readouterr().out
    assert not (tmp_path / "新闻" / "没搜到.jsonl").exists()


def test_snippets_and_empty_are_split(tmp_path, monkeypatch, capsys):
    def fetch(url, timeout):
        query = query_of(url)
        if query == "北京":
            return PROBE
        if query == '"甲"':
            return HIT
        if query == '"乙"':
            return EMPTY
        raise AssertionError(query)

    code = run(tmp_path, monkeypatch, ["甲", "乙"], fetch, batch=1)
    assert code == 0
    out = capsys.readouterr().out
    assert "甲" not in out and "乙" not in out
    missed = [json.loads(line) for line in (tmp_path / "新闻" / "没搜到.jsonl").read_text(encoding="utf-8").splitlines()]
    assert missed == [{"customer": "乙", "summary": "未查到", "url": "", "note": "未查到", "risk": "无", "retrieved": "2026-09-29"}]
    judged = json.loads((tmp_path / "新闻" / "要判断.json").read_text(encoding="utf-8"))
    assert [row["customer"] for row in judged] == ["甲"]
    assert judged[0]["hits"] == [
        {"title": "甲近闻", "url": "https://example.com/new", "summary": "资金链紧张", "published": "2026-09-28"}
    ]


def test_probe_failure_does_not_invent_a_miss(tmp_path, monkeypatch, capsys):
    def fetch(url, timeout):
        return "<html>blocked</html>"

    code = run(tmp_path, monkeypatch, ["甲"], fetch)
    captured = capsys.readouterr().out
    assert code == 2
    assert "搜索打不开" in captured
    assert "甲" not in captured
    assert not (tmp_path / "新闻" / "没搜到.jsonl").exists()
    assert not (tmp_path / "新闻" / "失败.txt").exists()


def test_customer_failure_is_not_a_miss(tmp_path, monkeypatch):
    def fetch(url, timeout):
        if query_of(url) == "北京":
            return PROBE
        raise TimeoutError("down")

    code = run(tmp_path, monkeypatch, ["甲"], fetch)
    assert code == 2
    assert (tmp_path / "新闻" / "失败.txt").read_text(encoding="utf-8").splitlines() == ["甲"]
    assert not (tmp_path / "新闻" / "没搜到.jsonl").exists()


def test_empty_result_is_not_written_when_search_goes_dark(tmp_path, monkeypatch):
    calls = {"probe": 0}

    def fetch(url, timeout):
        if query_of(url) == "北京":
            calls["probe"] += 1
            return PROBE if calls["probe"] == 1 else EMPTY
        return EMPTY

    code = run(tmp_path, monkeypatch, ["甲"], fetch, batch=1)
    assert code == 2
    assert (tmp_path / "新闻" / "失败.txt").read_text(encoding="utf-8").splitlines() == ["甲"]
    assert not (tmp_path / "新闻" / "没搜到.jsonl").exists()


def test_skips_customers_already_saved(tmp_path, monkeypatch):
    news = tmp_path / "新闻"
    news.mkdir()
    (news / "沿用.jsonl").write_text(
        json.dumps({"customer": "甲", "summary": "未查到", "url": "", "note": "未查到", "risk": "无"}, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    def fetch(url, timeout):
        query = query_of(url)
        if query == '"甲"':
            raise AssertionError(query)
        if query == "北京":
            return PROBE
        return EMPTY

    code = run(tmp_path, monkeypatch, ["甲", "乙"], fetch, batch=1)
    assert code == 0
    missed = (news / "没搜到.jsonl").read_text(encoding="utf-8")
    assert "乙" in missed and "甲" not in missed
