import json
import sys
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import news_plan  # noqa: E402


def write_facts(path: Path, names: list[str]) -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "事实"
    sheet.append(["销售", "客户", "交付月份"])
    for name in names:
        sheet.append(["甲销", name, 202609])
    workbook.save(path)


def write_prior(directory: Path, rows: list[dict]) -> None:
    directory.mkdir(exist_ok=True)
    (directory / "old.jsonl").write_text(
        "\n".join(json.dumps(row, ensure_ascii=False) for row in rows) + "\n",
        encoding="utf-8",
    )


def run(tmp_path: Path, names: list[str], prior_rows: list[dict] | None, retrieved: str, today: str):
    facts = tmp_path / "事实.xlsx"
    write_facts(facts, names)
    prior = tmp_path / "prior"
    if prior_rows is not None:
        write_prior(prior, prior_rows)
    out = tmp_path / "新闻"
    code = news_plan.main(
        [
            "--facts",
            str(facts),
            "--today",
            today,
            "--out",
            str(out),
            *(["--prior", str(prior), "--prior-retrieved", retrieved] if prior_rows is not None else []),
        ]
    )
    kept = []
    text = (out / "沿用.jsonl").read_text(encoding="utf-8") if code == 0 else ""
    for line in text.splitlines():
        if line.strip():
            kept.append(json.loads(line))
    search = (out / "必搜.txt").read_text(encoding="utf-8").splitlines() if code == 0 else []
    return code, kept, search


def test_no_prior_searches_everyone_once(tmp_path, capsys):
    code, kept, search = run(tmp_path, ["甲", "甲", "乙"], None, "", "2026-09-29")
    assert code == 0
    assert kept == []
    assert search == ["甲", "乙"]
    assert "甲" not in capsys.readouterr().out


def test_reuses_a_real_search_from_the_last_seven_days(tmp_path):
    rows = [
        {"customer": "甲", "summary": "未查到", "url": "", "note": "未查到", "risk": "无"},
        {"customer": "乙", "summary": "业绩下降", "url": "https://example.com/b", "note": "公告", "risk": "低", "reason": "看不出会拖欠。"},
    ]
    code, kept, search = run(tmp_path, ["甲", "乙", "丙"], rows, "2026-09-23", "2026-09-29")
    assert code == 0
    assert [row["customer"] for row in kept] == ["甲", "乙"]
    assert kept[0]["retrieved"] == "2026-09-23"
    assert search == ["丙"]


def test_seventh_day_reuses_and_eighth_day_searches_again(tmp_path):
    rows = [{"customer": "甲", "summary": "未查到", "url": "", "note": "未查到"}]
    code, kept, search = run(tmp_path, ["甲"], rows, "2026-09-22", "2026-09-29")
    assert code == 0 and [row["customer"] for row in kept] == ["甲"] and search == []
    code, kept, search = run(tmp_path, ["甲"], rows, "2026-09-21", "2026-09-29")
    assert code == 0 and kept == [] and search == ["甲"]


def test_placeholder_or_link_without_a_grade_is_searched_again(tmp_path):
    rows = [
        {"customer": "甲", "summary": "本次未检索", "url": "", "note": ""},
        {"customer": "乙", "summary": "有一篇", "url": "https://example.com/b", "note": "公告", "risk": "无"},
    ]
    code, kept, search = run(tmp_path, ["甲", "乙"], rows, "2026-09-29", "2026-09-29")
    assert code == 0
    assert kept == []
    assert search == ["甲", "乙"]


def test_prior_customer_not_in_this_period_is_dropped(tmp_path):
    rows = [{"customer": "外人", "summary": "未查到", "url": "", "note": "未查到", "risk": "无"}]
    code, kept, search = run(tmp_path, ["甲"], rows, "2026-09-29", "2026-09-29")
    assert code == 0
    assert kept == []
    assert search == ["甲"]
