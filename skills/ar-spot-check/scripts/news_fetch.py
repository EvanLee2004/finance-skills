#!/usr/bin/env python3
"""按必搜名单检索标题和摘录。不打开正文，不改风险等级。"""

from __future__ import annotations

import argparse
import html
import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.error import URLError
from urllib.parse import quote
from urllib.request import Request, urlopen
import xml.etree.ElementTree as ET

PROBE_QUERY = "北京"
QUERY_TAIL = "被执行 失信 破产 停产 裁员 亏损 业绩下降 资金链"
SEARCH = "https://cn.bing.com/search?format=rss&q="
HALF_YEAR_DAYS = 183
MAX_HITS = 5
EMPTY_BATCH = 20
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
CN_PUB = re.compile(r"(\d{1,2})\s+(\d{1,2})月\s+(\d{4})")


def ask(text: str) -> int:
    print("status=ask")
    print(f"ask={text}")
    return 2


def search_url(query: str) -> str:
    return SEARCH + quote(query, safe="")


def clean_text(value: str, limit: int) -> str:
    text = re.sub(r"<[^>]+>", " ", value or "")
    text = html.unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return text[:limit].rstrip()


def published_date(value: str):
    text = str(value or "").strip()
    if not text:
        return None
    matched = CN_PUB.search(text)
    if matched:
        day, month, year = (int(part) for part in matched.groups())
        try:
            return date(year, month, day)
        except ValueError:
            return None
    try:
        return parsedate_to_datetime(text).date()
    except (TypeError, ValueError, IndexError, OverflowError):
        pass
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def parse_items(text: str):
    raw = text.lstrip("\ufeff").lstrip()
    if not raw.startswith("<"):
        return None
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return None
    kind = root.tag.split("}")[-1].lower()
    if kind not in {"rss", "feed", "rdf"}:
        return None
    items = []
    for node in root.iter():
        if node.tag.split("}")[-1].lower() not in {"item", "entry"}:
            continue
        fields = {}
        for child in list(node):
            key = child.tag.split("}")[-1].lower()
            if key in fields:
                continue
            text_value = "".join(child.itertext()).strip()
            if key == "link" and not text_value:
                text_value = (child.attrib.get("href") or "").strip()
            fields[key] = text_value
        url = fields.get("link") or ""
        if not url.startswith("http"):
            url = fields.get("guid") or ""
        title = clean_text(fields.get("title") or "", 120)
        summary = clean_text(fields.get("description") or fields.get("summary") or "", 180)
        if not url.startswith("http") or not (title or summary):
            continue
        items.append(
            {
                "title": title,
                "url": url.strip(),
                "summary": summary,
                "published": (fields.get("pubdate") or fields.get("published") or "").strip(),
            }
        )
    return items


def keep_hit(item: dict, today: date):
    found = published_date(item.get("published") or "")
    if found is not None and (today - found).days > HALF_YEAR_DAYS:
        return None
    return {
        "title": item["title"],
        "url": item["url"],
        "summary": item["summary"],
        "published": found.isoformat() if found else "",
    }


def fetch_rss(url: str, timeout: float) -> str:
    request = Request(
        url,
        headers={
            "User-Agent": UA,
            "Accept": "application/rss+xml, application/xml, text/xml;q=0.9, */*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        return response.read().decode("utf-8", errors="replace")


def probe_ok(fetch, timeout: float) -> bool:
    try:
        text = fetch(search_url(PROBE_QUERY), timeout)
    except (URLError, TimeoutError, OSError, ValueError):
        return False
    return bool(parse_items(text))


def lookup(name: str, fetch, timeout: float, today: date):
    try:
        text = fetch(search_url(f"{name} {QUERY_TAIL}"), timeout)
    except (URLError, TimeoutError, OSError, ValueError):
        return "fail", []
    items = parse_items(text)
    if items is None:
        return "fail", []
    kept = []
    for item in items:
        hit = keep_hit(item, today)
        if hit:
            kept.append(hit)
        if len(kept) == MAX_HITS:
            break
    if kept:
        return "hits", kept
    return "empty", []


def read_names(path: Path) -> list[str]:
    names = []
    seen = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        name = line.strip().lstrip("\ufeff")
        if name and name not in seen:
            seen.add(name)
            names.append(name)
    return names


def already_done(directory: Path) -> set[str]:
    found = set()
    if not directory.exists():
        return found
    for path in sorted(directory.glob("*.jsonl")):
        for line in path.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            row = json.loads(line)
            name = str(row.get("customer") or "").strip()
            if name:
                found.add(name)
    judge_path = directory / "要判断.json"
    if not judge_path.exists():
        return found
    data = json.loads(judge_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("judge")
    for row in data:
        if isinstance(row, dict):
            name = str(row.get("customer") or "").strip()
            if name:
                found.add(name)
    return found


def load_judge(path: Path) -> list[dict]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("judge")
    return [row for row in data if isinstance(row, dict)]


def write_judge(path: Path, existing: list[dict], fresh: list[dict]) -> None:
    by_name = {str(row.get("customer") or "").strip(): row for row in fresh}
    merged = []
    seen = set()
    for row in existing:
        name = str(row.get("customer") or "").strip()
        if not name or name in seen:
            continue
        seen.add(name)
        merged.append(by_name.get(name, row))
    for row in fresh:
        name = row["customer"]
        if name not in seen:
            merged.append(row)
            seen.add(name)
    path.write_text(json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="检索必搜名单的标题和摘录")
    parser.add_argument("--names", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--today", required=True)
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--timeout", type=float, default=12)
    args = parser.parse_args(argv)
    today = published_date(args.today) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", args.today or "") else None
    if today is None or str(today) != args.today:
        return ask("今天的日期要写成 YYYY-MM-DD。")
    if not 1 <= args.workers <= 8:
        return ask("并发只接受 1 到 8。")
    if not 1 <= args.timeout <= 60:
        return ask("单家超时要在 1 到 60 秒。")
    if not args.names.exists():
        return ask("还没有必搜名单。先跑 news_plan。")
    try:
        names = read_names(args.names)
    except UnicodeDecodeError:
        return ask("必搜名单读不了。")
    args.out.mkdir(parents=True, exist_ok=True)
    try:
        done = already_done(args.out)
    except (json.JSONDecodeError, UnicodeDecodeError, ValueError):
        return ask("新闻目录里已有的结果读不了。")
    todo = [name for name in names if name not in done]
    skipped = len(names) - len(todo)
    if not todo:
        print("status=ok")
        print("search=0")
        print("empty=0")
        print("judge=0")
        print(f"skipped={skipped}")
        print("failed=0")
        return 0
    if not probe_ok(fetch_rss, args.timeout):
        return ask("搜索打不开。先别把客户写成未查到，也不要开网页补搜。")
    results = {}
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(lookup, name, fetch_rss, args.timeout, today): name for name in todo}
        for future in as_completed(futures):
            results[futures[future]] = future.result()
    failed = []
    judged = []
    confirmed = []
    pending = []
    blocked = False
    today_text = today.isoformat()
    for name in todo:
        kind, hits = results[name]
        if kind == "fail":
            failed.append(name)
            continue
        if kind == "hits":
            judged.append({"customer": name, "retrieved": today_text, "hits": hits})
            continue
        if blocked:
            failed.append(name)
            continue
        pending.append(name)
        if len(pending) >= EMPTY_BATCH:
            if probe_ok(fetch_rss, args.timeout):
                confirmed.extend(pending)
            else:
                failed.extend(pending)
                blocked = True
            pending = []
    if pending:
        if blocked or not probe_ok(fetch_rss, args.timeout):
            failed.extend(pending)
        else:
            confirmed.extend(pending)
    if confirmed:
        lines = [
            json.dumps(
                {"customer": name, "summary": "未查到", "url": "", "note": "未查到", "risk": "无", "retrieved": today_text},
                ensure_ascii=False,
            )
            for name in confirmed
        ]
        with (args.out / "没搜到.jsonl").open("a", encoding="utf-8") as handle:
            handle.write("\n".join(lines) + "\n")
    if judged:
        try:
            existing = load_judge(args.out / "要判断.json")
        except (json.JSONDecodeError, UnicodeDecodeError, ValueError):
            return ask("要判断.json 读不了。")
        write_judge(args.out / "要判断.json", existing, judged)
    fail_path = args.out / "失败.txt"
    if failed:
        fail_path.write_text("\n".join(failed) + "\n", encoding="utf-8")
    elif fail_path.exists():
        fail_path.unlink()
    print("status=ok" if not failed else "status=ask")
    print(f"search={len(todo)}")
    print(f"empty={len(confirmed)}")
    print(f"judge={len(judged)}")
    print(f"skipped={skipped}")
    print(f"failed={len(failed)}")
    if failed:
        print("ask=有客户这次没搜成。看失败.txt，重跑这一条。不要写成未查到，也不要开子代理补搜。")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
