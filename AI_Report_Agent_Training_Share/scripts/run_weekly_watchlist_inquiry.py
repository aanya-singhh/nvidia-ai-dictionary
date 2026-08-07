import argparse
import csv
import json
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "data" / "ai-models.csv"
JSON_PATH = ROOT / "weekly_ai_supplemental_research_ledger_2026-07-27.json"
MD_PATH = ROOT / "weekly_ai_supplemental_research_ledger_2026-07-27.md"
WINDOW_START = "2026-07-20"
WINDOW_END = "2026-07-27"
SEARCH_TERMS = "(release OR released OR launched OR announced OR preview OR beta OR update OR changelog OR safety)"
DATE_RE = re.compile(r"$^")


def current_monday(today: date | None = None) -> date:
    today = today or date.today()
    return today - timedelta(days=today.weekday())


def build_date_signal_regex(start: date, end: date) -> re.Pattern:
    variants = []
    cursor = start
    while cursor <= end:
        variants.extend(
            [
                cursor.isoformat(),
                cursor.strftime("%Y/%m/%d"),
                f"{cursor.strftime('%B')} {cursor.day}, {cursor.year}",
                f"{cursor.strftime('%b')} {cursor.day}, {cursor.year}",
            ]
        )
        cursor += timedelta(days=1)
    return re.compile("|".join(re.escape(value) for value in variants if value), re.IGNORECASE)


def load_targets():
    with CSV_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return [
            {
                "name": row["Name"].strip(),
                "maker": row["Maker"].strip(),
                "section": (row.get("Section") or row["Maker"]).strip(),
            }
            for row in reader
            if row.get("Name") and row.get("Maker")
        ]


def search_target(index_target):
    index, target = index_target
    query = (
        f'"{target["name"]}" "{target["maker"]}" {SEARCH_TERMS} '
        f"after:{(date.fromisoformat(WINDOW_START) - timedelta(days=1)).isoformat()} "
        f"before:{(date.fromisoformat(WINDOW_END) + timedelta(days=1)).isoformat()}"
    )
    url = "https://www.bing.com/search?format=rss&q=" + urllib.parse.quote(query)
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/136 Safari/537.36"
            )
        },
    )
    error = None
    results = []
    for attempt in range(3):
        try:
            with urllib.request.urlopen(request, timeout=25) as response:
                payload = response.read()
            root = ET.fromstring(payload)
            for item in root.findall("./channel/item")[:5]:
                title = (item.findtext("title") or "").strip()
                link = (item.findtext("link") or "").strip()
                description = re.sub(
                    r"<[^>]+>", " ", item.findtext("description") or ""
                )
                description = re.sub(r"\s+", " ", description).strip()
                results.append(
                    {
                        "title": title,
                        "url": link,
                        "snippet": description,
                        "in_window_date_signal": bool(
                            DATE_RE.search(f"{title} {description}")
                        ),
                    }
                )
            break
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
            time.sleep(1.5 * (attempt + 1))

    has_date_signal = any(result["in_window_date_signal"] for result in results)
    outcome = (
        "candidate-found"
        if has_date_signal
        else "checked-no-in-window-candidate-surfaced"
    )
    return index, {
        **target,
        "target_key": f'{target["name"]} | {target["maker"]} | {target["section"]}',
        "supplemental_query": query,
        "outcome": outcome if not error or results else "search-error",
        "results": results,
        "error": error if not results else None,
    }


def main():
    global CSV_PATH, JSON_PATH, MD_PATH, WINDOW_START, WINDOW_END, DATE_RE

    parser = argparse.ArgumentParser()
    parser.add_argument("--models-csv", default=str(CSV_PATH))
    parser.add_argument("--week-end", help="Report end date in YYYY-MM-DD format.")
    parser.add_argument("--output-dir", default=str(ROOT))
    parser.add_argument("--max-workers", type=int, default=8)
    args = parser.parse_args()

    end = date.fromisoformat(args.week_end) if args.week_end else current_monday()
    start = end - timedelta(days=7)
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    CSV_PATH = Path(args.models_csv).resolve()
    WINDOW_START = start.isoformat()
    WINDOW_END = end.isoformat()
    JSON_PATH = output_dir / f"weekly_ai_supplemental_research_ledger_{WINDOW_END}.json"
    MD_PATH = output_dir / f"weekly_ai_supplemental_research_ledger_{WINDOW_END}.md"
    DATE_RE = build_date_signal_regex(start, end)

    targets = load_targets()
    records = [None] * len(targets)
    with ThreadPoolExecutor(max_workers=args.max_workers) as pool:
        futures = [
            pool.submit(search_target, indexed) for indexed in enumerate(targets)
        ]
        for completed, future in enumerate(as_completed(futures), start=1):
            index, record = future.result()
            records[index] = record
            if completed % 25 == 0 or completed == len(futures):
                print(f"Completed {completed}/{len(futures)}")

    candidate_count = sum(
        record["outcome"] == "candidate-found" for record in records
    )
    error_count = sum(record["outcome"] == "search-error" for record in records)
    payload = {
        "pass_type": "supplemental exact-target discovery pass",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "report_window": {"start": WINDOW_START, "end": WINDOW_END},
        "source_csv": CSV_PATH.name,
        "live_target_count": len(records),
        "completed_target_count": len(records) - error_count,
        "candidate_count": candidate_count,
        "search_error_count": error_count,
        "method_note": (
            "One additional exact-name and maker discovery query was executed "
            "for every live CSV composite target. Candidate results require "
            "primary-source opening and event-date verification before inclusion."
        ),
        "records": records,
    }
    JSON_PATH.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    lines = [
        "# Supplemental Weekly AI Watchlist Research Pass",
        "",
        f"Report window: {WINDOW_START} to {WINDOW_END} (inclusive)",
        f"Live CSV composite targets: {len(records)}",
        f"Completed exact-target queries: {len(records) - error_count}",
        f"Candidates with an in-window date signal: {candidate_count}",
        f"Search errors: {error_count}",
        "",
        "This is an additional discovery pass. A search hit is not sufficient "
        "for report inclusion; the primary source and event date must be opened "
        "and verified.",
        "",
        "## Results",
        "",
    ]
    for record in records:
        lines.append(f"### {record['target_key']}")
        lines.append("")
        lines.append(f"- Outcome: {record['outcome']}")
        lines.append(f"- Query: {record['supplemental_query']}")
        if record["error"]:
            lines.append(f"- Error: {record['error']}")
        for result in record["results"]:
            signal = " — in-window date signal" if result["in_window_date_signal"] else ""
            lines.append(f"- [{result['title']}]({result['url']}){signal}")
        lines.append("")
    MD_PATH.write_text("\n".join(lines), encoding="utf-8")

    print(f"Wrote {JSON_PATH}")
    print(f"Wrote {MD_PATH}")
    print(
        f"Summary: {len(records) - error_count}/{len(records)} completed; "
        f"{candidate_count} candidates; {error_count} errors"
    )


if __name__ == "__main__":
    main()
