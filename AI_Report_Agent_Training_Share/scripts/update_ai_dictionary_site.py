import argparse
import csv
import json
import re
from datetime import date
from pathlib import Path
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parent.parent
DEFAULT_HTML = ROOT / "site" / "index.html"
DEFAULT_MODELS_CSV = ROOT / "data" / "ai-models.csv"

DATA_PATTERN = re.compile(
    r"const DATA = (\[.*?\]);\s*\n\nconst FILTERS", re.DOTALL
)
NEWS_PATTERN = re.compile(
    r"const NEWS_WEEKS = (\[.*?\]);\s*\n", re.DOTALL
)


def normalize(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (value or "").casefold())


def compact_week_label(start_value: str, end_value: str) -> str:
    start = date.fromisoformat(start_value)
    end = date.fromisoformat(end_value)
    if start.year == end.year and start.month == end.month:
        return f"{start.strftime('%b')} {start.day} – {end.day}, {end.year}"
    if start.year == end.year:
        return (
            f"{start.strftime('%b')} {start.day} – "
            f"{end.strftime('%b')} {end.day}, {end.year}"
        )
    return (
        f"{start.strftime('%b')} {start.day}, {start.year} – "
        f"{end.strftime('%b')} {end.day}, {end.year}"
    )


def source_url(item: dict) -> str:
    candidates = [
        item.get("summary_source"),
        item.get("best_for_source"),
        *[
            update.get("source")
            for update in item.get("updates", [])
            if isinstance(update, dict)
        ],
    ]
    for candidate in candidates:
        if isinstance(candidate, dict) and candidate.get("url"):
            return candidate["url"]
    return ""


def source_label(item: dict, url: str) -> str:
    company = (item.get("company") or "").strip()
    if company:
        return company
    hostname = urlparse(url).hostname or ""
    return hostname.removeprefix("www.") or "Source"


def news_tag(section_title: str, item: dict) -> str:
    combined = " ".join(
        [
            item.get("tool", ""),
            item.get("heading_verb", ""),
            item.get("date", ""),
        ]
    ).casefold()
    if "rumor" in combined or "unconfirmed" in combined:
        return "Rumor"
    if section_title == "New AI Models and Rumors":
        return "New model"
    if section_title == "Technological Advancements in AI":
        return "Research"
    if section_title == "NVIDIA AI News":
        return "NVIDIA"
    return "Product launch"


def make_news_item(section_title: str, item: dict) -> dict:
    url = source_url(item)
    return {
        "title": item.get("tool", "Untitled update"),
        "maker": item.get("company", ""),
        "date": " ".join(
            value
            for value in [item.get("heading_verb", ""), item.get("date", "")]
            if value
        ),
        "tag": news_tag(section_title, item),
        "summary": item.get("summary", ""),
        "bullets": [
            update.get("text", "")
            for update in item.get("updates", [])
            if isinstance(update, dict) and update.get("text")
        ],
        "bestFor": item.get("best_for", ""),
        "source": source_label(item, url),
        "url": url,
    }


def load_embedded_json(pattern: re.Pattern, html: str, label: str) -> list:
    match = pattern.search(html)
    if not match:
        raise ValueError(f"Could not find embedded {label} data in index.html")
    return json.loads(match.group(1))


def find_dictionary_entry(data: list[dict], tool: str, company: str) -> dict | None:
    tool_key = normalize(tool)
    company_key = normalize(company)
    exact = [row for row in data if normalize(row.get("name", "")) == tool_key]
    if len(exact) == 1:
        return exact[0]

    candidates = []
    for row in data:
        name_key = normalize(row.get("name", ""))
        maker_key = normalize(row.get("maker", ""))
        name_matches = tool_key and (
            tool_key in name_key or name_key in tool_key
        )
        maker_matches = not company_key or not maker_key or (
            company_key in maker_key or maker_key in company_key
        )
        if name_matches and maker_matches:
            candidates.append(row)
    return candidates[0] if len(candidates) == 1 else None


def infer_tags(item: dict) -> str:
    text = " ".join(
        [
            item.get("tool", ""),
            item.get("heading_verb", ""),
            item.get("summary", ""),
            *[
                update.get("text", "")
                for update in item.get("updates", [])
                if isinstance(update, dict)
            ],
        ]
    ).casefold()
    tags = []
    keyword_tags = [
        ("frontier", ["frontier"]),
        ("multi", ["multimodal", "vision", "image"]),
        ("reason", ["reasoning", "reasoner"]),
        ("code", ["code", "coding", "developer"]),
        ("agent", ["agent", "agentic"]),
        ("open", ["open-weight", "open weight", "open model"]),
        ("edge", ["on-device", "edge"]),
        ("world", ["world model"]),
        ("audio", ["audio", "speech", "voice"]),
        ("image", ["image generation"]),
        ("video", ["video"]),
    ]
    for tag, keywords in keyword_tags:
        if any(keyword in text for keyword in keywords):
            tags.append(tag)
    return "|".join(tags or ["frontier"])


def infer_section(company: str, csv_rows: list[dict]) -> str:
    company_key = normalize(company)
    candidates = [
        row["Section"]
        for row in csv_rows
        if company_key
        and (
            company_key in normalize(row.get("Maker", ""))
            or normalize(row.get("Maker", "")) in company_key
        )
    ]
    if candidates:
        return max(set(candidates), key=candidates.count)
    return company or "New AI Models"


def make_dictionary_entry(item: dict, section: str) -> dict:
    heading = item.get("heading_verb", "Updated")
    summary = item.get("summary", "")
    return {
        "name": item.get("tool", "Unnamed model"),
        "maker": item.get("company", ""),
        "section": section,
        "description": f"{heading} — {summary}".strip(" —"),
        "tags": infer_tags(item),
        "hot": "true",
        "newsUpdate": f"{heading} {item.get('date', '')}".strip(),
        "summary": summary,
        "bestFor": f"[NVIDIA employee] {item.get('best_for', '')}".strip(),
        "beginner": "false",
    }


def read_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["Name", "Maker", "Section"])
        writer.writeheader()
        writer.writerows(rows)


def update_site(input_path: Path, html_path: Path, models_csv: Path) -> dict:
    newsletter = json.loads(input_path.read_text(encoding="utf-8"))
    html = html_path.read_text(encoding="utf-8")
    data = load_embedded_json(DATA_PATTERN, html, "DATA")
    news_weeks = load_embedded_json(NEWS_PATTERN, html, "NEWS_WEEKS")
    csv_rows = read_csv(models_csv)

    week_label = compact_week_label(
        newsletter["date_range"]["start"], newsletter["date_range"]["end"]
    )
    current_week = {"week": week_label, "sections": []}
    report_items = []
    new_models = []

    for section in newsletter.get("sections", []):
        section_title = section.get("title", "Other")
        items = section.get("items", [])
        if not items:
            continue
        current_week["sections"].append(
            {
                "label": section_title,
                "items": [
                    make_news_item(section_title, item) for item in items
                ],
            }
        )
        report_items.extend(items)

        if section_title != "New AI Models and Rumors":
            continue
        for item in items:
            if news_tag(section_title, item) == "Rumor":
                continue
            if find_dictionary_entry(
                data, item.get("tool", ""), item.get("company", "")
            ):
                continue
            section_name = infer_section(item.get("company", ""), csv_rows)
            data.append(make_dictionary_entry(item, section_name))
            csv_rows.append(
                {
                    "Name": item.get("tool", ""),
                    "Maker": item.get("company", ""),
                    "Section": section_name,
                }
            )
            new_models.append(item.get("tool", ""))

    for row in data:
        row.pop("newsUpdate", None)
    for item in report_items:
        row = find_dictionary_entry(
            data, item.get("tool", ""), item.get("company", "")
        )
        if row:
            heading = item.get("heading_verb", "Updated")
            row["newsUpdate"] = (
                f"{heading} {item.get('date', '')} — "
                f"{item.get('summary', '')}"
            ).strip(" —")

    news_weeks = [
        current_week,
        *[week for week in news_weeks if week.get("week") != week_label],
    ]

    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    news_json = json.dumps(
        news_weeks, ensure_ascii=False, separators=(",", ":")
    )
    html = DATA_PATTERN.sub(
        lambda _: f"const DATA = {data_json};\n\nconst FILTERS",
        html,
        count=1,
    )
    html = NEWS_PATTERN.sub(
        lambda _: f"const NEWS_WEEKS = {news_json};\n", html, count=1
    )

    banner_range = week_label.replace(", 2026", "")
    html = re.sub(
        r"Updates in this week's AI News \([^<]+\)",
        f"Updates in this week's AI News ({banner_range})",
        html,
    )
    html = re.sub(
        r"\d+ AI models · Functional overview",
        f"{len(data)} AI models · Functional overview",
        html,
    )
    html = re.sub(
        r"\d+ AI models, tools &amp; companies · Updated [A-Za-z]+ \d{4}",
        (
            f"{len(data)} AI models, tools &amp; companies · Updated "
            f"{date.fromisoformat(newsletter['date_range']['end']).strftime('%B %Y')}"
        ),
        html,
    )

    html_path.write_text(html, encoding="utf-8")
    if new_models:
        write_csv(models_csv, csv_rows)

    return {
        "week": week_label,
        "news_items": sum(
            len(section["items"]) for section in current_week["sections"]
        ),
        "dictionary_entries": len(data),
        "new_models": new_models,
        "archived_weeks": len(news_weeks),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--html", default=str(DEFAULT_HTML))
    parser.add_argument("--models-csv", default=str(DEFAULT_MODELS_CSV))
    args = parser.parse_args()

    result = update_site(
        Path(args.input).resolve(),
        Path(args.html).resolve(),
        Path(args.models_csv).resolve(),
    )
    print(f"Website week updated: {result['week']}")
    print(f"Weekly news items published: {result['news_items']}")
    print(f"Dictionary entries: {result['dictionary_entries']}")
    print(f"Archived weeks: {result['archived_weeks']}")
    if result["new_models"]:
        print("New models added: " + "; ".join(result["new_models"]))
    else:
        print("New models added: 0")


if __name__ == "__main__":
    main()
