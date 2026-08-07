import argparse
import csv
import json
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path


DEFAULT_MODELS_CSV = "ai-models.csv"
REPORT_KEYWORDS = "release OR launched OR announced OR preview OR beta OR benchmark OR update OR changelog OR safety"


def current_monday(today: date | None = None) -> date:
    today = today or date.today()
    return today - timedelta(days=today.weekday())


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def load_models(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"Name", "Maker", "Section"}
        missing = required - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"Missing required column(s): {', '.join(sorted(missing))}")

        rows = []
        seen = set()
        for row in reader:
            name = (row.get("Name") or "").strip()
            maker = (row.get("Maker") or "").strip()
            section = (row.get("Section") or "").strip()
            if not name or not maker:
                continue
            key = (name.lower(), maker.lower(), section.lower())
            if key in seen:
                continue
            seen.add(key)
            rows.append({"name": name, "maker": maker, "section": section or maker})
    return rows


def build_model_queries(model: dict, start: date, end: date) -> list[str]:
    name = model["name"]
    maker = model["maker"]
    return [
        f'"{name}" "{maker}" ({REPORT_KEYWORDS}) after:{start.isoformat()} before:{(end + timedelta(days=1)).isoformat()}',
        f'site:{maker_domain_hint(maker)} "{name}" "{end.year}" "{end.strftime("%B")}"',
    ]


def maker_domain_hint(maker: str) -> str:
    maker_lower = maker.lower()
    domain_hints = {
        "openai": "openai.com",
        "anthropic": "anthropic.com",
        "google": "blog.google",
        "deepmind": "deepmind.google",
        "microsoft": "microsoft.com",
        "github": "github.blog",
        "meta": "ai.meta.com",
        "mistral": "mistral.ai",
        "xai": "x.ai",
        "databricks": "databricks.com",
        "nvidia": "nvidia.com",
        "amazon": "aws.amazon.com",
        "alibaba": "qwen.ai",
        "qwen": "qwen.ai",
        "huggingface": "huggingface.co",
        "cohere": "cohere.com",
        "perplexity": "perplexity.ai",
    }
    for needle, domain in domain_hints.items():
        if needle in maker_lower:
            return domain
    return "official website"


def grouped_models(models: list[dict]) -> dict[str, list[dict]]:
    grouped = defaultdict(list)
    for model in models:
        grouped[model["section"]].append(model)
    return {section: sorted(items, key=lambda item: (item["maker"], item["name"])) for section, items in sorted(grouped.items())}


def build_payload(models: list[dict], start: date, end: date) -> dict:
    grouped = grouped_models(models)
    return {
        "date_range": {
            "start": start.isoformat(),
            "end": end.isoformat(),
            "display": f"{start.strftime('%B')} {start.day}, {start.year} - {end.strftime('%B')} {end.day}, {end.year}",
        },
        "source": DEFAULT_MODELS_CSV,
        "model_count": len(models),
        "section_counts": {section: len(items) for section, items in grouped.items()},
        "rules": [
            "Only include a model/tool if the update itself is dated inside the report window.",
            "Do not carry forward prior-week model entries unless there is a new in-window update.",
            "Prefer primary sources: official blogs, release notes, docs, changelogs, papers, or company announcements.",
            "Use reputable reporting only when the primary source is unavailable, restricted, or the item is explicitly a rumor/watch item.",
            "If a CSV model has no in-window update, omit it from the report.",
            "A dated inquiry ledger must contain a terminal outcome for every live CSV target before report generation.",
        ],
        "models": [
            {
                **model,
                "queries": build_model_queries(model, start, end),
            }
            for model in models
        ],
    }


def write_markdown(payload: dict, output_path: Path) -> None:
    lines = [
        "# Weekly AI Model Research Plan",
        "",
        f"Date range: {payload['date_range']['display']}",
        f"Model watchlist source: {payload['source']}",
        f"Models in watchlist: {payload['model_count']}",
        "",
        "## Rules",
        "",
    ]
    lines.extend(f"- {rule}" for rule in payload["rules"])
    lines.extend(["", "## Section Counts", ""])
    for section, count in sorted(payload["section_counts"].items()):
        lines.append(f"- {section}: {count}")

    grouped = grouped_models(payload["models"])
    lines.extend(["", "## Model Targets", ""])
    for section, models in grouped.items():
        lines.extend([f"### {section}", ""])
        for model in models:
            lines.append(f"- {model['name']} ({model['maker']})")
        lines.append("")

    lines.extend(
        [
            "## Query Pattern",
            "",
            "For each model, search the model name and maker with release/update terms and the report date window.",
            "Use exact model names first, then maker-level release notes if exact results are sparse.",
            "",
        ]
    )
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--models-csv", default=DEFAULT_MODELS_CSV)
    parser.add_argument("--week-end", help="Current report Monday/end date in YYYY-MM-DD format.")
    parser.add_argument("--output-dir", default=".")
    args = parser.parse_args()

    models_path = Path(args.models_csv).resolve()
    if not models_path.exists():
        raise FileNotFoundError(f"Models CSV not found: {models_path}")

    end = parse_iso_date(args.week_end) if args.week_end else current_monday()
    start = end - timedelta(days=7)
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    models = load_models(models_path)
    payload = build_payload(models, start, end)

    stem = f"weekly_ai_model_research_plan_{end.isoformat()}"
    json_path = output_dir / f"{stem}.json"
    markdown_path = output_dir / f"{stem}.md"
    json_path.write_text(json.dumps(payload, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    write_markdown(payload, markdown_path)

    print(f"Models loaded: {len(models)}")
    print(f"Research plan JSON: {json_path}")
    print(f"Research plan Markdown: {markdown_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
