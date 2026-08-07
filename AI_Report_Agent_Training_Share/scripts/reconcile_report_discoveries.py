import argparse
import csv
import json
from datetime import date
from pathlib import Path

from build_weekly_model_research_plan import (
    build_payload,
    load_models,
    write_markdown,
)


def normalized_key(name: str, maker: str, section: str) -> tuple[str, str, str]:
    return (
        name.strip().casefold(),
        maker.strip().casefold(),
        section.strip().casefold(),
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


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--models-csv", required=True)
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    csv_path = Path(args.models_csv).resolve()
    report = json.loads(input_path.read_text(encoding="utf-8"))
    watchlist = report["model_watchlist"]
    ledger_path = input_path.parent / watchlist["research_ledger"]
    ledger = json.loads(ledger_path.read_text(encoding="utf-8"))

    report_items = {}
    for section in report.get("sections", []):
        for item in section.get("items", []):
            key = (
                (item.get("tool") or "").strip().casefold(),
                (item.get("company") or "").strip().casefold(),
            )
            report_items[key] = item

    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        csv_rows = list(csv.DictReader(handle))
    ledger_keys = {
        normalized_key(
            record.get("name", ""),
            record.get("maker", ""),
            record.get("section", ""),
        )
        for record in ledger.get("records", [])
    }

    additions = []
    for row in csv_rows:
        key = normalized_key(row["Name"], row["Maker"], row["Section"])
        if key in ledger_keys:
            continue
        item = report_items.get(
            (row["Name"].strip().casefold(), row["Maker"].strip().casefold())
        )
        if not item:
            raise ValueError(
                "A new watchlist target is missing from the verified weekly "
                f"report and cannot be reconciled safely: {row['Name']}"
            )
        url = source_url(item)
        if not url:
            raise ValueError(
                f"Verified weekly report item lacks a source URL: {row['Name']}"
            )
        additions.append(
            {
                "name": row["Name"],
                "maker": row["Maker"],
                "section": row["Section"],
                "target_key": (
                    f"{row['Name']} | {row['Maker']} | {row['Section']}"
                ),
                "supplemental_query": (
                    "Verified during weekly newsletter research from a "
                    f"primary source: \"{row['Name']}\" \"{row['Maker']}\""
                ),
                "outcome": "candidate-reviewed-include",
                "results": [
                    {
                        "title": row["Name"],
                        "url": url,
                        "snippet": item.get("summary", ""),
                        "in_window_date_signal": True,
                    }
                ],
                "error": None,
            }
        )

    if additions:
        ledger["records"].extend(additions)
    ledger["live_target_count"] = len(csv_rows)
    ledger["completed_target_count"] = len(csv_rows)
    ledger["candidate_count"] = int(ledger.get("candidate_count") or 0) + len(
        additions
    )
    ledger["search_error_count"] = 0
    ledger_path.write_text(
        json.dumps(ledger, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    end = date.fromisoformat(report["date_range"]["end"])
    start = date.fromisoformat(report["date_range"]["start"])
    models = load_models(csv_path)
    plan = build_payload(models, start, end)
    plan_stem = f"weekly_ai_model_research_plan_{end.isoformat()}"
    plan_json = input_path.parent / f"{plan_stem}.json"
    plan_markdown = input_path.parent / f"{plan_stem}.md"
    plan_json.write_text(
        json.dumps(plan, indent=2, ensure_ascii=True) + "\n",
        encoding="utf-8",
    )
    write_markdown(plan, plan_markdown)

    watchlist["model_count"] = len(csv_rows)
    watchlist["research_plan"] = plan_markdown.name
    report["model_watchlist"] = watchlist
    input_path.write_text(
        json.dumps(report, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    print(f"Live watchlist reconciled: {len(csv_rows)}")
    print(f"Verified report discoveries added to ledger: {len(additions)}")


if __name__ == "__main__":
    main()
