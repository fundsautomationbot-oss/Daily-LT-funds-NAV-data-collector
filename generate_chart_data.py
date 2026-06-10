#!/usr/bin/env python3
"""
Parse all pension_data_combined_*.html reports from docs/ and output
docs/chart_data.json for the interactive dashboard.
"""
import json
import re
import sys
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path

DOCS_DIR = Path(__file__).parent / "docs"
OUTPUT_FILE = DOCS_DIR / "chart_data.json"

DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")
BUCKET_RE = re.compile(r"(20\d{2}|19\d{2})[-–](20\d{2}|19\d{2})")
PRESERVATION_KEYWORDS = ["išsaugojimo", "preservation", "turto išsaugojimo"]

PROVIDER_COLORS = {
    "ALLIANZ":  "#2563EB",
    "ARTEA":    "#059669",
    "GOINDEX":  "#7C3AED",
    "LUMINOR":  "#EA580C",
    "SEB":      "#DC2626",
    "SWEDBANK": "#D97706",
}

BUCKET_ORDER = [
    "2003-2009",
    "1996-2002",
    "1989-1995",
    "1982-1988",
    "1975-1981",
    "1968-1974",
    "1961-1967",
    "preservation",
]


def extract_bucket(fund_name: str) -> str:
    name_lower = fund_name.lower()
    for kw in PRESERVATION_KEYWORDS:
        if kw in name_lower:
            return "preservation"
    m = BUCKET_RE.search(fund_name)
    if m:
        start = m.group(1)
        end = m.group(2)
        return f"{start}-{end}"
    return "other"


class TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows: list[list[str]] = []
        self.current_row: list[str] = []
        self.in_td = False
        self.in_provider_row = False
        self.current_provider: str | None = None

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        if tag == "tr":
            cls = attrs_dict.get("class", "")
            self.in_provider_row = "provider" in cls
        if tag in ("td", "th"):
            self.in_td = True

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self.in_td = False
        if tag == "tr" and self.current_row:
            if self.in_provider_row and len(self.current_row) == 1:
                self.current_provider = self.current_row[0].strip().upper()
            elif not self.in_provider_row and len(self.current_row) >= 4:
                self.rows.append([self.current_provider] + self.current_row)
            self.current_row = []

    def handle_data(self, data):
        if self.in_td:
            self.current_row.append(data.strip())


def parse_nav(raw: str) -> float | None:
    try:
        return float(raw.replace(",", "").replace(" ", ""))
    except (ValueError, AttributeError):
        return None


def parse_assets(raw: str) -> int | None:
    try:
        return int(raw.replace(",", "").replace(" ", "").replace("\xa0", ""))
    except (ValueError, AttributeError):
        return None


def parse_report(html_path: Path) -> list[dict]:
    """Return list of {provider, name, date, nav, assets} dicts."""
    parser = TableParser()
    parser.feed(html_path.read_text(encoding="utf-8"))
    records = []
    for row in parser.rows:
        if len(row) < 5:
            continue
        provider, name, date, nav_raw, assets_raw = row[0], row[1], row[2], row[3], row[4]
        if not DATE_RE.match(date):
            continue
        nav = parse_nav(nav_raw)
        assets = parse_assets(assets_raw)
        if nav is None:
            continue
        records.append(
            {
                "provider": provider or "UNKNOWN",
                "name": name,
                "date": date,
                "nav": nav,
                "assets": assets,
                "bucket": extract_bucket(name),
            }
        )
    return records


def build_chart_data(docs_dir: Path) -> dict:
    all_records: list[dict] = []
    report_files = sorted(docs_dir.glob("pension_data_combined_*.html"))
    if not report_files:
        print("No report files found in docs/")
        return {}

    for path in report_files:
        m = DATE_RE.search(path.name)
        if not m:
            continue
        records = parse_report(path)
        all_records.extend(records)
        print(f"  Parsed {path.name}: {len(records)} rows")

    # Build fund registry (keyed by provider+name)
    fund_map: dict[str, dict] = {}
    dates_set: set[str] = set()
    providers_set: set[str] = set()

    for r in all_records:
        key = f"{r['provider']}||{r['name']}"
        if key not in fund_map:
            fund_map[key] = {
                "provider": r["provider"],
                "name": r["name"],
                "bucket": r["bucket"],
                "nav": {},
                "assets": {},
            }
        fund_map[key]["nav"][r["date"]] = r["nav"]
        if r["assets"] is not None:
            fund_map[key]["assets"][r["date"]] = r["assets"]
        dates_set.add(r["date"])
        providers_set.add(r["provider"])

    dates = sorted(dates_set)
    providers = sorted(providers_set)
    funds = sorted(
        fund_map.values(),
        key=lambda f: (
            providers.index(f["provider"]) if f["provider"] in providers else 99,
            BUCKET_ORDER.index(f["bucket"]) if f["bucket"] in BUCKET_ORDER else 99,
        ),
    )

    return {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "dates": dates,
        "providers": providers,
        "providerColors": PROVIDER_COLORS,
        "bucketOrder": BUCKET_ORDER,
        "funds": funds,
    }


def main():
    print("Generating chart_data.json...")
    data = build_chart_data(DOCS_DIR)
    if not data:
        sys.exit(1)
    OUTPUT_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ Written {OUTPUT_FILE}  ({len(data['funds'])} funds, {len(data['dates'])} dates)")


if __name__ == "__main__":
    main()
