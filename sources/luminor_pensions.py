#!/usr/bin/env python3
"""
Luminor pension funds scraper.
Fetches II pillar fund data from server-rendered dnbPensionFunds payload.
"""
import json
import re
import sys
import urllib.request
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent))

from base_scraper import BaseScraper


EXCLUDED_FUNDS = {
    "Luminor ateitis 16–50",
    "Luminor ateitis 50–58",
    "Luminor ateitis 58+",
    "Luminor tvari ateitis index",
    "Luminor ateitis akciju index",
}

LUMINOR_II_FUND_IDS = ["15", "16", "17", "18", "19", "20", "23", "21"]
LUMINOR_URL_TEMPLATE = (
    "https://www.luminor.lt/lt/rinkis-fonda"
    "?fund_type=pension&currency=eur&period=3year&fund={fund_id}"
)


class LuminorPensionsScraper(BaseScraper):
    """Scrapes Luminor II pillar fund metrics from embedded page payload."""

    def __init__(self):
        super().__init__("luminor_pensions")

    def get_url(self) -> str:
        return "https://www.luminor.lt/lt/rinkis-fonda"

    def scrape_data(self, page) -> list:
        # Not used because run() is overridden to avoid brittle browser flows.
        return []

    def fetch_html(self, fund_id: str) -> str:
        url = LUMINOR_URL_TEMPLATE.format(fund_id=fund_id)
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                )
            },
        )
        with urllib.request.urlopen(request, timeout=45) as response:
            return response.read().decode("utf-8", "ignore")

    def extract_payload(self, html: str) -> dict:
        match = re.search(
            r'dnbPensionFunds"\s*:\s*(\{.*?\})\s*,\s*"urlIsAjaxTrusted"',
            html,
            flags=re.DOTALL,
        )
        if not match:
            return {}
        try:
            return json.loads(match.group(1))
        except Exception:
            return {}

    def build_row(self, payload: dict) -> dict:
        fund_rates = payload.get("fundRates") or {}
        fund_history = payload.get("fundRatesHistory") or {}

        fund_name = (
            fund_rates.get("name_alias_lt")
            or fund_rates.get("name_lt")
            or fund_rates.get("name")
            or fund_rates.get("title")
        )
        if not fund_name or fund_name in EXCLUDED_FUNDS:
            return {}

        data_date = None
        if isinstance(fund_history, dict):
            date_keys = [
                key
                for key in fund_history.keys()
                if isinstance(key, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", key)
            ]
            if date_keys:
                data_date = max(date_keys)

        unit_value = fund_rates.get("unit_price_eur") or fund_rates.get("unit_price")
        net_assets = fund_rates.get("nav_eur") or fund_rates.get("nav")
        if unit_value is None and net_assets is None:
            return {}

        return {
            "Fund name": fund_name,
            "Data": data_date,
            "Vieneto vertė": unit_value,
            "Grynieji aktyvai": net_assets,
        }

    def run(self):
        try:
            rows = []
            for fund_id in LUMINOR_II_FUND_IDS:
                try:
                    html = self.fetch_html(fund_id)
                    payload = self.extract_payload(html)
                    if not payload:
                        continue
                    row = self.build_row(payload)
                    if row:
                        rows.append(row)
                except Exception as exc:
                    print(f"Failed fund {fund_id}: {exc}")

            if not rows:
                print("No data scraped from luminor_pensions. Page structure may have changed.")
                return None

            # Deduplicate by fund name to keep latest parsed row only.
            deduped = {}
            for row in rows:
                deduped[row["Fund name"]] = row
            rows = list(deduped.values())

            print(f"Parsed {len(rows)} funds from server payload")
            df = pd.DataFrame(rows)
            if df.empty:
                print("No data parsed for luminor_pensions.")
                return None

            data_date = self._extract_data_date(df)
            filename = f"{self.source_name}_data_{data_date}.xlsx"
            filepath = self.save_to_excel(df, filename)
            if filepath:
                print(f"✅ Excel file created: {filename}")
            return filepath

        except Exception as exc:
            print(f"Error scraping {self.source_name}: {exc}")
            return None


if __name__ == "__main__":
    scraper = LuminorPensionsScraper()
    sys.exit(0 if scraper.run() else 1)
