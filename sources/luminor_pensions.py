#!/usr/bin/env python3
"""
Luminor pension funds scraper.
Extracts II pillar fund data from a simple static table.
"""

import re
from datetime import datetime
from playwright.sync_api import sync_playwright

URL = "https://www.luminor.lt/en/pension-funds"


def extract_value(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def scrape_luminor():
    results = []

    for attempt in range(3):
        self.dismiss_cookie_modal(page)
        try:
            page.wait_for_selector("text=Unit value", timeout=10000)
            break
        except Exception:
            if attempt == 2:
                raise
            page.wait_for_timeout(2000)

    body = page.inner_text("body")

    data_date = None
    m = re.search(
        r"(?:Unit value date|Vieneto vertės data)\s*[^\d]*(\d{4}-\d{2}-\d{2})",
        body
    )
    if m:
        data_date = m.group(1)

    print(f"Data date: {data_date}")

    blocks = re.split(r"\bFund Luminor\b", body)

    for block in blocks[1:]:
        text = "Fund Luminor " + block[:600]

        fund_match = re.match(r"Fund Luminor\s+([^\n|]+)", text)
        if not fund_match:
            continue

        fund_name = fund_match.group(1).strip()

        if fund_name in EXCLUDED_FUNDS:
            continue

        unit_value_match = re.search(r"Unit value\s+([\d.,]+)", text)
        nav_match = re.search(r"Net asset value\s+([\d\s.,]+)", text)

        if not unit_value_match or not nav_match:
            continue

        unit_value = unit_value_match.group(1).replace(",", "")
        net_assets = nav_match.group(1).replace(" ", "").replace(",", "")

        results.append({
            "Fund name": fund_name,
            "Data": data_date,
            "Vieneto vertė": unit_value,
            "Grynieji aktyvai": net_assets,
        })

    print(f"Extracted {len(results)} funds")
    return results
