#!/usr/bin/env python3
"""
Luminor pension funds scraper (II pillar NAV data).
"""

import re
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://www.luminor.lt/en/pension-funds"

EXCLUDED_FUNDS = set()  # add fund names here if needed


def dismiss_cookie_modal(page):
    """Try to close cookie popup if it exists."""
    try:
        page.locator("button:has-text('Accept')").click(timeout=3000)
    except Exception:
        pass


def scrape_luminor():
    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        page.goto(URL, timeout=60000)

        for attempt in range(3):
            dismiss_cookie_modal(page)

            try:
                page.wait_for_selector("text=Unit value", timeout=10000)
                break
            except PlaywrightTimeoutError:
                if attempt == 2:
                    browser.close()
                    raise
                page.wait_for_timeout(2000)

        body = page.inner_text("body")

        # Extract data date
        data_date = None
        date_match = re.search(
            r"(?:Unit value date|Vieneto vertės data)\s*[^\d]*(\d{4}-\d{2}-\d{2})",
            body
        )
        if date_match:
            data_date = date_match.group(1)

        print(f"Data date: {data_date}")

        # Split into fund blocks
        blocks = re.split(r"\bFund Luminor\b", body)

        for block in blocks[1:]:
            text = "Fund Luminor " + block[:800]

            fund_match = re.match(r"Fund Luminor\s+([^\n|]+)", text)
            if not fund_match:
                continue

            fund_name = fund_match.group(1).strip()

            if fund_name in EXCLUDED_FUNDS:
                continue

            unit_value_match = re.search(r"Unit value\s*([\d.,]+)", text)
            nav_match = re.search(r"Net asset value\s*([\d\s.,]+)", text)

            if not unit_value_match or not nav_match:
                continue

            unit_value = unit_value_match.group(1).replace(",", "").strip()
            net_assets = nav_match.group(1).replace(" ", "").replace(",", "").strip()

            results.append({
                "Fund name": fund_name,
                "Date": data_date,
                "Unit value": unit_value,
                "Net assets": net_assets,
            })

        browser.close()

    print(f"Extracted {len(results)} funds")
    return results


if __name__ == "__main__":
    data = scrape_luminor()
    for item in data:
        print(item)
