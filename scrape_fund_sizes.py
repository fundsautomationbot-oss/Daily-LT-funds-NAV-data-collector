from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import sys

URL = "https://www.swedbank.lt/private/pensions/pillar2/allFunds?language=LIT"

results = []

with sync_playwright() as p:
    print("Starting browser...")
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    print(f"Opening: {URL}")
    page.goto(URL, timeout=60000)
    page.wait_for_timeout(2000)

    # Aggressively dismiss all cookie modal buttons
    print("Dismissing cookie modal...")
    for attempt in range(5):
        try:
            buttons = page.query_selector_all('ui-cookie-consent button')
            for btn in buttons:
                try:
                    btn.click(force=True)
                except Exception:
                    pass
            page.wait_for_timeout(300)
        except Exception:
            pass

    # Wait until fund list loads (using actual table rows, not data-testid)
    print("Waiting for fund rows...")
    page.wait_for_selector("tbody tr", timeout=60000)

    fund_rows = page.query_selector_all("tbody tr")

    print(f"Found {len(fund_rows)} funds")

    if not fund_rows:
        browser.close()
        print("No fund rows found. The page structure may have changed.")
        sys.exit(1)

    for index in range(len(fund_rows)):
        try:
            # Re-select rows every time (page may change)
            fund_rows = page.query_selector_all("tbody tr")
            if index >= len(fund_rows):
                print(f"Row {index} no longer available, stopping.")
                break

            row = fund_rows[index]
            cells = row.query_selector_all("td")
            
            # First cell is checkbox, second is fund name
            if len(cells) < 2:
                continue
            
            fund_name = cells[1].inner_text().strip()
            print(f"[{index+1}/{len(fund_rows)}] Opening fund: {fund_name}")

            # Click the link inside the fund name cell
            link = row.query_selector("a")
            if not link:
                print(f"  No link found, skipping")
                continue
                
            link.click(timeout=15000, force=True)

            # Wait for details panel to appear
            try:
                page.wait_for_selector("text=Fondo dydis", timeout=60000)
            except Exception as e:
                print(f"  Warning: Details panel did not load: {e}")
                page.go_back()
                page.wait_for_selector("tbody tr", timeout=30000)
                continue

            # Find all elements containing "Fondo dydis"
            detail_blocks = page.query_selector_all("div")

            fondo_dydis_value = None
            fondo_dydis_date = None

            for block in detail_blocks:
                try:
                    text = block.inner_text().strip()

                    if text.startswith("Fondo dydis"):
                        # Example: "Fondo dydis (2026-04-16)\n123 456 789 EUR"
                        lines = text.split("\n")

                        if len(lines) >= 2:
                            fondo_dydis_date = lines[0]
                            fondo_dydis_value = lines[1]

                        break
                except Exception:
                    continue

            results.append({
                "Fund name": fund_name,
                "Fondo dydis value": fondo_dydis_value
            })

            # Go back to fund list
            page.go_back()
            page.wait_for_selector("tbody tr", timeout=30000)
            page.wait_for_timeout(1000)

        except Exception as e:
            print(f"  Error processing row {index}: {e}")
            try:
                page.go_back()
                page.wait_for_selector("tbody tr", timeout=30000)
            except Exception:
                pass
            continue

    browser.close()

# Save results
df = pd.DataFrame(results)

if df.empty:
    print("No data was collected from funds.")
    sys.exit(1)

print(f"Collected data for {len(df)} funds")

filename = f"swedbank_fondo_dydis_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
df.to_excel(filename, index=False)

print(f"✅ Saved file: {filename}")