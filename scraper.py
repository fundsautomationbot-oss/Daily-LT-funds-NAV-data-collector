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

    # The page renders fund data in an HTML table.
    print("Waiting for table rows...")
    page.wait_for_selector("tbody tr", timeout=60000)

    rows = page.query_selector_all("tbody tr")
    print(f"Rows found: {len(rows)}")

    if not rows:
        browser.close()
        print("No fund rows found. The page structure may have changed.")
        sys.exit(1)

    for row in rows:
        cells = row.query_selector_all("td")

        # Expected shape: checkbox + 9 data columns.
        if len(cells) < 10:
            continue

        try:
            results.append({
                "Fund name": cells[1].inner_text().strip(),
                "Date": cells[3].inner_text().strip(),
                "GAV": cells[4].inner_text().strip(),
            })
        except Exception:
            continue

    browser.close()

df = pd.DataFrame(results)

if df.empty:
    print("No data parsed from table rows.")
    sys.exit(1)

df["GAV"] = pd.to_numeric(
    df["GAV"].astype(str).str.replace(",", ".", regex=False).str.strip(),
    errors="coerce",
)

filename = f"swedbank_pensions_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
df.to_excel(filename, index=False)

print(f"✅ Excel file created: {filename}")