#!/usr/bin/env python3
"""
Luminor pension funds scraper.
Extracts II pillar fund data from a simple static table.
"""

import re
from datetime 
import datetime  from playwright.sync_api 
import sync_playwright

URL = "https://www.luminor.lt/en/pension-funds"

def extract_value(pattern, text): 
    match = re.search(pattern, text, re.IGNORECASE)
return match.group(1).strip() if match 
else None

def scrape_luminor():
    results = []
    with sync_playwright() as p:  
    browser =
    p.chromium.launch(headless=True)  
    page = browser.new_page()

for attempt in range(3):  
    page.goto(URL, timeout=30000) 
try:  
    page.locator("button:has-text('Accept')").click(timeout=3000)
except:  
pass 
try:
    page.wait_for_selector("text=Unit value", timeout=10000)
    break
except:
if attempt == 2:
raise
page.wait_for_timeout (2000)
body = page.inner_text ("body")
date_match = re.search(
r"(?:Unit value date|Vieneto verčiy data)
1d]*(1d{43-10{23-1d{2})",
body
)
nav_date = date_match.group(1) if
date_match else None
blocks = re.split(r"|bFund Luminor|b", body)
for block in blocks [1:]:
text = "Fund Luminor " + block[:600]
fund_match = re.match(r"Fund Luminor|s+
([^\n]+)", text)
if not fund_match:
continue
fund_name = fund_match.group (1).strip()
unit_value = extract_value(
r"Unit value\s+([\d.,]+)",
text
)
nav = extract_value(
r"Net asset value\s+ ([ld\s.,]+)", text
)
if not unit_value or not nav:
continue
results.append ({
"fund_name": fund_name,
"unit_value": unit_value.replace(",", ""),
"nav": nav.replace("' "' '''),
"date": nav_date or
datetime.today().strftime("%Y-%m-%d")
})
browser.close()
return results
if name == "main":
data = scrape_luminor()
for row in data:
print(row)
