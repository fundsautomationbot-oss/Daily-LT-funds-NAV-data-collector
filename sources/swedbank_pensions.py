#!/usr/bin/env python3
"""
Swedbank pension fund scraper - extracts performance metrics.
"""
import sys
import re
from pathlib import Path

# Add parent directory to path so we can import base_scraper
sys.path.insert(0, str(Path(__file__).parent.parent))

from base_scraper import BaseScraper
import pandas as pd


class SwedBankPerformanceScraper(BaseScraper):
    """Scrapes Swedbank pension fund performance data."""
    
    def __init__(self):
        super().__init__("swedbank_pensions")
    
    def get_url(self) -> str:
        return "https://www.swedbank.lt/private/pensions/pillar2/allFunds?language=LIT"

    def normalize_date_text(self, value: str) -> str:
        raw = " ".join(str(value).split())
        if not raw:
            return ""

        # Support YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD, YYYY MM DD.
        m = re.search(r"(\d{4})[\s./-](\d{2})[\s./-](\d{2})", raw)
        if m:
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

        # Support DD-MM-YYYY, DD.MM.YYYY, DD/MM/YYYY, DD MM YYYY.
        m = re.search(r"(\d{2})[\s./-](\d{2})[\s./-](\d{4})", raw)
        if m:
            return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"

        return ""

    def extract_page_report_date(self, page) -> str:
        try:
            body_text = " ".join(page.inner_text("body").split())
        except Exception:
            return ""

        # Prefer dates that appear near labels like Data/Atnaujinta.
        preferred = [
            r"(?:Data|Atnaujinta|Paskutin(?:is|ė)\s+atnaujinimas)\D{0,40}(\d{4}[\s./-]\d{2}[\s./-]\d{2})",
            r"(?:Data|Atnaujinta|Paskutin(?:is|ė)\s+atnaujinimas)\D{0,40}(\d{2}[\s./-]\d{2}[\s./-]\d{4})",
        ]
        for pattern in preferred:
            match = re.search(pattern, body_text, flags=re.IGNORECASE)
            if match:
                normalized = self.normalize_date_text(match.group(1))
                if normalized:
                    return normalized

        fallback = re.search(r"(\d{4}[\s./-]\d{2}[\s./-]\d{2})", body_text)
        if fallback:
            return self.normalize_date_text(fallback.group(1))

        return ""

    def dismiss_cookie_modal(self, page):
        """Try to dismiss cookie consent so row links are clickable."""
        for _ in range(5):
            try:
                buttons = page.query_selector_all("ui-cookie-consent button")
                for btn in buttons:
                    try:
                        btn.click(force=True)
                    except Exception:
                        pass
                page.wait_for_timeout(300)
            except Exception:
                pass

    def extract_fondo_dydis(self, text: str):
        """Extract numeric fund-size value from details page text."""
        pattern = re.compile(
            r"Fondo\s+dydis\s*\([^)]+\):?\s*([0-9\s\u00A0.,]+)\s*EUR",
            re.IGNORECASE,
        )
        match = pattern.search(text)
        if not match:
            return None
        return " ".join(match.group(1).replace("\u00A0", " ").split()) + " EUR"

    def collect_fund_sizes(self, page, fund_names):
        """Open each fund details page and capture Fondo dydis value."""
        sizes = {}
        for idx, fund_name in enumerate(fund_names, start=1):
            try:
                print(f"  [{idx}/{len(fund_names)}] Getting fund size for {fund_name}...")

                # Re-acquire table rows after each navigation.
                page.wait_for_selector("tbody tr", timeout=60000)
                rows = page.query_selector_all("tbody tr")
                target_row = None
                for row in rows:
                    cells = row.query_selector_all("td")
                    if len(cells) < 2:
                        continue
                    row_name = " ".join(cells[1].inner_text().split())
                    if row_name == fund_name:
                        target_row = row
                        break

                if not target_row:
                    print(f"    Row not found for {fund_name}")
                    continue

                link = target_row.query_selector("a")
                if not link:
                    print(f"    No details link found for {fund_name}")
                    continue

                link.click(force=True, timeout=15000)
                page.wait_for_url("**/list/details", timeout=30000)
                page.wait_for_load_state("networkidle", timeout=30000)

                body_text = page.inner_text("body")
                sizes[fund_name] = self.extract_fondo_dydis(body_text)

                page.go_back(timeout=30000)
                page.wait_for_load_state("domcontentloaded", timeout=30000)
                self.dismiss_cookie_modal(page)
            except Exception:
                try:
                    page.goto(self.get_url(), timeout=60000)
                    page.wait_for_load_state("domcontentloaded", timeout=30000)
                    self.dismiss_cookie_modal(page)
                except Exception:
                    pass

        return sizes
    
    def scrape_data(self, page) -> list:
        """Extract fund performance data from table."""
        results = []
        page_report_date = self.extract_page_report_date(page)
        
        self.dismiss_cookie_modal(page)
        print("Waiting for table rows...")
        page.wait_for_selector("tbody tr", timeout=60000)
        
        rows = page.query_selector_all("tbody tr")
        print(f"Rows found: {len(rows)}")
        
        for row in rows:
            cells = row.query_selector_all("td")
            
            # Expected shape: checkbox + 9 data columns.
            if len(cells) < 10:
                continue
            
            try:
                fund_name = cells[1].inner_text().strip()
                if not fund_name:
                    continue
                if "tradicin" in fund_name.lower():
                    continue

                # Date column position can shift; try the expected cell first, then scan row.
                row_date = ""
                if len(cells) > 3:
                    row_date = self.normalize_date_text(cells[3].inner_text())
                if not row_date:
                    for cell in cells:
                        row_date = self.normalize_date_text(cell.inner_text())
                        if row_date:
                            break
                if not row_date:
                    row_date = page_report_date

                results.append({
                    "Fund name": " ".join(fund_name.split()),
                    "Date": row_date,
                    "GAV": cells[4].inner_text().strip(),
                })

            except Exception:
                continue
        
        # Parse numeric columns
        if results:
            fund_sizes = self.collect_fund_sizes(page, [item["Fund name"] for item in results])
            for item in results:
                item["Fondo dydis value"] = fund_sizes.get(item["Fund name"])

            df = pd.DataFrame(results)
            df["GAV"] = pd.to_numeric(
                df["GAV"].astype(str).str.replace(",", ".", regex=False).str.strip(),
                errors="coerce",
            )
            results = df.to_dict('records')
        
        return results


if __name__ == "__main__":
    scraper = SwedBankPerformanceScraper()
    sys.exit(0 if scraper.run() else 1)
