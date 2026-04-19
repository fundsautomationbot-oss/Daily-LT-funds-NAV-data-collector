#!/usr/bin/env python3
"""
Swedbank pension fund scraper - extracts performance metrics.
"""
import sys
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
    
    def scrape_data(self, page) -> list:
        """Extract fund performance data from table."""
        results = []
        
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
                results.append({
                    "Fund name": cells[1].inner_text().strip(),
                    "Date": cells[3].inner_text().strip(),
                    "GAV": cells[4].inner_text().strip(),
                })
            except Exception:
                continue
        
        # Parse numeric columns
        if results:
            df = pd.DataFrame(results)
            df["GAV"] = pd.to_numeric(
                df["GAV"].astype(str).str.replace(",", ".", regex=False).str.strip(),
                errors="coerce",
            )
            results = df.to_dict('records')
        
        return results


if __name__ == "__main__":
    scraper = SwedBankPerformanceScraper()
    scraper.run()
