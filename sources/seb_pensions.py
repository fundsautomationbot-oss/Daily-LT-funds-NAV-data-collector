#!/usr/bin/env python3
"""
SEB pension funds scraper.
Extracts II pillar fund data from a static table.
Table columns: fund name, 1d%, 1m%, 1y%, date (DD.MM), unit value, currency
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from base_scraper import BaseScraper

# Only II pillar funds; skip III pillar
EXCLUDED_PREFIXES = ("SEB index.", "SEB pensija 18+", "SEB pensija 50+", "SEB pensija 58+")
II_PILLAR_FUNDS = {
    "SEB pensija 1961-1967", "SEB pensija 1968-1974", "SEB pensija 1975-1981",
    "SEB pensija 1982-1988", "SEB pensija 1989-1995", "SEB pensija 1996-2002",
    "SEB pensija 2003-2009", "SEB turto išsaugojimo fondas",
}


class SEBPensionsScraper(BaseScraper):
    """Scrapes SEB II pillar pension fund table."""

    URL = "https://e.seb.lt/web/ipank.p?sesskey=&act=VPFOND&filterCode=P&lang=lit&frnam=X&unetmenuhigh="

    def __init__(self):
        super().__init__("seb_pensions")

    def get_url(self) -> str:
        return self.URL

    def get_net_assets(self, page, fund_link_locator) -> str:
        """Click fund link, extract net assets from detail page, return to list."""
        try:
            fund_link_locator.click(timeout=10000)
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(1500)

            body = page.inner_text("body")
            m = re.search(r"Grynųjų aktyvų vertė\s+([\d\s]+)", body)
            value = m.group(1).strip() if m else None

            page.go_back()
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(1000)
            return value
        except Exception as e:
            print(f"  Error getting net assets: {e}")
            try:
                page.go_back()
                page.wait_for_load_state("domcontentloaded")
            except Exception:
                pass
            return None

    def scrape_data(self, page) -> list:
        results = []

        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(2000)

        # Extract year from the last full date on the page (most recent = today's date)
        year = None
        page_date = None
        try:
            body = page.inner_text("body")
            matches = re.findall(r"\d{4}-\d{2}-\d{2}", body)
            if matches:
                page_date = matches[-1]  # last date is today's page timestamp
                year = page_date[:4]
        except Exception:
            pass

        # Find the table containing SEB pension fund rows
        tables = page.query_selector_all("table")
        print(f"  Found {len(tables)} tables")

        target_table = None
        for table in tables:
            txt = table.inner_text()
            if "SEB pensija 1961" in txt or "SEB turto išsaugojimo" in txt:
                target_table = table
                break

        if not target_table:
            print("  Could not find SEB pension fund table.")
            return results

        rows = target_table.query_selector_all("tr")
        print(f"  Rows in target table: {len(rows)}")

        # Pass 1: collect fund name, date, unit value, and detail URL from table
        fund_data = []
        for row in rows:
            cells = row.query_selector_all("td")
            if len(cells) < 6:
                continue

            fund_name = " ".join(cells[1].inner_text().split())
            if not fund_name or fund_name not in II_PILLAR_FUNDS:
                continue

            date_raw = cells[5].inner_text().strip()
            data_date = None
            if year and re.match(r"\d{2}\.\d{2}", date_raw):
                day, month = date_raw.split(".")[:2]
                data_date = f"{year}-{month}-{day}"
                # Table shows DD.MM only: early-January page with late-December NAVs.
                if page_date and data_date > page_date:
                    data_date = f"{int(year) - 1}-{month}-{day}"

            unit_value = cells[6].inner_text().strip()

            link = row.query_selector("a")
            detail_href = link.get_attribute("href") if link else None

            fund_data.append({
                "Fund name": fund_name,
                "Data": data_date,
                "Vieneto vertė": unit_value,
                "detail_href": detail_href,
            })

        # The output filename uses the latest row date, so a table where SEB has
        # updated only some funds would be published as a complete snapshot.
        # Refuse to save until every II pillar fund is present with the same NAV date.
        found_names = {fund["Fund name"] for fund in fund_data}
        missing_funds = sorted(II_PILLAR_FUNDS - found_names)
        nav_dates = {fund["Data"] for fund in fund_data}
        if missing_funds or None in nav_dates or len(nav_dates) != 1:
            print(
                "  Incomplete SEB snapshot, not saving: "
                f"missing funds={missing_funds}, NAV dates={sorted(str(d) for d in nav_dates)}"
            )
            for fund in fund_data:
                print(f"    {fund['Fund name']}: {fund['Data']}")
            return []

        # Pass 2: visit each detail page to get net assets
        base_url = "https://e.seb.lt/web/ipank.p"
        for idx, fund in enumerate(fund_data, start=1):
            print(f"  [{idx}/{len(fund_data)}] Getting net assets for {fund['Fund name']}...")
            net_assets = None
            try:
                detail_link = page.locator(f"a:has-text('{fund['Fund name']}')").first
                detail_link.click(timeout=10000)
                page.wait_for_load_state("domcontentloaded")
                page.wait_for_timeout(1200)

                body = page.inner_text("body")
                m = re.search(r"Grynųjų aktyvų vertė\s+([\d\s]+)", body)
                if m:
                    net_assets = m.group(1).strip()

                page.go_back()
                page.wait_for_load_state("domcontentloaded")
                page.wait_for_timeout(800)
            except Exception as e:
                print(f"    Error: {e}")
                try:
                    page.go_back()
                    page.wait_for_load_state("domcontentloaded")
                except Exception:
                    pass

            results.append({
                "Fund name": fund["Fund name"],
                "Data": fund["Data"],
                "Vieneto vertė": fund["Vieneto vertė"],
                "Grynieji aktyvai": net_assets,
            })

        return results


if __name__ == "__main__":
    scraper = SEBPensionsScraper()
    sys.exit(0 if scraper.run() else 1)
