#!/usr/bin/env python3
"""
Luminor pension funds scraper.
Extracts II pillar fund data and writes an XLSX via BaseScraper.
"""
import re
import sys
from pathlib import Path

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
EN_FALLBACK_URL = "https://www.luminor.lt/en/pension-funds"


class LuminorPensionsScraper(BaseScraper):
    """Scrapes Luminor II pillar pension fund data from multiple page layouts."""

    def __init__(self):
        super().__init__("luminor_pensions")

    def setup_browser(self):
        """Use a realistic browser fingerprint for Luminor pages."""
        from playwright.sync_api import sync_playwright

        self._playwright = sync_playwright().start()
        self.browser = self._playwright.chromium.launch(
            headless=self._is_headless(),
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
                "--disable-gpu",
            ],
        )
        self.context = self.browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="lt-LT",
            timezone_id="Europe/Vilnius",
        )
        self.page = self.context.new_page()
        try:
            self.page.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
        except Exception:
            pass
        return self.page

    def get_url(self) -> str:
        return "https://www.luminor.lt/lt/rinkis-fonda"

    def dismiss_cookie_modal(self, page):
        selectors = [
            "button:has-text('PRIIMTI VISUS')",
            "button:has-text('Priimti visus')",
            "button:has-text('Accept all')",
            "button:has-text('Accept')",
            "#onetrust-accept-btn-handler",
        ]
        for selector in selectors:
            try:
                page.locator(selector).first.click(timeout=3000, force=True)
                break
            except Exception:
                pass
        page.wait_for_timeout(400)

    def extract_fund_ids(self, page) -> list:
        """Discover currently available II-pillar fund IDs from settings or links."""
        # Primary source in current Luminor page: drupal settings payload.
        try:
            ids_from_settings = page.evaluate(
                '''() => {
                    const settings = window.drupalSettings || (window.Drupal && window.Drupal.settings) || {};
                    const dnb = settings.dnbPensionFunds || {};
                    const groups = dnb.funds || {};
                    const values = new Set();

                    const visitGroup = (group) => {
                        if (!group || typeof group !== 'object') return;
                        for (const entry of Object.values(group)) {
                            const type = (entry.type || '').toLowerCase();
                            const id = String(entry.fund_id || '');
                            if (type === 'pension' && id) values.add(id);
                        }
                    };

                    if (Array.isArray(groups)) {
                        for (const group of groups) visitGroup(group);
                    } else if (groups && typeof groups === 'object') {
                        for (const group of Object.values(groups)) visitGroup(group);
                    }

                    return Array.from(values);
                }''',
            )
            if isinstance(ids_from_settings, list) and ids_from_settings:
                return [str(item) for item in ids_from_settings]
        except Exception:
            pass

        # Fallback source: links in rendered DOM.
        try:
            ids = page.evaluate(
                '''() => {
                    const values = new Set();
                    const re = /[?&]fund=(\\d+)/;
                    const nodes = document.querySelectorAll("a[href*='fund_type=pension'][href*='fund=']");
                    for (const node of nodes) {
                        const href = node.getAttribute('href') || '';
                        const match = href.match(re);
                        if (match) values.add(match[1]);
                    }
                    return Array.from(values);
                }''',
            )
            if isinstance(ids, list) and ids:
                return [str(item) for item in ids]
        except Exception:
            pass
        return []

    def extract_fund_payload(self, page) -> tuple:
        """Read fund payload from modern drupalSettings or legacy Drupal.settings."""
        payload = page.evaluate(
            '''() => {
                const legacy = window.Drupal && window.Drupal.settings ? window.Drupal.settings : null;
                const modern = window.drupalSettings || null;
                const settings = modern || legacy || {};
                const dnb = settings.dnbPensionFunds || null;
                if (!dnb) return { rates: null, history: null };
                return {
                    rates: dnb.fundRates || dnb.fund_rates || null,
                    history: dnb.fundRatesHistory || dnb.fund_rates_history || null,
                };
            }''',
        )
        if not isinstance(payload, dict):
            return None, None
        return payload.get("rates"), payload.get("history")

    def get_selected_fund_id(self, page) -> str:
        try:
            selected = page.evaluate(
                '''() => {
                    const settings = window.drupalSettings || (window.Drupal && window.Drupal.settings) || {};
                    const dnb = settings.dnbPensionFunds || {};
                    const defaults = dnb.defaultValues || {};
                    return defaults.fund ? String(defaults.fund) : '';
                }''',
            )
            return str(selected) if selected else ""
        except Exception:
            return ""

    def parse_text_fallback(self, page) -> list:
        """Fallback parser for the english text-heavy page layout."""
        results = []
        body = page.inner_text("body")

        data_date = None
        date_match = re.search(
            r"(?:Unit value date|Vieneto vertes data)\s*[^\d]*(\d{4}[-.]\d{2}[-.]\d{2})",
            body,
            flags=re.IGNORECASE,
        )
        if date_match:
            data_date = date_match.group(1).replace(".", "-")

        blocks = re.split(r"\bFund Luminor\b", body)
        for block in blocks[1:]:
            text = "Fund Luminor " + block[:900]
            fund_match = re.match(r"Fund Luminor\s+([^\n|]+)", text)
            if not fund_match:
                continue

            fund_name = fund_match.group(1).strip()
            if not fund_name or fund_name in EXCLUDED_FUNDS:
                continue

            unit_value_match = re.search(r"Unit value\s*([\d.,]+)", text, flags=re.IGNORECASE)
            nav_match = re.search(r"Net asset value\s*([\d\s.,]+)", text, flags=re.IGNORECASE)
            if not unit_value_match and not nav_match:
                continue

            unit_value = unit_value_match.group(1).strip() if unit_value_match else None
            net_assets = nav_match.group(1).strip() if nav_match else None

            results.append(
                {
                    "Fund name": fund_name,
                    "Data": data_date,
                    "Vieneto vertė": unit_value,
                    "Grynieji aktyvai": net_assets,
                }
            )

        return results

    def scrape_data(self, page) -> list:
        results = []

        # 1) Primary route: iterate fund detail views and read drupal settings payload.
        try:
            self.dismiss_cookie_modal(page)
            page.wait_for_timeout(1000)

            discovered_ids = self.extract_fund_ids(page)
            fund_ids = discovered_ids if discovered_ids else LUMINOR_II_FUND_IDS
            if discovered_ids:
                print(f"  Discovered {len(discovered_ids)} fund IDs from live page")
            else:
                print("  Using fallback known Luminor fund IDs")

            for fund_id in fund_ids:
                fund_url = f"{self.get_url()}?fund_type=pension&fund={fund_id}&currency=eur&period=3year"
                page.goto(fund_url, wait_until="domcontentloaded", timeout=90000)
                self.dismiss_cookie_modal(page)
                try:
                    page.wait_for_function(
                        f'''() => {{
                            const settings = window.drupalSettings || (window.Drupal && window.Drupal.settings) || {{}};
                            const dnb = settings.dnbPensionFunds || {{}};
                            const defaults = dnb.defaultValues || {{}};
                            return String(defaults.fund || '') === '{fund_id}';
                        }}''',
                        timeout=7000,
                    )
                except Exception:
                    pass
                page.wait_for_timeout(700)

                selected_fund_id = self.get_selected_fund_id(page)
                if selected_fund_id and selected_fund_id != str(fund_id):
                    # When CI ignores query params and keeps default fund, skip this pass.
                    continue

                fund_rates, fund_history = self.extract_fund_payload(page)
                if not isinstance(fund_rates, dict):
                    continue

                fund_name = (
                    fund_rates.get("name_alias_lt")
                    or fund_rates.get("name_lt")
                    or fund_rates.get("name")
                    or fund_rates.get("title")
                )
                if not fund_name or fund_name in EXCLUDED_FUNDS:
                    continue

                data_date = None
                if isinstance(fund_history, dict):
                    date_keys = [
                        key
                        for key in fund_history.keys()
                        if isinstance(key, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", key)
                    ]
                    if date_keys:
                        data_date = max(date_keys)
                if not data_date:
                    candidate_date = fund_rates.get("date") or fund_rates.get("updated_at")
                    if isinstance(candidate_date, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", candidate_date):
                        data_date = candidate_date

                unit_value = fund_rates.get("unit_price_eur") or fund_rates.get("unit_price")
                net_assets = fund_rates.get("nav_eur") or fund_rates.get("nav")
                if unit_value is None and net_assets is None:
                    continue

                results.append(
                    {
                        "Fund name": fund_name,
                        "Data": data_date,
                        "Vieneto vertė": unit_value,
                        "Grynieji aktyvai": net_assets,
                    }
                )

            # Remove duplicates if the same fund appears across multiple passes.
            if results:
                deduped = {}
                for row in results:
                    deduped[row["Fund name"]] = row
                results = list(deduped.values())

            if results:
                print(f"  Parsed {len(results)} funds from live selector payload")
                return results
        except Exception as exc:
            print(f"  Selector payload parse failed: {exc}")

        # 2) Legacy table fallback.
        rows = []
        table_selector = 'table[aria-describedby="funds-table-label"] tbody tr'
        for attempt in range(1, 4):
            page.wait_for_load_state("domcontentloaded")
            self.dismiss_cookie_modal(page)
            try:
                page.wait_for_selector(table_selector, timeout=20000)
            except Exception:
                pass
            rows = page.query_selector_all(table_selector)
            print(f"  Legacy table attempt {attempt}: found {len(rows)} rows")
            if len(rows) >= 6:
                break
            if attempt < 3:
                page.wait_for_timeout(1500)
                try:
                    page.reload(wait_until="domcontentloaded", timeout=60000)
                except Exception:
                    pass

        if rows:
            data_date = None
            try:
                body = page.inner_text("body")
                match = re.search(r"Vieneto verciu data[:\s]+(\d{4})[.-](\d{2})[.-](\d{2})", body)
                if match:
                    data_date = f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
            except Exception:
                pass

            for row in rows:
                cells = row.query_selector_all("td")
                if len(cells) < 4:
                    continue

                fund_name = cells[0].inner_text().strip()
                if not fund_name or fund_name in EXCLUDED_FUNDS:
                    continue

                unit_value = cells[1].inner_text().strip().replace("EUR", "").strip()
                net_assets = cells[3].inner_text().strip()

                results.append(
                    {
                        "Fund name": fund_name,
                        "Data": data_date,
                        "Vieneto vertė": unit_value,
                        "Grynieji aktyvai": net_assets,
                    }
                )

            if results:
                print(f"  Parsed {len(results)} funds from legacy table")
                return results

        # 3) English text page fallback.
        try:
            page.goto(EN_FALLBACK_URL, wait_until="domcontentloaded", timeout=90000)
            self.dismiss_cookie_modal(page)
            text_results = self.parse_text_fallback(page)
            if text_results:
                print(f"  Parsed {len(text_results)} funds from english text fallback")
                return text_results
        except Exception as exc:
            print(f"  English fallback parse failed: {exc}")

        return results


if __name__ == "__main__":
    scraper = LuminorPensionsScraper()
    sys.exit(0 if scraper.run() else 1)
