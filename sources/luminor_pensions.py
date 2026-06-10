#!/usr/bin/env python3
"""
Luminor pension funds scraper.
Fetches II pillar fund data from server-rendered dnbPensionFunds payload.
"""
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
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
LUMINOR_BASE_URLS = [
    "https://luminor.lt/lt/rinkis-fonda",
    "https://www.luminor.lt/lt/rinkis-fonda",
]


class LuminorPensionsScraper(BaseScraper):

    def __init__(self):
        super().__init__("luminor_pensions")

    def setup_browser(self):
        """Initialize browser with Luminor-specific proxy settings."""
        luminor_proxy = os.environ.get("LUMINOR_PROXY_SERVER", "")

        if luminor_proxy:
            os.environ["PLAYWRIGHT_PROXY_SERVER"] = luminor_proxy

            luminor_user = os.environ.get("LUMINOR_PROXY_USERNAME", "")
            luminor_pass = os.environ.get("LUMINOR_PROXY_PASSWORD", "")

            if luminor_user:
                os.environ["PLAYWRIGHT_PROXY_USERNAME"] = luminor_user
            if luminor_pass:
                os.environ["PLAYWRIGHT_PROXY_PASSWORD"] = luminor_pass

        return super().setup_browser()

    def get_url(self) -> str:
        return LUMINOR_BASE_URLS[0]

    def scrape_data(self, page) -> list:
        return []

    def build_url(self, base_url: str, fund_id: str) -> str:
        return f"{base_url}?fund_type=pension&currency=eur&period=3year&fund={fund_id}"

    def resolve_http_proxy(self) -> str:
        explicit_proxy = os.getenv("LUMINOR_HTTP_PROXY", "").strip()
        if explicit_proxy:
            return explicit_proxy

        proxy_server = os.getenv("LUMINOR_PROXY_SERVER", "").strip()
        if not proxy_server:
            return ""

        if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", proxy_server):
            proxy_server = f"http://{proxy_server}"

        username = os.getenv("LUMINOR_PROXY_USERNAME", "").strip()
        password = os.getenv("LUMINOR_PROXY_PASSWORD", "").strip()
        if not username:
            return proxy_server

        parsed = urllib.parse.urlsplit(proxy_server)
        if "@" in parsed.netloc:
            return proxy_server

        auth = urllib.parse.quote(username, safe="")
        if password:
            auth += ":" + urllib.parse.quote(password, safe="")

        netloc = f"{auth}@{parsed.netloc}"
        return urllib.parse.urlunsplit((parsed.scheme, netloc, parsed.path, parsed.query, parsed.fragment))

    def is_retryable_browser_error(self, error_message: str) -> bool:
        retryable_markers = (
            "chrome-error://chromewebdata",
            "ERR_PROXY",
            "ERR_CONNECTION",
            "ERR_TUNNEL",
            "ERR_TIMED_OUT",
            "ERR_NETWORK_CHANGED",
            "Navigation timeout",
        )
        return any(marker in error_message for marker in retryable_markers)

    # ✅ ✅ UPDATED: proxy-enabled urllib
    def fetch_html(self, fund_id: str) -> str:
        last_error = None
        proxy_url = self.resolve_http_proxy()

        for base_url in LUMINOR_BASE_URLS:
            url = self.build_url(base_url, fund_id)

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

            try:
                # ✅ ✅ PROXY LOGIC
                if proxy_url:
                    proxy_handler = urllib.request.ProxyHandler({
                        "http": proxy_url,
                        "https": proxy_url,
                    })
                    opener = urllib.request.build_opener(proxy_handler)
                else:
                    opener = urllib.request.build_opener()

                with opener.open(request, timeout=45) as response:
                    return response.read().decode("utf-8", "ignore")

            except Exception as exc:
                last_error = exc

            # ✅ optional delay (prevents blocking)
            time.sleep(1.5)

        raise last_error

    def fetch_html_via_browser_navigation(self, fund_id: str) -> str:
        if not self.page:
            raise RuntimeError("Browser page is not initialized")

        last_error = None

        for base_url in LUMINOR_BASE_URLS:
            url = self.build_url(base_url, fund_id)

            try:
                try:
                    homepage = base_url.split("/lt/rinkis-fonda", 1)[0]
                    self.page.goto(homepage + "/lt", wait_until="domcontentloaded", timeout=30000)
                    self.page.wait_for_timeout(800)
                except Exception:
                    pass

                response = self.page.goto(url, wait_until="domcontentloaded", timeout=90000)

                if response and response.status >= 400:
                    raise RuntimeError(f"HTTP {response.status}")

                self.page.wait_for_timeout(1500)
                return self.page.content()

            except Exception as exc:
                last_error = exc

        raise last_error

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
            blocked_fund_ids = []

            for fund_id in LUMINOR_II_FUND_IDS:
                try:
                    html = self.fetch_html(fund_id)
                    payload = self.extract_payload(html)

                    if not payload:
                        continue

                    row = self.build_row(payload)

                    if row:
                        rows.append(row)

                except urllib.error.HTTPError as exc:
                    if exc.code == 403:
                        blocked_fund_ids.append(fund_id)
                        continue

                    print(f"Failed fund {fund_id}: {exc}")

                except Exception as exc:
                    print(f"Failed fund {fund_id}: {exc}")

            print(f"Parsed {len(rows)} funds from server payload")

            df = pd.DataFrame(rows)

            if df.empty:
                print("No data parsed.")
                return None

            data_date = self._extract_data_date(df)
            filename = f"{self.source_name}_data_{data_date}.xlsx"
            filepath = self.save_to_excel(df, filename)

            print(f"✅ Excel created: {filename}")

            return filepath

        except Exception as exc:
            print(f"Error scraping: {exc}")
            return None


if __name__ == "__main__":
    scraper = LuminorPensionsScraper()
    sys.exit(0 if scraper.run() else 1)