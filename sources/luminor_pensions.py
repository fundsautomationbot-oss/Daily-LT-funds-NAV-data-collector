#!/usr/bin/env python3
"""
Luminor pension funds scraper.
Fetches II pillar fund data from server-rendered dnbPensionFunds payload.
"""
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import ipaddress
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
        self._current_browser_uses_proxy = None

    def setup_browser(self, use_proxy: bool = True):
        """Initialize browser with Luminor-specific proxy settings."""
        self._current_browser_uses_proxy = use_proxy

        luminor_proxy = os.environ.get("LUMINOR_PROXY_SERVER", "") if use_proxy else ""

        os.environ.pop("PLAYWRIGHT_PROXY_SERVER", None)
        os.environ.pop("PLAYWRIGHT_PROXY_USERNAME", None)
        os.environ.pop("PLAYWRIGHT_PROXY_PASSWORD", None)

        if luminor_proxy:
            os.environ["PLAYWRIGHT_PROXY_SERVER"] = luminor_proxy

            luminor_user = os.environ.get("LUMINOR_PROXY_USERNAME", "")
            luminor_pass = os.environ.get("LUMINOR_PROXY_PASSWORD", "")

            if luminor_user:
                os.environ["PLAYWRIGHT_PROXY_USERNAME"] = luminor_user
            if luminor_pass:
                os.environ["PLAYWRIGHT_PROXY_PASSWORD"] = luminor_pass

        return super().setup_browser()

    def ensure_browser_mode(self, use_proxy: bool):
        if self.page and self._current_browser_uses_proxy == use_proxy:
            return

        self.cleanup_browser()
        self.setup_browser(use_proxy=use_proxy)

    def get_url(self) -> str:
        return LUMINOR_BASE_URLS[0]

    def scrape_data(self, page) -> list:
        return []

    def build_url(self, base_url: str, fund_id: str) -> str:
        return f"{base_url}?fund_type=pension&currency=eur&period=3year&fund={fund_id}"

    def _looks_like_host(self, value: str) -> bool:
        try:
            ipaddress.ip_address(value)
            return True
        except ValueError:
            pass

        return bool(re.match(r"^[A-Za-z0-9.-]+$", value))

    def resolve_http_proxy(self) -> str:
        proxy_server = os.getenv("LUMINOR_PROXY_SERVER", "").strip().strip("'\"")

        if not proxy_server:
            return ""

        username = os.getenv("LUMINOR_PROXY_USERNAME", "").strip().strip("'\"")
        password = os.getenv("LUMINOR_PROXY_PASSWORD", "").strip().strip("'\"")

        # Support common proxy-cheap raw layout: IP:PORT:USERNAME:PASSWORD
        raw_parts = proxy_server.split(":")
        if (
            len(raw_parts) == 4
            and self._looks_like_host(raw_parts[0])
            and raw_parts[1].isdigit()
        ):
            proxy_server = f"{raw_parts[0]}:{raw_parts[1]}"
            if not username:
                username = raw_parts[2]
            if not password:
                password = raw_parts[3]

        # If proxy server is provided as USER:PASS@HOST:PORT, extract credentials.
        if "@" in proxy_server and not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", proxy_server):
            auth_part, host_part = proxy_server.rsplit("@", 1)
            if ":" in auth_part and not username:
                split_auth = auth_part.split(":", 1)
                username = split_auth[0]
                password = split_auth[1]
            proxy_server = host_part

        if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", proxy_server):
            proxy_server = f"http://{proxy_server}"

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

    def is_cloudflare_challenge(self, html: str) -> bool:
        if not html:
            return False

        markers = (
            "Checking your browser before accessing",
            "Just a moment...",
            "cf-chl-",
            "__cf_bm",
            "Cloudflare Ray ID",
            "Attention Required! | Cloudflare",
        )
        return any(marker in html for marker in markers)

    # ✅ ✅ UPDATED: proxy-enabled urllib
    def fetch_html(self, fund_id: str, use_proxy: bool = True) -> str:
        last_error = None
        proxy_url = self.resolve_http_proxy() if use_proxy else ""

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

    def fetch_payload_via_curl(self, fund_id: str, use_proxy: bool = True) -> dict:
        proxy_url = self.resolve_http_proxy() if use_proxy else ""
        last_error = None

        for base_url in LUMINOR_BASE_URLS:
            url = self.build_url(base_url, fund_id)
            command = [
                "curl",
                "-sS",
                "-L",
                "--compressed",
                "--http1.1",
                "--max-time",
                "45",
                "-A",
                (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                "-H",
                "Accept-Language: lt-LT,lt;q=0.9,en-US;q=0.8,en;q=0.7",
                "-H",
                f"Referer: {base_url}",
            ]

            if proxy_url:
                command.extend(["--proxy", proxy_url])

            command.append(url)

            try:
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=55,
                    check=False,
                )
            except Exception as exc:
                last_error = exc
                continue

            if result.returncode != 0:
                last_error = RuntimeError(result.stderr.strip() or f"curl failed with code {result.returncode}")
                continue

            payload = self.extract_payload(result.stdout)
            if payload:
                return payload

            if self.is_cloudflare_challenge(result.stdout):
                last_error = RuntimeError(
                    "Cloudflare challenge page returned by curl"
                    + (" (via proxy)" if proxy_url else " (direct)")
                )

        if last_error:
            raise last_error
        return {}

    def fetch_html_via_browser_navigation(self, fund_id: str, use_proxy: bool = True) -> str:
        self.ensure_browser_mode(use_proxy=use_proxy)

        if not self.page:
            raise RuntimeError("Browser page is not initialized")

        last_error = None

        def accept_cookies_if_visible() -> None:
            if not self.page:
                return

            cookie_selectors = [
                "button:has-text('PRIIMTI VISUS')",
                "button:has-text('Priimti visus')",
                "button:has-text('Accept all')",
                "#onetrust-accept-btn-handler",
            ]

            for selector in cookie_selectors:
                try:
                    button = self.page.locator(selector).first
                    if button.count() and button.is_visible(timeout=800):
                        button.click(timeout=1500)
                        self.page.wait_for_timeout(400)
                        return
                except Exception:
                    continue

        for base_url in LUMINOR_BASE_URLS:
            url = self.build_url(base_url, fund_id)

            try:
                try:
                    homepage = base_url.split("/lt/rinkis-fonda", 1)[0]
                    self.page.goto(homepage + "/lt", wait_until="domcontentloaded", timeout=30000)
                    accept_cookies_if_visible()
                    self.page.wait_for_timeout(800)
                except Exception:
                    pass

                response = self.page.goto(url, wait_until="domcontentloaded", timeout=90000)
                accept_cookies_if_visible()

                if response and response.status >= 400:
                    raise RuntimeError(f"HTTP {response.status}")

                self.page.wait_for_timeout(1500)
                return self.page.content()

            except Exception as exc:
                last_error = exc

        raise last_error

    def extract_payload(self, html: str) -> dict:
        key_match = re.search(r"(?:['\"])?dnbPensionFunds(?:['\"])?\s*:", html)
        if not key_match:
            return {}

        start = html.find("{", key_match.end())
        if start == -1:
            return {}

        depth = 0
        in_string = False
        escaped = False

        for idx in range(start, len(html)):
            ch = html[idx]

            if in_string:
                if escaped:
                    escaped = False
                elif ch == "\\":
                    escaped = True
                elif ch == '"':
                    in_string = False
                continue

            if ch == '"':
                in_string = True
                continue

            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    raw_json = html[start : idx + 1]
                    try:
                        return json.loads(raw_json)
                    except Exception:
                        return {}

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
            cloudflare_fund_ids = []
            used_browser_fallback = False
            payload_missing_ids = []
            proxy_bypass_ids = []

            for fund_id in LUMINOR_II_FUND_IDS:
                try:
                    payload = self.fetch_payload_via_curl(fund_id, use_proxy=True)
                    if not payload:
                        payload = self.fetch_payload_via_curl(fund_id, use_proxy=False)
                        if payload:
                            proxy_bypass_ids.append(fund_id)

                    if payload:
                        row = self.build_row(payload)
                        if row:
                            rows.append(row)
                        continue

                    html = self.fetch_html(fund_id)
                    if self.is_cloudflare_challenge(html):
                        cloudflare_fund_ids.append(fund_id)
                    payload = self.extract_payload(html)

                    if not payload:
                        payload_missing_ids.append(fund_id)
                        if not self.page:
                            self.setup_browser()

                        html = self.fetch_html_via_browser_navigation(fund_id)
                        payload = self.extract_payload(html)
                        if payload:
                            used_browser_fallback = True
                        else:
                            continue

                    row = self.build_row(payload)

                    if row:
                        rows.append(row)

                except urllib.error.HTTPError as exc:
                    if exc.code == 403:
                        try:
                            html = self.fetch_html(fund_id, use_proxy=False)
                            payload = self.extract_payload(html)
                            if payload:
                                row = self.build_row(payload)
                                if row:
                                    rows.append(row)
                                    proxy_bypass_ids.append(fund_id)
                                    continue
                        except Exception:
                            pass

                        try:
                            html = self.fetch_html_via_browser_navigation(fund_id, use_proxy=True)
                            payload = self.extract_payload(html)
                            if payload:
                                row = self.build_row(payload)
                                if row:
                                    rows.append(row)
                                    used_browser_fallback = True
                                    continue
                        except Exception as browser_exc:
                            print(f"Browser fallback failed fund {fund_id}: {browser_exc}")

                        try:
                            html = self.fetch_html_via_browser_navigation(fund_id, use_proxy=False)
                            payload = self.extract_payload(html)
                            if payload:
                                row = self.build_row(payload)
                                if row:
                                    rows.append(row)
                                    used_browser_fallback = True
                                    proxy_bypass_ids.append(fund_id)
                                    continue
                        except Exception as browser_no_proxy_exc:
                            print(f"Browser no-proxy fallback failed fund {fund_id}: {browser_no_proxy_exc}")

                        blocked_fund_ids.append(fund_id)
                        continue

                    print(f"Failed fund {fund_id}: {exc}")

                except Exception as exc:
                    if "Cloudflare challenge" in str(exc):
                        cloudflare_fund_ids.append(fund_id)
                    print(f"Failed fund {fund_id}: {exc}")

            print(f"Parsed {len(rows)} funds from server payload")
            if payload_missing_ids:
                print(
                    "Payload was missing for HTTP fetch on fund IDs: "
                    + ", ".join(payload_missing_ids)
                )
            if used_browser_fallback:
                print("Recovered some funds using browser navigation fallback")
            if proxy_bypass_ids:
                print("Recovered funds by bypassing proxy on IDs: " + ", ".join(proxy_bypass_ids))
            if blocked_fund_ids:
                print("HTTP 403 on fund IDs: " + ", ".join(blocked_fund_ids))
            if cloudflare_fund_ids:
                unique_cf_ids = sorted(set(cloudflare_fund_ids), key=lambda x: int(x))
                print("Cloudflare challenge detected on IDs: " + ", ".join(unique_cf_ids))

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
        finally:
            self.cleanup_browser()


if __name__ == "__main__":
    scraper = LuminorPensionsScraper()
    sys.exit(0 if scraper.run() else 1)