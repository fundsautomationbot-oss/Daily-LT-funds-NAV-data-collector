#!/usr/bin/env python3
"""
Luminor pension funds scraper.
Fetches II pillar fund data from server-rendered dnbPensionFunds payload.
"""
import json
import html
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
LUMINOR_TABLE_URLS = [
    "https://luminor.lt/lt/pensiju-fondai",
    "https://www.luminor.lt/lt/pensiju-fondai",
]
DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")


class LuminorPensionsScraper(BaseScraper):

    def __init__(self):
        super().__init__("luminor_pensions")
        self._current_browser_uses_proxy = None

    def setup_browser(self, use_proxy: bool = True):
        """Initialize browser with Luminor-specific proxy settings."""
        self._current_browser_uses_proxy = use_proxy

        luminor_proxy = self.resolve_playwright_proxy_server() if use_proxy else ""

        os.environ.pop("PLAYWRIGHT_PROXY_SERVER", None)
        os.environ.pop("PLAYWRIGHT_PROXY_USERNAME", None)
        os.environ.pop("PLAYWRIGHT_PROXY_PASSWORD", None)

        if luminor_proxy:
            os.environ["PLAYWRIGHT_PROXY_SERVER"] = luminor_proxy

            luminor_user = self.resolve_proxy_username()
            luminor_pass = self.resolve_proxy_password()

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

    def build_table_url(self, base_url: str) -> str:
        return base_url

    def _accept_cookies_if_visible(self) -> None:
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

    def _looks_like_host(self, value: str) -> bool:
        try:
            ipaddress.ip_address(value)
            return True
        except ValueError:
            pass

        return bool(re.match(r"^[A-Za-z0-9.-]+$", value))

    def resolve_proxy_username(self) -> str:
        username = os.getenv("LUMINOR_PROXY_USERNAME", "").strip().strip("'\"")
        if username:
            return username

        proxy_server = os.getenv("LUMINOR_PROXY_SERVER", "").strip().strip("'\"")
        if "@" in proxy_server and ":" in proxy_server.split("@", 1)[0]:
            return proxy_server.split("@", 1)[0].split(":", 1)[0]

        raw_parts = proxy_server.split(":")
        if (
            len(raw_parts) == 4
            and self._looks_like_host(raw_parts[0])
            and raw_parts[1].isdigit()
        ):
            return raw_parts[2]

        return ""

    def resolve_proxy_password(self) -> str:
        password = os.getenv("LUMINOR_PROXY_PASSWORD", "").strip().strip("'\"")
        if password:
            return password

        proxy_server = os.getenv("LUMINOR_PROXY_SERVER", "").strip().strip("'\"")
        if "@" in proxy_server and ":" in proxy_server.split("@", 1)[0]:
            return proxy_server.split("@", 1)[0].split(":", 1)[1]

        raw_parts = proxy_server.split(":")
        if (
            len(raw_parts) == 4
            and self._looks_like_host(raw_parts[0])
            and raw_parts[1].isdigit()
        ):
            return raw_parts[3]

        return ""

    def resolve_playwright_proxy_server(self) -> str:
        proxy_url = self.resolve_http_proxy()
        if not proxy_url:
            return ""

        parsed = urllib.parse.urlsplit(proxy_url)
        host = parsed.hostname or ""
        port = parsed.port
        if not host or not port:
            return ""

        scheme = parsed.scheme.lower()
        if scheme == "socks5h":
            scheme = "socks5"

        if scheme not in ("http", "https", "socks5"):
            scheme = "http"

        return f"{scheme}://{host}:{port}"

    def playwright_proxy_is_supported(self) -> bool:
        """Chromium does not support SOCKS5 proxy authentication via Playwright proxy settings."""
        server = self.resolve_playwright_proxy_server()
        if not server:
            return True

        parsed = urllib.parse.urlsplit(server)
        if parsed.scheme != "socks5":
            return True

        username = self.resolve_proxy_username()
        password = self.resolve_proxy_password()
        return not (username or password)

    def resolve_http_proxy(self) -> str:
        proxy_server = os.getenv("LUMINOR_PROXY_SERVER", "").strip().strip("'\"")

        if not proxy_server:
            return ""

        username = self.resolve_proxy_username()
        password = self.resolve_proxy_password()

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

    def fetch_table_html_via_curl(self, use_proxy: bool = True) -> str:
        proxy_url = self.resolve_http_proxy() if use_proxy else ""
        last_error = None

        for base_url in LUMINOR_TABLE_URLS:
            for url in (
                self.build_table_url(base_url),
                f"{base_url}?fund_type=pension&currency=eur&period=3year&fund=15",
            ):
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
                    url,
                ]

                if proxy_url:
                    command[1:1] = ["--proxy", proxy_url]

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

                if self.is_cloudflare_challenge(result.stdout):
                    last_error = RuntimeError(
                        "Cloudflare challenge page returned by curl"
                        + (" (via proxy)" if proxy_url else " (direct)")
                    )
                    continue

                if "data-label=\"Fondas\"" in result.stdout or "data-label='Fondas'" in result.stdout:
                    return result.stdout

                last_error = RuntimeError("Table rows not found in pensiju-fondai response")

        if last_error:
            raise last_error
        return ""

    def fetch_table_html_via_browser_navigation(self, use_proxy: bool = True) -> str:
        if use_proxy and not self.playwright_proxy_is_supported():
            raise RuntimeError(
                "Skipping browser proxy attempt: Chromium/Playwright does not support SOCKS5 proxy authentication"
            )

        self.ensure_browser_mode(use_proxy=use_proxy)

        if not self.page:
            raise RuntimeError("Browser page is not initialized")

        last_error = None

        for base_url in LUMINOR_TABLE_URLS:
            try:
                try:
                    homepage = base_url.split("/lt/pensiju-fondai", 1)[0]
                    self.page.goto(homepage + "/lt", wait_until="domcontentloaded", timeout=30000)
                    self._accept_cookies_if_visible()
                    self.page.wait_for_timeout(800)
                except Exception:
                    pass

                response = self.page.goto(base_url, wait_until="domcontentloaded", timeout=90000)
                self._accept_cookies_if_visible()

                if response and response.status >= 400:
                    raise RuntimeError(f"HTTP {response.status}")

                self.page.wait_for_timeout(1500)
                html_content = self.page.content()

                if self.is_cloudflare_challenge(html_content):
                    raise RuntimeError("Cloudflare challenge page returned by browser")

                if "data-label=\"Fondas\"" in html_content or "data-label='Fondas'" in html_content:
                    return html_content

                raise RuntimeError("Table rows not found in browser response")
            except Exception as exc:
                last_error = exc

        raise last_error

    def _strip_html(self, value: str) -> str:
        clean = re.sub(r"<[^>]+>", " ", value)
        clean = html.unescape(clean)
        return re.sub(r"\s+", " ", clean).strip()

    def _extract_fund_name(self, fund_cell_html: str) -> str:
        anchor_match = re.search(
            r"<a\b[^>]*href=[\"']?/lt/pensiju-fondu-forma\?fund=\d+[\"']?[^>]*>(.*?)</a>",
            fund_cell_html,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if anchor_match:
            return self._strip_html(anchor_match.group(1))
        return self._strip_html(fund_cell_html)

    def _extract_visible_td_value(self, cell_html: str) -> str:
        # Prefer non-mobile-only div content because that is the actual displayed value.
        div_matches = re.findall(
            r"<div\b(?![^>]*mobile-only)[^>]*>(.*?)</div>",
            cell_html,
            flags=re.IGNORECASE | re.DOTALL,
        )
        for div_html in reversed(div_matches):
            value = self._strip_html(div_html)
            if value:
                return value

        return self._strip_html(cell_html)

    def latest_existing_output_file(self) -> str:
        candidates = []
        for path in Path(".").glob("luminor_pensions_data_*.xlsx"):
            match = DATE_RE.search(path.name)
            if not match:
                continue
            candidates.append((match.group(1), path))

        if not candidates:
            return ""

        candidates.sort(key=lambda item: item[0])
        return str(candidates[-1][1])

    def parse_rows_from_table_html(self, table_html: str) -> list:
        rows = []
        row_blocks = re.findall(r"<tr\b[^>]*>(.*?)</tr>", table_html, flags=re.IGNORECASE | re.DOTALL)

        for row_html in row_blocks:
            label_cells = {
                label.strip(): content
                for label, content in re.findall(
                    r'<td\b[^>]*data-label=["\']([^"\']+)["\'][^>]*>(.*?)</td>',
                    row_html,
                    flags=re.IGNORECASE | re.DOTALL,
                )
            }

            fund_cell = label_cells.get("Fondas", "")
            unit_cell = label_cells.get("Apskaitos vieneto vertė", "")
            assets_cell = label_cells.get("Grynųjų aktyvų vertė", "")

            fund_name = self._extract_fund_name(fund_cell)
            if not fund_name or fund_name in EXCLUDED_FUNDS:
                continue

            unit_value_text = self._extract_visible_td_value(unit_cell)
            assets_text = self._extract_visible_td_value(assets_cell)

            if not unit_value_text and not assets_text:
                continue

            rows.append(
                {
                    "Fund name": fund_name,
                    "Data": None,
                    "Vieneto vertė": unit_value_text,
                    "Grynieji aktyvai": assets_text,
                }
            )

        return rows

    def fetch_html_via_browser_navigation(self, fund_id: str, use_proxy: bool = True) -> str:
        if use_proxy and not self.playwright_proxy_is_supported():
            raise RuntimeError(
                "Skipping browser proxy attempt: Chromium/Playwright does not support SOCKS5 proxy authentication"
            )

        self.ensure_browser_mode(use_proxy=use_proxy)

        if not self.page:
            raise RuntimeError("Browser page is not initialized")

        last_error = None

        for base_url in LUMINOR_BASE_URLS:
            url = self.build_url(base_url, fund_id)

            try:
                try:
                    homepage = base_url.split("/lt/rinkis-fonda", 1)[0]
                    self.page.goto(homepage + "/lt", wait_until="domcontentloaded", timeout=30000)
                    self._accept_cookies_if_visible()
                    self.page.wait_for_timeout(800)
                except Exception:
                    pass

                response = self.page.goto(url, wait_until="domcontentloaded", timeout=90000)
                self._accept_cookies_if_visible()

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
            table_blocked = False
            used_browser_fallback = False
            payload_missing_ids = []
            proxy_bypass_ids = []

            table_attempts = [
                ("curl via proxy", lambda: self.fetch_table_html_via_curl(use_proxy=True), True),
                ("curl direct", lambda: self.fetch_table_html_via_curl(use_proxy=False), False),
                ("browser via proxy", lambda: self.fetch_table_html_via_browser_navigation(use_proxy=True), True),
                ("browser direct", lambda: self.fetch_table_html_via_browser_navigation(use_proxy=False), False),
            ]

            last_table_error = None
            for attempt_name, fetch_fn, used_proxy in table_attempts:
                try:
                    table_html = fetch_fn()
                    parsed_rows = self.parse_rows_from_table_html(table_html)
                    if not parsed_rows:
                        print(f"Table attempt {attempt_name} returned no parsable rows")
                        continue

                    rows = parsed_rows
                    if not used_proxy:
                        proxy_bypass_ids = LUMINOR_II_FUND_IDS.copy()
                    print(f"Parsed {len(rows)} funds from pensiju-fondai table ({attempt_name})")
                    break
                except Exception as exc:
                    last_table_error = exc
                    print(f"Table attempt {attempt_name} failed: {exc}")
                    if "Cloudflare" in str(exc) or "HTTP 403" in str(exc):
                        table_blocked = True

            if not rows and last_table_error:
                print(f"Table scrape path failed, falling back to payload path: {last_table_error}")

            for fund_id in ([] if rows else LUMINOR_II_FUND_IDS):
                try:
                    payload = {}

                    try:
                        payload = self.fetch_payload_via_curl(fund_id, use_proxy=True)
                    except Exception as proxy_payload_exc:
                        print(f"Proxy payload attempt failed fund {fund_id}: {proxy_payload_exc}")

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
                allow_stale = os.getenv("LUMINOR_ALLOW_STALE_ON_BLOCK", "false").strip().lower() in (
                    "1",
                    "true",
                    "yes",
                )
                fully_blocked = table_blocked or len(set(cloudflare_fund_ids)) >= len(LUMINOR_II_FUND_IDS)
                if allow_stale and fully_blocked:
                    stale_file = self.latest_existing_output_file()
                    if stale_file:
                        print(
                            "No fresh Luminor data (Cloudflare/proxy block). "
                            f"Using last successful file: {stale_file}"
                        )
                        return stale_file
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