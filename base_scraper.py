#!/usr/bin/env python3
"""
Base class for pension fund scrapers.
Provides shared functionality for browser setup, error handling, and file output.
"""
import os
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
from pathlib import Path
import sys
import re
from abc import ABC, abstractmethod


class BaseScraper(ABC):
    """Abstract base class for pension fund scrapers."""
    
    def __init__(self, source_name: str):
        """
        Initialize scraper.
        
        Args:
            source_name: Name of the data source (e.g., 'swedbank')
        """
        self.source_name = source_name
        self.results = []
        self._playwright = None
        self.browser = None
        self.context = None
        self.page = None
    
    @abstractmethod
    def get_url(self) -> str:
        """Return the URL to scrape."""
        pass
    
    @abstractmethod
    def scrape_data(self, page) -> list:
        """
        Scrape data from the page. Subclasses must implement this.
        
        Args:
            page: Playwright page object
            
        Returns:
            List of dictionaries with scraped data
        """
        pass
    
    def _is_headless(self) -> bool:
        return os.getenv("PLAYWRIGHT_HEADLESS", "true").strip().lower() not in ("false", "0", "no", "off")

    def _get_proxy_settings(self):
        proxy_server = (
            os.getenv("PLAYWRIGHT_PROXY_SERVER")
            or os.getenv("HTTPS_PROXY")
            or os.getenv("HTTP_PROXY")
        )
        if not proxy_server:
            return None

        proxy_settings = {"server": proxy_server}
        proxy_username = os.getenv("PLAYWRIGHT_PROXY_USERNAME") or os.getenv("PROXY_USERNAME")
        proxy_password = os.getenv("PLAYWRIGHT_PROXY_PASSWORD") or os.getenv("PROXY_PASSWORD")
        if proxy_username or proxy_password:
            proxy_settings["username"] = proxy_username or ""
            proxy_settings["password"] = proxy_password or ""

        return proxy_settings

    def setup_browser(self):
        """Initialize browser and page."""
        headless_mode = self._is_headless()
        proxy_settings = self._get_proxy_settings()
        if proxy_settings:
            print(
                f"Starting browser (headless={headless_mode}) with proxy={proxy_settings['server']}"
            )
        else:
            print(f"Starting browser (headless={headless_mode})...")

        self._playwright = sync_playwright().start()
        # Provide robust launch args for CI environments (xvfb, no-sandbox, etc.)
        launch_options = {
            "headless": headless_mode,
            "args": [
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
                "--disable-gpu",
            ],
        }
        if proxy_settings:
            launch_options["proxy"] = proxy_settings

        self.browser = self._playwright.chromium.launch(**launch_options)

        # Default context options to make pages appear like real users
        context_args = {
            "user_agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "locale": "lt-LT",
            "timezone_id": "Europe/Vilnius",
            "viewport": {"width": 1280, "height": 800},
        }

        self.context = self.browser.new_context(**context_args)
        self.page = self.context.new_page()
        # Mask webdriver flag to reduce bot detection surface
        try:
            self.page.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
        except Exception:
            pass
        # Apply playwright-stealth if available to reduce bot fingerprint
        try:
            from playwright_stealth import stealth_sync
            stealth_sync(self.page)
        except Exception:
            pass

        # Longer default timeout for network-heavy pages in CI
        self.page.set_default_timeout(int(os.getenv("PLAYWRIGHT_TIMEOUT_MS", "60000")))
        return self.page
    
    def cleanup_browser(self):
        """Close browser if open."""
        if self.page:
            try:
                self.page.close()
            except Exception:
                pass
        if self.context:
            try:
                self.context.close()
            except Exception:
                pass
        if self.browser:
            try:
                self.browser.close()
            except Exception:
                pass
        if self._playwright:
            try:
                self._playwright.stop()
            except Exception:
                pass

    def _extract_data_date(self, df: pd.DataFrame) -> str:
        """Extract a normalized YYYY-MM-DD date from Data/Date columns when possible."""
        candidate_columns = [col for col in ("Data", "Date") if col in df.columns]
        for column in candidate_columns:
            values = []
            for raw in df[column].dropna().astype(str):
                normalized = re.sub(r"[\s/.]", "-", raw.strip())
                match = re.search(r"\b\d{4}-\d{2}-\d{2}\b", normalized)
                if match:
                    values.append(match.group(0))
            if values:
                # Pick latest valid date string for deterministic output naming.
                return max(values)
        return ""
    
    def save_to_excel(self, df: pd.DataFrame, filename: str) -> str:
        """
        Save DataFrame to Excel file.
        
        Args:
            df: DataFrame to save
            filename: Output filename
            
        Returns:
            Full path to created file
        """
        if df.empty:
            print(f"No data to save for {self.source_name}.")
            return None
        
        filepath = Path(filename)
        df.to_excel(filepath, index=False)
        return str(filepath)
    
    def run(self):
        """
        Main execution method. Handles browser lifecycle and error handling.
        
        Returns:
            Path to created Excel file, or None if failed
        """
        try:
            # Setup
            self.setup_browser()
            
            # Scrape
            url = self.get_url()
            print(f"Opening: {url}")
            self.page.goto(url, wait_until="domcontentloaded", timeout=90000)
            
            self.results = self.scrape_data(self.page)
            
            if not self.results:
                print(f"No data scraped from {self.source_name}. Page structure may have changed.")
                return None
            
            # Save
            df = pd.DataFrame(self.results)
            
            if df.empty:
                print(f"No data parsed for {self.source_name}.")
                return None
            
            data_date = self._extract_data_date(df)
            if not data_date:
                print(
                    f"No valid source date found in scraped data for {self.source_name}. "
                    "Skipping file creation to avoid wrong fallback date."
                )
                return None
            
            filename = f"{self.source_name}_data_{data_date}.xlsx"
            
            filepath = self.save_to_excel(df, filename)
            
            if filepath:
                print(f"✅ Excel file created: {filename}")
            
            return filepath
        
        except Exception as e:
            print(f"Error scraping {self.source_name}: {e}")
            return None
        
        finally:
            self.cleanup_browser()
