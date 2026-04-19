#!/usr/bin/env python3
"""
Merge pension fund performance data with fund size data into a single Excel file.
"""
import pandas as pd
from datetime import datetime
import sys
from pathlib import Path


def get_latest_file(pattern):
    files = list(Path(".").glob(pattern))
    if not files:
        return None
    return max(files, key=lambda path: path.stat().st_mtime)


# Read the two Excel files
pensions_file = get_latest_file("swedbank_pensions_*.xlsx")
sizes_file = get_latest_file("swedbank_fondo_dydis_*.xlsx")

if not pensions_file:
    print("Error: no pension file found. Run scraper.py first.")
    sys.exit(1)

if not sizes_file:
    print("Warning: no fund size file found. Run scrape_fund_sizes.py first.")

print("Reading pension fund data...")
df_pensions = pd.read_excel(pensions_file)
print(f"  Loaded {len(df_pensions)} funds")

if sizes_file:
    print("Reading fund size data...")
    df_sizes = pd.read_excel(sizes_file)
    print(f"  Loaded {len(df_sizes)} funds")
    
    # Merge on Fund name (left join to keep all pension funds)
    print("Merging data...")
    df_combined = df_pensions.merge(df_sizes, on="Fund name", how="left")
    print(f"  Combined: {len(df_combined)} rows, {len(df_combined.columns)} columns")
else:
    df_combined = df_pensions

# Save to new Excel file
today = datetime.today().strftime('%Y-%m-%d')
output_file = f"swedbank_pension_data_combined_{today}.xlsx"

print(f"Writing to {output_file}...")
df_combined.to_excel(output_file, index=False)

print(f"\n✅ Merged file created: {output_file}")
print(f"Rows: {len(df_combined)}")
print(f"Columns: {list(df_combined.columns)}")
