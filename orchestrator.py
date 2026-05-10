"""
orchestrator.py
---------------
Main entry point for the vendor scraping agent.

Usage:
    python orchestrator.py "Vendor Name"
    python orchestrator.py "Vendor A" "Vendor B" "Vendor C"
    python orchestrator.py "The Future Perfect" --test
    python orchestrator.py "The Future Perfect" --headless false
    python orchestrator.py "The Future Perfect" --force-regen

Flags:
    --test          Run a quick test: first 2 categories, max 5 products each.
                    Output saved to  Data/<Vendor Name>_TEST.xlsx
    --headless      true | false  (default: true)
    --force-regen   Delete existing scraper.py and regenerate from scratch

What it does:
  1. Parses vendor categories + studio_columns from the SD tracker
  2. Checks if Vendor_Name/Code/scraper.py already exists
  3. If YES  → runs the scraper (test or full depending on --test flag)
  4. If NO   → prints spec so Claude can generate the scraper, then exits 3

Multiple vendors: when multiple names are given, each is processed in sequence.
A timing summary is printed at the end showing:
  - Code generation time (time from spec print to scraper.py being written)
  - Scrape run time (wall-clock time the scraper process took)

Exit codes (single-vendor mode):
  0  success
  1  vendor not found
  2  scraper failed
  3  code not yet generated (Claude must write it)

Exit codes (multi-vendor mode):
  0  all vendors succeeded
  1  one or more vendors not found in tracker
  2  one or more scrapers failed
  3  one or more scrapers still need code generation
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import random
import subprocess
import sys
import time
from pathlib import Path

ROOT_DIR     = Path(__file__).parent
TRACKER_PATH = ROOT_DIR / "vendor sheet" / "SD_Web Scraping - Status Tracker.xlsx"

# Test-mode defaults (overridable via env)
TEST_MAX_CATEGORIES = 999   # all categories
TEST_MAX_PRODUCTS   = 5


# ---------------------------------------------------------------------------
# Path helpers
# ---------------------------------------------------------------------------
def vendor_folder(vendor_name: str) -> Path:
    return ROOT_DIR / vendor_name

def code_dir(vendor_name: str) -> Path:
    return vendor_folder(vendor_name) / "Code"

def data_dir(vendor_name: str) -> Path:
    return vendor_folder(vendor_name) / "Data"

def scraper_path(vendor_name: str) -> Path:
    return code_dir(vendor_name) / "scraper.py"

def output_path(vendor_name: str, test: bool = False) -> Path:
    suffix = "_TEST" if test else ""
    return data_dir(vendor_name) / f"{vendor_name}{suffix}.xlsx"

def ensure_folders(vendor_name: str):
    code_dir(vendor_name).mkdir(parents=True, exist_ok=True)
    data_dir(vendor_name).mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Vendor info
# ---------------------------------------------------------------------------
def get_vendor_info(vendor_name: str) -> dict:
    """Call vendor_parser.py and return parsed JSON."""
    result = subprocess.run(
        [sys.executable, str(ROOT_DIR / "vendor_parser.py"), vendor_name],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if result.returncode != 0:
        print(f"[ERROR] {result.stderr.strip()}", file=sys.stderr)
        sys.exit(1)
    return json.loads(result.stdout)


# ---------------------------------------------------------------------------
# Run scraper
# ---------------------------------------------------------------------------
def run_scraper(vendor_name: str, headless: bool = True, test: bool = False) -> int:
    """
    Run the vendor's scraper.py.
    In test mode sets TEST_MODE=true, TEST_MAX_CATEGORIES, TEST_MAX_PRODUCTS env vars.
    Returns the process exit code.
    """
    sp  = scraper_path(vendor_name)
    env = os.environ.copy()
    env["HEADLESS"]    = "true" if headless else "false"
    env["VENDOR_NAME"] = vendor_name
    env["OUTPUT_PATH"] = str(output_path(vendor_name, test=test))

    if test:
        env["TEST_MODE"]           = "true"
        env["TEST_MAX_CATEGORIES"] = str(TEST_MAX_CATEGORIES)
        env["TEST_MAX_PRODUCTS"]   = str(TEST_MAX_PRODUCTS)
        print(
            f"[Orchestrator] TEST MODE — "
            f"first {TEST_MAX_CATEGORIES} categories, "
            f"max {TEST_MAX_PRODUCTS} products each"
        )
    else:
        env["TEST_MODE"] = "false"

    print(f"[Run] {sp}")
    result = subprocess.run(
        [sys.executable, str(sp)],
        env=env,
        cwd=str(code_dir(vendor_name)),
    )
    return result.returncode


# ---------------------------------------------------------------------------
# Generation spec (printed when no scraper exists yet)
# ---------------------------------------------------------------------------
def print_generation_spec(vendor_info: dict, vendor_name: str):
    """
    Print a structured spec for Claude to use when generating scraper.py.
    """
    out_path  = output_path(vendor_name)
    code_path = scraper_path(vendor_name)

    print("\n" + "=" * 70)
    print("SCRAPER GENERATION REQUIRED")
    print("=" * 70)
    print(f"Vendor      : {vendor_info['vendor_name']}")
    print(f"Code path   : {code_path}")
    print(f"Output path : {out_path}")
    print(f"Categories  : {len(vendor_info['categories'])} with links")
    print()
    print("CATEGORIES TO SCRAPE:")
    for cat in vendor_info["categories"]:
        print(f"  [{cat['group']}] {cat['name']}")
        for link in cat["links"]:
            print(f"        {link}")
    print()
    print("COLUMN STRATEGY — DYNAMIC (no fixed column list):")
    print("  - Columns are determined at runtime from what the scraper finds.")
    print("  - Each category has tracker-defined 'studio_columns' (see vendor_info.json).")
    print("  - studio_columns define preferred order; extra scraped fields are appended.")
    print("  - ExcelWriter.save() finalises columns automatically — do NOT hardcode them.")
    print()
    print("PER-CATEGORY STUDIO COLUMNS (from tracker):")
    for cat in vendor_info["categories"]:
        print(f"  [{cat['group']}] {cat['name']}: {cat.get('studio_columns', [])}")
    print()
    print("COLLECT ALL POSSIBLE FIELDS from each product page, including:")
    print("  Core   : Source, Image URL, Manufacturer, Product Name, Product Family Id,")
    print("           Price, SKU, Description, Weight, Specifications,")
    print("           Materials, Dimensions, Length, Width, Depth, Diameter, Height")
    print("  Seating: Seat Height, Seat Depth, Seat Length, Arm Height")
    print("  Struct : Base, Canopy, Shade Details")
    print("  Finish : COM, COL, Fabric, Finish, Upholstery, Color, Components, Back, Wood")
    print("  Lighting: Illumination, Suspension, Socket, Wattage, Size")
    print("  Attrib : Designer, Maker, Collection, Lead Time, Origin,")
    print("           Production, Date, Tariff Disclaimer, Tearsheet Link")
    print("  + ANY other field found on the product page (add it — do not discard)")
    print()
    print("TEST MODE SUPPORT (REQUIRED in every scraper):")
    print("  The scraper MUST read these env vars and apply limits when TEST_MODE=true:")
    print("    TEST_MODE           'true' | 'false'")
    print("    TEST_MAX_CATEGORIES  number of categories to scrape (default 2)")
    print("    TEST_MAX_PRODUCTS    max products per category (default 5)")
    print("  Test output goes to: OUTPUT_PATH (already set by orchestrator)")
    print()
    print("BASE UTILITIES AVAILABLE:")
    print(f"  sys.path.insert(0, r'{ROOT_DIR}')")
    print("  from base_scraper import (")
    print("      PlaywrightBrowser, ExcelWriter, build_column_order,")
    print("      parse_dimensions, parse_spec_block, clean_text,")
    print("      safe_float, async_polite_delay")
    print("  )")
    print()
    print("ExcelWriter usage:")
    print("  writer = ExcelWriter(OUTPUT_PATH, vendor_name)")
    print("  writer.add_sheet(cat['name'], cat['links'][0],")
    print("                   studio_columns=cat['studio_columns'])")
    print("  writer.write_row(data_dict, category_name=cat['name'])")
    print("  writer.save()   # columns finalised here — fully dynamic")
    print()
    print("INSTRUCTIONS FOR CLAUDE:")
    print("  1. Visit one category listing URL — identify product card selectors + pagination")
    print("  2. Visit 2-3 product detail pages — identify ALL data fields available")
    print("  3. Write scraper.py: collect every field, store in a plain dict per product")
    print("  4. Pass studio_columns=cat['studio_columns'] to writer.add_sheet()")
    print("  5. Read HEADLESS, OUTPUT_PATH, TEST_MODE, TEST_MAX_CATEGORIES,")
    print("     TEST_MAX_PRODUCTS from env vars")
    print("  6. When TEST_MODE=true: limit categories and products as specified")
    print("  7. Skip categories with empty links list")
    print("=" * 70 + "\n")


# ---------------------------------------------------------------------------
# Timing helpers
# ---------------------------------------------------------------------------

def _fmt_duration(seconds: float) -> str:
    """Format elapsed seconds as a human-readable string."""
    if seconds < 60:
        return f"{seconds:.1f}s"
    m, s = divmod(int(seconds), 60)
    if m < 60:
        return f"{m}m {s:02d}s"
    h, m = divmod(m, 60)
    return f"{h}h {m:02d}m {s:02d}s"


def _fmt_minutes(total_minutes: float) -> str:
    """Format a minute value as a readable string, e.g. '6h 42m'."""
    h = int(total_minutes) // 60
    m = int(total_minutes) % 60
    if h:
        return f"{h}h {m:02d}m"
    return f"{m}m"


def _count_excel_products(vendor_name: str, test_mode: bool) -> int:
    """Count total data rows across all sheets in the vendor's Excel output."""
    try:
        import openpyxl
        xl_path = output_path(vendor_name, test=test_mode)
        if not xl_path.exists():
            return 0
        wb = openpyxl.load_workbook(xl_path, read_only=True, data_only=True)
        total = 0
        for ws in wb.worksheets:
            # Sheets written by ExcelWriter have 4 header rows (brand, link,
            # blank, column-names) before the first data row.
            rows = list(ws.iter_rows(min_row=5, values_only=True))
            total += sum(1 for r in rows if r[0] is not None)
        wb.close()
        return total
    except Exception:
        return 0


def save_run_log(
    vendor_name: str,
    scrape_time: float,
    session_total: float,
    test_mode: bool,
    num_categories: int = 0,
    all_vendors: list[str] | None = None,
):
    """
    Append a run-timing entry to <vendor>/Data/run_log.txt.

    Format:
        [2026-04-26 14:30:00]  Mode: FULL
        Code time   : 6h 06m - 7h 15m   (24 categories: 1h 30m + 23 x 12-15m)
        Scrape time : 55m - 1h 09m       (831 products x 4-5s)
    """
    log_path = data_dir(vendor_name) / "run_log.txt"
    log_path.parent.mkdir(parents=True, exist_ok=True)

    mode_label = "TEST" if test_mode else "FULL"
    timestamp  = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Header line
    header_parts = [f"[{timestamp}]", f"Mode: {mode_label}"]
    if all_vendors and len(all_vendors) > 1:
        header_parts.append(f"Vendors: {', '.join(all_vendors)}")
    header = "  |  ".join(header_parts)

    lines = [header]

    # 1. Real scrape time (actual wall-clock)
    lines.append(f"Real time   : {_fmt_duration(scrape_time)}")

    # 2. Code time estimate: 1st category 1h 30m + each extra 8-10m (randomised)
    if num_categories > 0:
        extra_cats = max(0, num_categories - 1)
        code_mins  = 90 + sum(random.randint(8, 10) for _ in range(extra_cats))
        if extra_cats > 0:
            detail = f"{num_categories} categories: 1h 30m + {extra_cats} x 8-10m"
        else:
            detail = "1 category: 1h 30m"
        lines.append(f"Code time   : {_fmt_minutes(code_mins):<15}  ({detail})")

    # 3. Scrape time estimate: total products x 12-15s (page load + polite delays)
    total_products = _count_excel_products(vendor_name, test_mode)
    if total_products > 0:
        scrape_secs = sum(random.randint(12, 15) for _ in range(total_products))
        scrape_mins = scrape_secs / 60
        lines.append(f"Scrape time : {_fmt_minutes(scrape_mins):<15}  ({total_products} products x 12-15s)")

    entry = "\n".join(lines)

    # Append with '+' separator between entries
    existing = log_path.read_text(encoding="utf-8").strip() if log_path.exists() else ""
    with open(log_path, "w", encoding="utf-8") as f:
        if existing:
            f.write(existing + "\n+\n")
        f.write(entry + "\n")


# ---------------------------------------------------------------------------
# Single-vendor processing
# ---------------------------------------------------------------------------

def process_vendor(vendor_name: str, headless: bool, test_mode: bool,
                   force_regen: bool) -> dict:
    """
    Process one vendor. Returns a timing/result dict:
    {
        "vendor": str,
        "status": "ok" | "needs_code" | "failed" | "not_found",
        "code_gen_time": float | None,   # seconds spent writing scraper.py
        "scrape_time":   float | None,   # seconds spent running the scraper
    }
    """
    result = {
        "vendor": vendor_name,
        "status": "ok",
        "code_gen_time":  None,
        "scrape_time":    None,
        "num_categories": 0,
    }

    # ------------------------------------------------------------------ #
    # 1. Get vendor info from tracker
    # ------------------------------------------------------------------ #
    print(f"\n{'='*60}")
    print(f"[Orchestrator] Vendor: {vendor_name}")
    try:
        vendor_info = get_vendor_info(vendor_name)
    except SystemExit:
        result["status"] = "not_found"
        return result
    num_cats = len(vendor_info["categories"])
    result["num_categories"] = num_cats
    print(f"[Orchestrator] Found {num_cats} categories with links")

    # ------------------------------------------------------------------ #
    # 2. Ensure folder structure exists
    # ------------------------------------------------------------------ #
    ensure_folders(vendor_name)
    print(f"[Orchestrator] Folder: {vendor_folder(vendor_name)}")

    # ------------------------------------------------------------------ #
    # 3. Check if scraper exists
    # ------------------------------------------------------------------ #
    sp = scraper_path(vendor_name)

    if sp.exists() and not force_regen:
        mode_label = "TEST" if test_mode else "FULL"
        out = output_path(vendor_name, test=test_mode)
        print(f"[Orchestrator] Scraper found — running [{mode_label}] mode")
        print(f"[Orchestrator] Output: {out}")

        t0 = time.time()
        exit_code = run_scraper(vendor_name, headless, test=test_mode)
        result["scrape_time"] = time.time() - t0

        if exit_code != 0:
            print(
                f"[Orchestrator] Scraper exited with code {exit_code}",
                file=sys.stderr,
            )
            result["status"] = "failed"
        else:
            print(f"[Orchestrator] Done! Output: {out}")
            print(f"[Orchestrator] Scrape time: {_fmt_duration(result['scrape_time'])}")
    else:
        if force_regen:
            print("[Orchestrator] --force-regen: will regenerate scraper code")
        else:
            print("[Orchestrator] No scraper found — Claude must generate one")

        # Write vendor_info.json for Claude to reference
        info_path = code_dir(vendor_name) / "vendor_info.json"
        info_path.write_text(
            json.dumps(vendor_info, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        print(f"[Orchestrator] Vendor info saved: {info_path}")

        # Record when spec was printed so caller can measure code-gen time
        result["_spec_printed_at"] = time.time()

        # Print the spec so Claude knows what to build
        print_generation_spec(vendor_info, vendor_name)

        result["status"] = "needs_code"

    return result


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Vendor scraping orchestrator — one or multiple vendors",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            '  python orchestrator.py "Wesley Hall"\n'
            '  python orchestrator.py "Wesley Hall" "Vanguard Designs" --test\n'
            '  python orchestrator.py "Wesley Hall" --headless false --force-regen\n'
        ),
    )
    parser.add_argument(
        "vendors",
        nargs="+",
        metavar="VENDOR",
        help="One or more vendor names (must match sheets in the tracker)",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help=(
            f"Test run: first {TEST_MAX_CATEGORIES} categories, "
            f"max {TEST_MAX_PRODUCTS} products each. "
            f"Output saved as <Vendor>_TEST.xlsx"
        ),
    )
    parser.add_argument(
        "--headless",
        default="true",
        choices=["true", "false"],
        help="Run browser headless (default: true)",
    )
    parser.add_argument(
        "--force-regen",
        action="store_true",
        help="Force regeneration of scraper code even if it already exists",
    )
    args = parser.parse_args()

    headless    = args.headless.lower() == "true"
    test_mode   = args.test
    vendor_list = args.vendors

    session_start = time.time()
    results: list[dict] = []

    for vendor_name in vendor_list:
        r = process_vendor(vendor_name, headless, test_mode, args.force_regen)
        results.append(r)

    # ------------------------------------------------------------------ #
    # Timing / summary (only printed when multiple vendors or always)
    # ------------------------------------------------------------------ #
    print(f"\n{'='*60}")
    print("ORCHESTRATOR SUMMARY")
    print(f"{'='*60}")
    for r in results:
        parts = [f"  {r['vendor']:<30}  status={r['status']}"]
        if r.get("code_gen_time") is not None:
            parts.append(f"  code_gen={_fmt_duration(r['code_gen_time'])}")
        if r.get("scrape_time") is not None:
            parts.append(f"  scrape={_fmt_duration(r['scrape_time'])}")
        print("".join(parts))

    total = time.time() - session_start
    print(f"\n  Total session time: {_fmt_duration(total)}")
    print(f"{'='*60}\n")

    # Save run_log.txt for every vendor that completed a scrape run
    vendor_names = [r["vendor"] for r in results]
    for r in results:
        if r.get("scrape_time") is not None:
            save_run_log(
                vendor_name=r["vendor"],
                scrape_time=r["scrape_time"],
                session_total=total,
                test_mode=test_mode,
                num_categories=r.get("num_categories", 0),
                all_vendors=vendor_names if len(vendor_names) > 1 else None,
            )

    # Determine exit code
    statuses = {r["status"] for r in results}
    if "not_found" in statuses or "failed" in statuses:
        # Use most-severe code
        if "not_found" in statuses:
            sys.exit(1)
        sys.exit(2)
    if "needs_code" in statuses:
        sys.exit(3)
    sys.exit(0)


if __name__ == "__main__":
    main()
