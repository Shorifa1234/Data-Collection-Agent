import asyncio
import io
import json
import os
import re
import sys
import time
from pathlib import Path

import pdfplumber
import requests as _requests

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

_PDF_SESSION = _requests.Session()
_PDF_SESSION.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Niermann Weeks")
HEADLESS = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://niermannweeks.com"

# PDF links that are not tearsheets
_SKIP_PDF_KEYWORDS = ["credit-card", "com-id", "authorization"]


def _is_tearsheet(href: str) -> bool:
    href_lower = href.lower()
    if not href_lower.endswith(".pdf"):
        return False
    return not any(kw in href_lower for kw in _SKIP_PDF_KEYWORDS)


# Known section labels in tearsheet PDFs
_TS_LABELS = {"ITEM #", "DIMENSIONS", "FINISH", "NOTES", "LIGHTS", "COLLABORATION"}
# Footer text to stop parsing at
_TS_FOOTER_RE = re.compile(r"P 4\d{2}|NIERMANNWEEKS\.COM|ALL RIGHTS RESERVED", re.IGNORECASE)


def _parse_tearsheet_pdf(pdf_url: str) -> dict:
    """Download and parse a Niermann Weeks tearsheet PDF.

    Returns a dict with any of: Finish, Notes, Collaboration, Lamp Quantity.
    Silently returns {} on any error so one bad PDF never kills the scrape.
    """
    try:
        resp = _PDF_SESSION.get(pdf_url, timeout=20)
        resp.raise_for_status()
        with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception as exc:
        print(f"  [PDF WARN] {pdf_url}: {exc}")
        return {}

    # Split into lines; strip whitespace
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    result: dict = {}
    i = 0
    while i < len(lines):
        line = lines[i]

        # Stop at footer
        if _TS_FOOTER_RE.search(line):
            break

        # Strip trailing colon to get label
        label = line.rstrip(":").upper()

        if label in _TS_LABELS and i + 1 < len(lines):
            # Collect value lines until next label, footer, or body description text.
            # Body text is identified as: line > 50 chars AND contains lowercase letters.
            val_lines: list[str] = []
            j = i + 1
            while j < len(lines):
                next_line = lines[j]
                if _TS_FOOTER_RE.search(next_line):
                    break
                next_label = next_line.rstrip(":").upper()
                if next_label in _TS_LABELS:
                    break
                # Stop if this looks like body description text (NW values are always ALL CAPS)
                if any(c.islower() for c in next_line):
                    break
                val_lines.append(next_line)
                j += 1

            value = " ".join(val_lines).strip()
            if value:
                if label == "FINISH":
                    result["Finish"] = value
                elif label == "NOTES":
                    result["Notes"] = value
                elif label == "COLLABORATION":
                    result["Designer"] = value
                elif label == "LIGHTS":
                    try:
                        result["Lamp Quantity"] = int(value)
                    except ValueError:
                        result["Lamp Quantity"] = value
            i = j
        else:
            i += 1

    return result


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from a category listing, following pagination."""
    links: list[str] = []
    seen: set[str] = set()
    current_url = listing_url

    while current_url:
        await page.goto(current_url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)

        hrefs = await page.eval_on_selector_all(
            "a[href]",
            "els => els.map(e => e.href)",
        )
        for href in hrefs:
            # Only product detail pages (not the generic /products/ index)
            if "/products/" in href and href.rstrip("/") != f"{BASE_URL}/products":
                base = href.split("?")[0].rstrip("/")
                if base not in seen:
                    seen.add(base)
                    links.append(base)

        # Follow pagination: look for <link rel="next"> via JS
        next_page = await page.evaluate(
            """() => {
                const el = document.querySelector('link[rel="next"]');
                return el ? el.href : null;
            }"""
        )
        current_url = next_page if next_page else None

    return links


async def scrape_product(page, url: str) -> list[dict]:
    """Return a list of dicts (one per product — NW has no variant system)."""
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)

    row: dict = {"Source URL": url, "Manufacturer": VENDOR_NAME}

    # Product name
    h1 = await page.query_selector("h1")
    if h1:
        row["Product Name"] = clean_text(await h1.inner_text())

    # First product image from wp-content/uploads
    imgs = await page.query_selector_all("img[src*='wp-content/uploads']")
    if imgs:
        src = await imgs[0].get_attribute("src")
        if src:
            row["Image URL"] = src

    # Tearsheet link — first PDF that is a tearsheet; then parse it for extra fields
    all_anchors = await page.query_selector_all("a[href]")
    for a in all_anchors:
        href = await a.get_attribute("href") or ""
        if _is_tearsheet(href):
            row["Tearsheet Link"] = href
            ts_data = _parse_tearsheet_pdf(href)
            row.update(ts_data)  # adds Finish, Notes, Designer, Lamp Quantity
            break

    # Parse all <p> elements for structured data
    paragraphs = await page.query_selector_all("p")
    desc_parts: list[str] = []
    item_found = False
    dims_pending = False  # set when Item# paragraph had no Dimensions

    for p_el in paragraphs:
        raw = clean_text(await p_el.inner_text())
        if not raw or "COPYRIGHT" in raw:
            continue

        # ── Item # line (may contain Dimensions too) ────────────────────────
        item_match = re.search(
            r"Item\s*#\s*[:\-]?\s*([A-Z0-9\-]+)",
            raw,
            re.IGNORECASE,
        )
        if item_match:
            item_found = True
            row["SKU"] = item_match.group(1).strip()

            # Dimensions on same line?
            dim_match = re.search(
                r"Dimensions?\s*[:\-]?\s*(.+)",
                raw,
                re.IGNORECASE,
            )
            if dim_match:
                row["Dimensions"] = dim_match.group(1).strip()
                dims_pending = False
            else:
                dims_pending = True
            continue

        # ── Standalone Dimensions line (when split from Item #) ─────────────
        if dims_pending:
            dim_match = re.search(
                r"Dimensions?\s*[:\-]?\s*(.+)",
                raw,
                re.IGNORECASE,
            )
            if dim_match:
                row["Dimensions"] = dim_match.group(1).strip()
                dims_pending = False
                continue
            # Could be "77.5 W X 2 D x 62 H" directly as next paragraph
            if re.search(r"\d+(\.\d+)?\s*(W|H|D|DIA)", raw, re.IGNORECASE):
                row["Dimensions"] = raw
                dims_pending = False
                continue

        # ── Retail Price ────────────────────────────────────────────────────
        price_match = re.search(
            r"(?:Retail\s*)?Price\s*[:\-]?\s*\$?([\d,\.]+)",
            raw,
            re.IGNORECASE,
        )
        if price_match:
            row["Price"] = clean_price(price_match.group(1))
            continue

        # ── Lamp / lights count ─────────────────────────────────────────────
        lights_match = re.match(r"^(\d+)\s+lights?$", raw, re.IGNORECASE)
        if lights_match:
            row["Lamp Quantity"] = int(lights_match.group(1))
            continue

        # ── Collection name (ends with "Collection" or starts with known prefix)
        if re.search(r"\bCollection\b", raw, re.IGNORECASE) and len(raw.split()) <= 6:
            row["Collection"] = raw
            continue

        # ── Everything else → Description ────────────────────────────────────
        desc_parts.append(raw)

    if desc_parts:
        row["Description"] = " ".join(desc_parts)

    # Derive sub-dimension fields from Dimensions string
    if "Dimensions" in row:
        dim_str = row["Dimensions"]
        parsed = parse_dimensions(dim_str)
        row["Dimensions"] = parsed.get("Dimensions", dim_str)
        for field in ("Width", "Height", "Depth", "Diameter", "Length"):
            if parsed.get(field) and field not in row:
                row[field] = parsed[field]

    # Product Family Id
    if row.get("Product Name") and not row.get("Product Family Id"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    return [row]


async def main() -> None:
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        print(f"[TEST: max {TEST_MAX_PRODUCTS} products per category, all {len(categories)} categories]")

    start_total = time.time()

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        cat_count = 0
        for cat in categories:
            if not cat["links"]:
                continue
            if TEST_MODE and cat_count >= TEST_MAX_CATEGORIES:
                break
            cat_count += 1

            cat_name = cat["name"]
            writer.add_sheet(
                cat_name,
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )
            print(f"\n[{cat_name}] Collecting product links...")

            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"[{cat_name}] {len(all_product_urls)} products to scrape")

            global_idx = 1
            for url in all_product_urls:
                try:
                    variant_rows = await scrape_product(page, url)
                    for variant in variant_rows:
                        if not variant.get("SKU"):
                            variant["SKU"] = generate_sku(
                                info["vendor_name"], cat_name, global_idx
                            )
                        if not variant.get("Product Family Id") and variant.get("Product Name"):
                            variant["Product Family Id"] = extract_family_id(
                                variant["Product Name"]
                            )
                        writer.write_row(variant, category_name=cat_name)
                        global_idx += 1
                except Exception as exc:
                    print(f"  [ERROR] {url}: {exc}")
                await async_polite_delay()

    writer.save()
    elapsed = time.time() - start_total
    print(f"\nDone. Total time: {elapsed:.1f}s  ->  {OUTPUT_PATH}")

    # Append to run log
    log_path = OUTPUT_PATH.parent / "run_log.txt"
    with open(log_path, "a", encoding="utf-8") as f:
        mode = "TEST" if TEST_MODE else "FULL"
        f.write(f"[{mode}] {time.strftime('%Y-%m-%d %H:%M:%S')}  elapsed={elapsed:.1f}s\n")


if __name__ == "__main__":
    asyncio.run(main())
