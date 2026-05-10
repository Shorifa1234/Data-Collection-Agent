"""
scraper.py  —  Kravet
-----------------------
Platform: kravet.com (Magento 2 + Algolia InstantSearch)
Uses Playwright — products are rendered client-side via Algolia JavaScript.

Site structure:
  - Listing: /furniture?furniture_subtype=Nightstands etc.
    Products rendered via Algolia into .product-item / .product-item-info elements.
    Each card contains two `a.result` links (both pointing to the same URL).
    Algolia query params (?queryID=...&objectID=...&indexName=...) are stripped.
    Pagination: query param ?page=N (Algolia-style, NOT Magento ?p=N).
    Next page button: .ais-Pagination-item--nextPage a
  - Product detail: Magento 2 product page
    - Name: itemprop="name"
    - SKU:  itemprop="sku" (may be name-based for custom/ICreate pieces)
    - Description: div.product-description-container
    - Image: img[src*='brandfolder.io'] (first large image, width=1200)
    - Specs: table#product-attribute-specs-table (Width/Depth/Height in inches)
    - Collection, Lead Time, Origin, Wood Type: from spec table rows
    - Tearsheet: first .pdf link in table#product-resources-table

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Kravet"
    python orchestrator.py "Kravet" --test
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup

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
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Kravet")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://www.kravet.com"
TIMEOUT_MS = 60_000
MAX_PAGES  = 200


async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Collect all product URLs from a Kravet Algolia-rendered listing page.

    Product cards render as .product-item-info with two a.result links per card
    (both pointing to the same URL). Algolia query params are stripped.

    Pagination: &page=N  (NOT &p=N which Magento uses — Kravet uses Algolia)
    Next button: .ais-Pagination-item--nextPage a
    """
    links: list[str] = []
    seen:  set[str]  = set()
    page_num = 1

    while page_num <= MAX_PAGES:
        url = listing_url if page_num == 1 else f"{listing_url}&page={page_num}"
        print(f"    [Listing p{page_num}] {url}")

        try:
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await page.wait_for_timeout(3000)   # let Algolia JS render
        except Exception as e:
            print(f"    [WARN] Navigation failed: {e}")
            break

        # Wait for product cards to appear
        try:
            await page.wait_for_selector(".product-item-info", timeout=10_000)
        except Exception:
            print(f"    [WARN] No .product-item-info elements on page {page_num}")
            break

        # a.result links — two per card, both same URL; dedup by stripping Algolia params
        hrefs: list[str] = await page.eval_on_selector_all(
            ".product-item-info a.result",
            "els => els.map(a => a.href).filter(Boolean)",
        )

        added = 0
        for href in hrefs:
            # Strip Algolia tracking params (queryID, objectID, indexName)
            clean = re.sub(r"\?.*", "", href).rstrip("/")
            if clean and clean not in seen and BASE_URL in clean:
                seen.add(clean)
                links.append(clean)
                added += 1

        print(f"    [Listing p{page_num}] +{added} products (total {len(links)})")

        if added == 0:
            break

        # Algolia pagination next button
        next_btn = await page.query_selector(".ais-Pagination-item--nextPage a")
        if not next_btn:
            break

        page_num += 1
        await async_polite_delay(2.0, 3.0)

    return links


def _clean_dim_value(raw: str) -> str:
    """Strip trailing 'in' and excessive whitespace from spec-table dimension values."""
    val = re.sub(r"\s+", " ", raw).strip()
    val = re.sub(r"\s*in\.?$", "", val, flags=re.I).strip()
    return val


async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape one Kravet product detail page (Magento 2 + Algolia).
    Returns list[dict] — one per swatch variant if present, otherwise single row.

    Kravet does not have a Product JSON-LD on detail pages. All fields are
    extracted from HTML: itemprop attributes, div.product-description-container,
    table#product-attribute-specs-table, and img[src*='brandfolder.io'].
    """
    base: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"    [WARN] Load failed: {url} — {e}")
        return [base]

    html = await page.content()
    soup = BeautifulSoup(html, "lxml")

    # ── 1. Product Name ──────────────────────────────────────────────────────
    name_el = soup.find(itemprop="name")
    if name_el:
        name = clean_text(name_el.get_text())
        if name:
            base["Product Name"] = name.upper()
            base["Product Family Id"] = extract_family_id(base["Product Name"])
    if not base.get("Product Name"):
        h1 = soup.find("h1")
        if h1:
            name = clean_text(h1.get_text())
            base["Product Name"] = name.upper()
            base["Product Family Id"] = extract_family_id(base["Product Name"])

    # ── 2. SKU ────────────────────────────────────────────────────────────────
    # itemprop="sku" may contain a proper code (OPL150.NIGHTSTAND.0) or the
    # product name text for ICreate custom pieces — accept it either way.
    sku_el = soup.find(itemprop="sku")
    if sku_el:
        base["SKU"] = clean_text(sku_el.get_text())

    # ── 3. Description ────────────────────────────────────────────────────────
    desc_el = soup.find("div", class_="product-description-container")
    if desc_el:
        desc = clean_text(desc_el.get_text())
        if len(desc) > 10:
            base["Description"] = desc

    # ── 4. Image URL ─────────────────────────────────────────────────────────
    # Kravet uses cdn.brandfolder.io for product images. Request 1200px.
    for img in soup.find_all("img"):
        src = (
            img.get("data-zoom-image")
            or img.get("data-src")
            or img.get("src", "")
        )
        if not src or src.startswith("data:"):
            continue
        if "brandfolder.io" in src or "product" in src.lower() or "catalog" in src.lower():
            # Upgrade to 1200px
            src = re.sub(r"width=\d+", "width=1200", src)
            src = re.sub(r"height=\d+", "height=1200", src)
            # Remove pad=true noise
            base["Image URL"] = src
            break

    # ── 5. Spec table (table#product-attribute-specs-table) ──────────────────
    # Kravet may have two tables with this id: main specs + "End Use" table.
    spec_parts: list[str] = []
    dim_parts: list[str] = []

    for tbl in soup.find_all("table", id="product-attribute-specs-table"):
        for row in tbl.find_all("tr"):
            cells = row.find_all(["th", "td"])
            if len(cells) < 2:
                continue
            key   = clean_text(cells[0].get_text())
            value = clean_text(cells[1].get_text())
            if not key or not value:
                continue

            spec_parts.append(f"{key}: {value}")
            key_lower = key.lower()

            # Dimensions — spec table has separate Width/Depth/Height rows
            if key_lower in ("width", "depth", "height", "length", "diameter"):
                cleaned_val = _clean_dim_value(value)
                if cleaned_val and cleaned_val not in ("0", "0.00"):
                    base.setdefault(key.capitalize(), cleaned_val)
                    dim_parts.append(f"{key[0].upper()} {cleaned_val}")

            # Named mappings
            elif key_lower == "collection":
                base.setdefault("Collection", value)
            elif "delivery" in key_lower or "lead time" in key_lower:
                base.setdefault("Lead Time", value)
            elif key_lower in ("country of manufacture", "origin", "made in"):
                base.setdefault("Origin", value)
            elif key_lower == "designer":
                base.setdefault("Designer", value)
            elif key_lower in ("wood type", "wood", "material"):
                base.setdefault("Material", value)
            elif key_lower == "end use":
                base.setdefault("Use", value)

    if spec_parts:
        base["Specifications"] = " | ".join(spec_parts)

    # Build Dimensions string from extracted parts (if not already set)
    if dim_parts and not base.get("Dimensions"):
        base["Dimensions"] = " x ".join(dim_parts)

    # ── 6. Tearsheet Link ─────────────────────────────────────────────────────
    resources_tbl = soup.find("table", id="product-resources-table")
    if resources_tbl:
        for a in resources_tbl.find_all("a", href=True):
            href = a["href"]
            text = a.get_text(strip=True)
            # Prefer tearsheet PDF (not the generic finishes sheet)
            if "tearsheet" in href.lower() or "tearsheet" in text.lower():
                base["Tearsheet Link"] = href if href.startswith("http") else BASE_URL + href
                break
        # Fallback to any PDF in the resources table
        if not base.get("Tearsheet Link"):
            for a in resources_tbl.find_all("a", href=re.compile(r"\.pdf", re.I)):
                href = a["href"]
                base["Tearsheet Link"] = href if href.startswith("http") else BASE_URL + href
                break

    # ── 7. Variant swatches ───────────────────────────────────────────────────
    swatch_attrs = await page.query_selector_all(".swatch-attribute")
    if not swatch_attrs:
        return [base]

    variant_rows: list[dict] = []
    for attr_el in swatch_attrs:
        attr_name_el = await attr_el.query_selector(".swatch-attribute-label")
        attr_name = clean_text(await attr_name_el.inner_text()) if attr_name_el else "Finish"

        swatch_options = await attr_el.query_selector_all(".swatch-option")
        for opt in swatch_options:
            label = (
                await opt.get_attribute("aria-label")
                or await opt.get_attribute("title")
                or clean_text(await opt.inner_text())
            )
            if label and label not in ("", "undefined"):
                row = dict(base)
                row[attr_name] = clean_text(label)
                variant_rows.append(row)

    return variant_rows if variant_rows else [base]


async def main():
    info       = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer     = ExcelWriter(OUTPUT_PATH, info["vendor_name"])
    categories = info["categories"]

    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor  : {info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")
    print(f"[Scraper] Headless: {HEADLESS}")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                print(f"[Skip] {cat['name']} — no links")
                continue

            cat_url = cat["links"][0]
            writer.add_sheet(cat["name"], cat_url, studio_columns=cat["studio_columns"])

            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)
                if TEST_MODE and len(all_product_urls) >= TEST_MAX_PRODUCTS:
                    break

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"\n[Category] {cat['name']}: {len(all_product_urls)} products")

            global_idx = 1
            for product_url in all_product_urls:
                print(f"  [{global_idx}/{len(all_product_urls)}] {product_url}")
                try:
                    variant_rows = await scrape_product(page, product_url)
                    for row in variant_rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [ERROR] {product_url}: {e}")
                await async_polite_delay(1.5, 3.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
