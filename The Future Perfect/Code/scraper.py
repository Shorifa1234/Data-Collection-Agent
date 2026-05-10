"""
scraper.py  —  The Future Perfect
----------------------------------
Scrapes ALL product data from thefutureperfect.com for every category
defined in the SD tracker, with fully dynamic column output.

Columns are NOT fixed — every field found on the product page is collected
and written to the Excel. Column order is driven by each category's
studio_columns from the tracker, with extra fields appended dynamically.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "The Future Perfect"           # full run
    python orchestrator.py "The Future Perfect" --test    # test run

Env vars (set by orchestrator):
    HEADLESS             true | false   (default: true)
    OUTPUT_PATH          absolute path to output .xlsx
    VENDOR_NAME          vendor name string
    TEST_MODE            true | false   (default: false)
    TEST_MAX_CATEGORIES  max categories in test mode (default: 2)
    TEST_MAX_PRODUCTS    max products per category in test mode (default: 5)
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

# ── project root so base_scraper is importable ──────────────────────────────
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from bs4 import BeautifulSoup

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
    parse_spec_block,
    safe_float,
)

# ── config ───────────────────────────────────────────────────────────────────
VENDOR_NAME = os.environ.get("VENDOR_NAME", "The Future Perfect")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)

# Test-mode limits (set by orchestrator via env)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://www.thefutureperfect.com"
TIMEOUT_MS = 45_000
MAX_PAGES  = 500


# ── listing page helpers ─────────────────────────────────────────────────────
async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from a category listing, handling pagination."""
    links: list[str] = []
    seen:  set[str]  = set()
    page_num = 1

    while page_num <= MAX_PAGES:
        url = listing_url if page_num == 1 else f"{listing_url}?paged={page_num}"
        print(f"    [Listing p{page_num}] {url}")

        try:
            await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await page.wait_for_timeout(2500)
        except Exception as e:
            print(f"    [WARN] Listing load failed: {e}")
            break

        # All anchors pointing to /product/ paths
        raw_hrefs: list[str] = await page.eval_on_selector_all(
            "a[href*='/product/']",
            "els => els.map(a => a.href)",
        )

        added = 0
        for href in raw_hrefs:
            href = href.split("?")[0].rstrip("/") + "/"
            if "/product/" in href and href not in seen:
                links.append(href)
                seen.add(href)
                added += 1

        print(f"    [Listing p{page_num}] +{added} new products (total {len(links)})")

        if added == 0:
            break

        # Detect next page
        has_next = await page.eval_on_selector_all(
            "a[href*='paged=']",
            f"els => els.some(a => a.href.includes('paged={page_num + 1}'))",
        )
        if not has_next:
            next_link = await page.query_selector("a.next, a[rel='next']")
            if next_link is None:
                break

        page_num += 1
        await async_polite_delay(1.5, 3.0)

    return links


# ── detail page scraper ──────────────────────────────────────────────────────
async def scrape_product(page, url: str) -> dict:
    """
    Navigate to a product detail page and collect every field available.
    Uses BeautifulSoup on the static HTML + JSON-LD schema for reliable extraction.
    Returns a flat dict — all keys go directly to the Excel row.
    """
    data: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)
    except Exception as e:
        print(f"    [WARN] Failed to load: {url} — {e}")
        return data

    html = await page.content()
    soup = BeautifulSoup(html, "lxml")

    # ── 1. JSON-LD Product schema (most reliable — present in static HTML) ──
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            d = json.loads(script.string or "{}")
            # Handle @graph wrapper
            if "@graph" in d:
                candidates = d["@graph"]
            elif isinstance(d, list):
                candidates = d
            else:
                candidates = [d]
            for obj in candidates:
                if obj.get("@type") == "Product":
                    name = clean_text(obj.get("name", ""))
                    if name:
                        data["Product Name"] = name.upper()
                        data["Product Family Id"] = extract_family_id(data["Product Name"])
                    if obj.get("description"):
                        data["Description"] = clean_text(obj["description"])
                    if obj.get("sku"):
                        data["SKU"] = str(obj["sku"])
                    offers = obj.get("offers", {})
                    if isinstance(offers, list):
                        offers = offers[0]
                    price = offers.get("price") or offers.get("lowPrice")
                    if price:
                        data["Price"] = clean_price(str(price))
                    break
        except Exception:
            pass

    # ── 2. Image URL — from product detail page gallery ──────────────────────
    for selector in [
        "div.woocommerce-product-gallery__image img",
        "figure.woocommerce-product-gallery__wrapper img",
        ".product-gallery img",
        "img.wp-post-image",
    ]:
        img_tag = soup.select_one(selector)
        if img_tag:
            src = (
                img_tag.get("data-large_image")
                or img_tag.get("nitro-lazy-src")
                or img_tag.get("data-src")
                or img_tag.get("src", "")
            )
            if src and not src.startswith("data:"):
                data["Image URL"] = src
                break

    # ── 2b. JSON-LD image fallback if gallery not found ──────────────────────
    if not data.get("Image URL"):
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                d = json.loads(script.string or "{}")
                candidates = d["@graph"] if "@graph" in d else ([d] if isinstance(d, dict) else d)
                for obj in candidates:
                    if obj.get("@type") == "Product":
                        img = obj.get("image")
                        if img:
                            data["Image URL"] = img if isinstance(img, str) else img[0]
                        break
            except Exception:
                pass
            if data.get("Image URL"):
                break

    # ── 3. Spec fields from .details .spec (h6=label, p=value) ──────────────
    spec_parts: list[str] = []
    for spec_el in soup.find_all("div", class_="spec"):
        h6 = spec_el.find("h6")
        p  = spec_el.find("p")
        if h6 and p:
            key   = clean_text(h6.get_text())
            value = clean_text(p.get_text())
            if key and value:
                spec_parts.append(f"{key}: {value}")
                parsed = parse_spec_block(f"{key}: {value}")
                for k, v in parsed.items():
                    data.setdefault(k, v)

    if spec_parts:
        data["Specifications"] = " | ".join(spec_parts)

    # ── 4. Parse dimensions into individual numeric fields ───────────────────
    dim_raw = data.get("Dimensions", "")
    if dim_raw:
        dims = parse_dimensions(dim_raw)
        for k, v in dims.items():
            data[k] = v

    # ── 5. Tearsheet link ────────────────────────────────────────────────────
    # Try .details .downloads area first, then any download link
    tearsheet_tag = None
    details_div = soup.find("div", class_="details")
    if details_div:
        tearsheet_tag = details_div.find("a", rel="download") or details_div.find(
            "a", href=re.compile(r"/tearsheet/", re.I)
        )
    if not tearsheet_tag:
        tearsheet_tag = soup.find("a", rel="download") or soup.find(
            "a", href=re.compile(r"/tearsheet/", re.I)
        )

    if tearsheet_tag:
        href = tearsheet_tag.get("href", "")
        data["Tearsheet Link"] = (BASE_URL + href) if href.startswith("/") else href
    else:
        slug = url.rstrip("/").split("/")[-1]
        data["Tearsheet Link"] = f"{BASE_URL}/tearsheet/product/{slug}"

    return data


# ── category scraper ─────────────────────────────────────────────────────────
async def scrape_category(
    page,
    writer: ExcelWriter,
    cat: dict,
    max_products: int | None = None,
):
    """
    Scrape one category.
    max_products: if set, stop after this many products (used in test mode).
    """
    name           = cat["name"]
    links          = cat["links"]
    studio_columns = cat.get("studio_columns", [])
    primary_link   = links[0] if links else ""

    mode_tag = f" [TEST: max {max_products} products]" if max_products else ""
    print(f"\n[Category] {name}{mode_tag}")
    writer.add_sheet(name, primary_link, studio_columns=studio_columns)

    # Collect all product URLs across all category links (e.g. "Link 2")
    all_urls: list[str] = []
    seen_urls: set[str] = set()
    for link in links:
        page_urls = await get_product_links(page, link)
        for u in page_urls:
            if u not in seen_urls:
                all_urls.append(u)
                seen_urls.add(u)
        # In test mode, stop collecting once we have enough
        if max_products and len(all_urls) >= max_products:
            break

    # Apply product cap
    if max_products:
        all_urls = all_urls[:max_products]

    print(f"  [Category] {len(all_urls)} products to scrape")

    vendor_name_env = os.environ.get("VENDOR_NAME", "The Future Perfect")

    for i, product_url in enumerate(all_urls, 1):
        print(f"  [{i}/{len(all_urls)}] {product_url}")
        try:
            data = await scrape_product(page, product_url)

            # ── Mandatory field: SKU ────────────────────────────────────
            # Generate if the site did not provide one
            if not data.get("SKU"):
                data["SKU"] = generate_sku(vendor_name_env, name, i)
                print(f"    [SKU generated] {data['SKU']}")

            # ── Mandatory field: Product Family Id ──────────────────────
            # Fallback: derive from Product Name if still missing
            if not data.get("Product Family Id") and data.get("Product Name"):
                data["Product Family Id"] = extract_family_id(data["Product Name"])

            writer.write_row(data, category_name=name)
        except Exception as e:
            print(f"  [ERROR] {product_url}: {e}")
        await async_polite_delay(1.0, 2.5)

    print(f"  [Category] Done — {len(all_urls)} rows buffered")


# ── main ──────────────────────────────────────────────────────────────────────
async def main():
    # Load vendor info (categories + studio_columns) written by orchestrator
    info_path = Path(__file__).parent / "vendor_info.json"
    if info_path.exists():
        vendor_info = json.loads(info_path.read_text(encoding="utf-8"))
    else:
        sys.path.insert(0, str(PROJECT_ROOT))
        from vendor_parser import parse_vendor
        vendor_info = parse_vendor(VENDOR_NAME)

    categories = vendor_info["categories"]

    # Apply test-mode category limit
    max_products: int | None = None
    if TEST_MODE:
        categories = [c for c in categories if c.get("links")][:TEST_MAX_CATEGORIES]
        max_products = TEST_MAX_PRODUCTS

    print(f"\n[Scraper] Vendor  : {vendor_info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")
    print(f"[Scraper] Headless: {HEADLESS}")
    print(f"[Scraper] Cats    : {len(categories)}"
          + (f" (capped at {TEST_MAX_CATEGORIES})" if TEST_MODE else ""))
    if TEST_MODE:
        print(f"[Scraper] Max products/cat: {TEST_MAX_PRODUCTS}")

    writer = ExcelWriter(OUTPUT_PATH, vendor_info["vendor_name"])

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        page.set_default_timeout(TIMEOUT_MS)

        for cat in categories:
            if not cat.get("links"):
                print(f"[Skip] {cat['name']} — no links")
                continue
            await scrape_category(page, writer, cat, max_products=max_products)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
