"""
scraper.py  —  Asthetic Decor
-------------------------------
Platform: aestheticdecor.com (WooCommerce / WordPress)
Uses requests + BeautifulSoup — no Playwright needed.

Site structure notes:
  - Listing pages (/tables/, /seating/, /cabinets/, /lighting/) contain all
    products on a single page with no pagination.
  - Some tracker categories link DIRECTLY to individual product pages
    (e.g. Lounge Chairs, Ottomans) — these are handled as single-product URLs.
  - Multiple tracker categories may share the same listing URL (/tables/ holds
    Coffee Tables, Side Tables, Dining Tables, Consoles, Desks). A global
    dedup set prevents scraping the same product twice.
  - Products use WooCommerce structure: JSON-LD for core fields + HTML spec
    table for dimensions/materials.
  - No interactive variant selectors — options are listed as static text.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Asthetic Decor"
    python orchestrator.py "Asthetic Decor" --test
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
    parse_spec_block,
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Asthetic Decor")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.aestheticdecor.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


def get_html(url: str) -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        if r.status_code != 200:
            print(f"    HTTP {r.status_code}: {url}")
            return None
        return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        print(f"    Fetch error {url}: {e}")
        return None


def get_product_links_from_listing(listing_url: str) -> list[str]:
    """
    Extract all /product/ hrefs from a WooCommerce listing page.
    These pages have no pagination — all products appear on one page.
    """
    soup = get_html(listing_url)
    if not soup:
        return []

    links: list[str] = []
    seen: set[str] = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "/product/" in href:
            # Normalise to absolute
            if href.startswith("/"):
                href = BASE_URL + href
            href = href.rstrip("/") + "/"
            if href not in seen:
                seen.add(href)
                links.append(href)
    return links


def scrape_product(product_url: str) -> dict:
    """
    Scrape one WooCommerce product detail page.
    Uses JSON-LD for core fields + HTML spec table for extra fields.
    Returns a flat dict.
    """
    data: dict = {"Source": product_url}
    soup = get_html(product_url)
    if not soup:
        return data

    # ── 1. JSON-LD ──────────────────────────────────────────────────────────
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            obj = json.loads(script.string or "{}")
            # Unwrap @graph if present
            candidates = obj.get("@graph", [obj]) if isinstance(obj, dict) else [obj]
            for item in candidates:
                if item.get("@type") == "Product":
                    if item.get("name"):
                        data["Product Name"] = clean_text(item["name"])
                        data["Product Family Id"] = extract_family_id(data["Product Name"])
                    if item.get("description"):
                        data["Description"] = clean_text(item["description"])
                    if item.get("sku"):
                        data["SKU"] = str(item["sku"])
                    # Price from offers
                    offers = item.get("offers", {})
                    if isinstance(offers, list):
                        offers = offers[0] if offers else {}
                    price = offers.get("price") or offers.get("lowPrice")
                    if price:
                        data["Price"] = clean_price(str(price))
                    # Image
                    img = item.get("image")
                    if img:
                        data["Image URL"] = img if isinstance(img, str) else img[0]
                    break
        except Exception:
            pass

    # ── 2. Fallback: product name from h1 ───────────────────────────────────
    if not data.get("Product Name"):
        h1 = soup.find("h1", class_=re.compile(r"product[_-]title|entry-title", re.I))
        if not h1:
            h1 = soup.find("h1")
        if h1:
            data["Product Name"] = clean_text(h1.get_text())
            data["Product Family Id"] = extract_family_id(data["Product Name"])

    # ── 3. Gallery image (higher-res from WooCommerce gallery) ──────────────
    if not data.get("Image URL"):
        for selector in [
            ".woocommerce-product-gallery__image img",
            ".wp-post-image",
            ".product-image img",
        ]:
            img_tag = soup.select_one(selector)
            if img_tag:
                src = (
                    img_tag.get("data-large_image")
                    or img_tag.get("data-src")
                    or img_tag.get("src", "")
                )
                if src and not src.startswith("data:"):
                    data["Image URL"] = src
                    break

    # ── 4. Specification table ───────────────────────────────────────────────
    # Asthetic Decor uses an HTML table for specs: row = (label, value)
    spec_parts: list[str] = []
    spec_table = soup.find("table", class_=re.compile(r"woocommerce-product-attributes|shop_attributes", re.I))
    if not spec_table:
        # fallback: look for any table inside .product or .entry-content
        spec_table = soup.select_one(".product-details table, .entry-content table, .product table")

    if spec_table:
        for row in spec_table.find_all("tr"):
            cells = row.find_all(["th", "td"])
            if len(cells) >= 2:
                key   = clean_text(cells[0].get_text())
                value = clean_text(cells[1].get_text())
                if key and value:
                    spec_parts.append(f"{key}: {value}")
                    key_lower = key.lower()
                    # Map common spec keys to canonical columns
                    if "dimension" in key_lower or "size" in key_lower:
                        dims = parse_dimensions(value)
                        for k, v in dims.items():
                            data.setdefault(k, v)
                    elif "wood" in key_lower:
                        data.setdefault("Wood", value)
                    elif "finish" in key_lower:
                        data.setdefault("Finish", value)
                    elif "material" in key_lower:
                        data.setdefault("Materials", value)
                    elif "upholster" in key_lower or "uph" in key_lower:
                        data.setdefault("Upholstery", value)
                    elif "com" == key_lower.strip():
                        data.setdefault("COM", value)
                    elif "rush" in key_lower:
                        data.setdefault("Rush", value)
                    elif "weight" in key_lower:
                        data.setdefault("Weight", safe_float(re.sub(r"[^\d.]", "", value)))
                    else:
                        # Keep unknown keys directly
                        data.setdefault(key, value)

    # Also try a div-based spec block if no table found
    if not spec_parts:
        for spec_div in soup.find_all(["div", "section"], class_=re.compile(r"spec|detail|attribute", re.I)):
            text = spec_div.get_text(separator=" | ")
            parsed = parse_spec_block(text)
            for k, v in parsed.items():
                data.setdefault(k, v)
                spec_parts.append(f"{k}: {v}")

    if spec_parts:
        data["Specifications"] = " | ".join(spec_parts)

    # ── 5. Dimensions from description if not already set ───────────────────
    if not data.get("Dimensions") and data.get("Description"):
        dims = parse_dimensions(data["Description"])
        for k, v in dims.items():
            data.setdefault(k, v)

    return data


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

    # Global dedup: multiple categories share the same listing URL
    global_seen_urls: set[str] = set()
    global_idx = 1

    for cat in categories:
        if not cat["links"]:
            continue

        cat_url = cat["links"][0]
        writer.add_sheet(cat["name"], cat_url, studio_columns=cat["studio_columns"])

        # Collect product URLs for this category
        product_urls: list[str] = []
        for link in cat["links"]:
            if "/product/" in link:
                # Direct product URL (e.g. Lounge Chairs, Ottomans in tracker)
                norm = link.rstrip("/") + "/"
                if norm not in global_seen_urls:
                    product_urls.append(norm)
            else:
                # Listing page — extract all product links
                for u in get_product_links_from_listing(link):
                    if u not in global_seen_urls:
                        product_urls.append(u)

        if TEST_MODE:
            product_urls = product_urls[:TEST_MAX_PRODUCTS]

        print(f"  {cat['name']}: {len(product_urls)} products")

        for product_url in product_urls:
            global_seen_urls.add(product_url)
            try:
                data = scrape_product(product_url)
                if not data.get("SKU"):
                    data["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                if not data.get("Product Family Id") and data.get("Product Name"):
                    data["Product Family Id"] = extract_family_id(data["Product Name"])
                writer.write_row(data, category_name=cat["name"])
                global_idx += 1
            except Exception as e:
                print(f"    ERROR {product_url}: {e}")
            time.sleep(0.8)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
