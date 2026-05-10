"""
scraper.py  —  Century
------------------------
Shopify store: shop.centuryfurniture.com
Uses the Shopify /products.json API with client-side product_type filtering.

The tracker category URLs include Shopify storefront filter params like:
  ?filter.p.product_type=Nightstands
These are not supported by the products.json API, so we:
  1. Extract the collection handle from the URL path
  2. Extract the product_type filter values from query params
  3. Fetch all products in that collection
  4. Filter client-side by matching product.product_type

NOTE: The "Accessories > Table Lamps" category links to theodorealexander.com —
a different vendor. That category is automatically skipped since its URL host
doesn't match BASE_URL.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Century"
    python orchestrator.py "Century" --test
"""

from __future__ import annotations

import asyncio
import json
import os
import sys
import time
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import re

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
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Century")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://shop.centuryfurniture.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


# Matches Century's inline dimension lines: "Overall Height: 30 in." etc.
_DIM_LINE_RE = re.compile(
    r"\b(?:Overall\s+)?\w[\w\s]*:\s*[\d./]+\s*(?:in|lbs?|kg)\.?",
    re.IGNORECASE,
)

# Specific pattern for extracting dim values from body_html text
_INLINE_DIM_RE = re.compile(
    r"(?:Overall\s+)?(Height|Width|Depth|Length|Diameter|Weight)"
    r"\s*:\s*([\d./]+)\s*(in|lbs?|kg)?\.?",
    re.IGNORECASE,
)


def clean_html(html: str) -> str:
    if not html:
        return ""
    text = clean_text(BeautifulSoup(html, "html.parser").get_text(separator=" "))
    # Strip any "Label: N unit" lines that belong in dimension columns
    text = _DIM_LINE_RE.sub("", text).strip()
    return text


def parse_inline_dims(html: str) -> dict:
    """Extract Height/Width/Depth/Weight etc. from body_html inline text.

    Century includes these in the Shopify body_html as plain lines like:
      'Overall Height: 30 in.  Overall Depth: 20 in.  Overall Width: 32 in.'
    Parsing here avoids a separate HTTP request and works even when the
    accordion is empty or JavaScript-rendered.
    """
    if not html:
        return {}
    text = BeautifulSoup(html, "html.parser").get_text(separator=" ")
    result: dict = {}
    for m in _INLINE_DIM_RE.finditer(text):
        col = m.group(1).strip().title()   # "Height", "Width", etc.
        val = safe_float(m.group(2).strip())
        if val is not None:
            result[col] = val
    return result


# Maps Century dimension label prefixes (lowercase) → output column names
_DIM_LABEL_MAP = {
    "overall height":   "Height",
    "overall width":    "Width",
    "overall depth":    "Depth",
    "overall length":   "Length",
    "overall diameter": "Diameter",
    "seat height":      "Seat Height",
    "seat depth":       "Seat Depth",
    "seat width":       "Seat Width",
    "arm height":       "Arm Height",
    "weight":           "Weight",
}


def fetch_product_page_data(handle: str) -> dict:
    """Fetch the Century product detail page and extract accordion data.

    Century stores dimensions (and sometimes materials) in <details>/<summary>
    accordions that are NOT in the Shopify /products.json body_html.

    Returns a dict with dimension fields (Height, Width, Depth, Weight, …)
    and optionally 'Material'.
    """
    result: dict = {}
    try:
        r = requests.get(
            f"{BASE_URL}/products/{handle}",
            headers=HEADERS, timeout=20,
        )
        if r.status_code != 200:
            return result
        soup = BeautifulSoup(r.text, "html.parser")

        for details in soup.find_all("details"):
            summary = details.find("summary")
            if not summary:
                continue
            label = summary.get_text(strip=True).lower()
            content = details.find("div", class_="accordion__content")
            if not content:
                continue

            if label == "dimensions":
                # Each spec is a <p>Label: Value unit.</p>
                for p in content.find_all("p"):
                    raw = p.get_text(strip=True)
                    if ":" not in raw:
                        continue
                    key_raw, _, val_raw = raw.partition(":")
                    key_lower = key_raw.strip().lower()
                    val_str = val_raw.strip().rstrip(".")
                    # Strip unit suffixes (" in", " lbs", " kg")
                    for unit in (" in", " lbs", " kg", " lb"):
                        val_str = val_str.removesuffix(unit)
                    val_str = val_str.strip()
                    col = _DIM_LABEL_MAP.get(key_lower)
                    if col:
                        num = safe_float(val_str)
                        result[col] = num if num is not None else val_str

            elif label == "materials":
                text = content.get_text(separator=" ", strip=True)
                if text:
                    result["Material"] = clean_text(text)

    except Exception as e:
        print(f"    Page-data error for {handle}: {e}")
    return result


def parse_product_types(url: str) -> list[str]:
    """Extract all filter.p.product_type values from a Shopify filter URL."""
    parsed = urlparse(url)
    params = parse_qs(parsed.query)
    types: list[str] = []
    for key, values in params.items():
        if "product_type" in key:
            types.extend(values)
    return [t.strip() for t in types if t.strip()]


def get_collection_products(collection_handle: str) -> list[dict]:
    products: list[dict] = []
    page = 1
    while True:
        url = f"{BASE_URL}/collections/{collection_handle}/products.json"
        try:
            r = requests.get(
                url, params={"limit": 250, "page": page},
                headers=HEADERS, timeout=25,
            )
            if r.status_code != 200:
                print(f"    Collection {collection_handle} returned {r.status_code}")
                break
            data = r.json().get("products", [])
            if not data:
                break
            products.extend(data)
            if len(data) < 250:
                break
            page += 1
            time.sleep(0.5)
        except Exception as e:
            print(f"    Listing error page {page}: {e}")
            break
    return products


def get_products_from_url(listing_url: str) -> list[dict]:
    """
    Extract handle + product_type filters from the URL, fetch from API,
    then filter by product type if filters are present.
    Skips URLs from other domains (e.g. theodorealexander.com).
    """
    parsed = urlparse(listing_url)
    if parsed.netloc and BASE_URL not in parsed.netloc and parsed.netloc not in BASE_URL:
        print(f"    Skipping external URL: {listing_url}")
        return []

    path = listing_url.rstrip("/")
    if "/collections/" not in path:
        print(f"    Unknown URL pattern, skipping: {listing_url}")
        return []

    # Extract collection handle (path segment after /collections/)
    handle = path.split("/collections/")[-1].split("?")[0].split("/")[0]
    products = get_collection_products(handle)

    # Apply product_type filter if present in URL
    type_filters = parse_product_types(listing_url)
    if type_filters:
        # Case-insensitive match against any of the filter values
        type_filters_lower = {t.lower() for t in type_filters}
        products = [
            p for p in products
            if p.get("product_type", "").strip().lower() in type_filters_lower
        ]
        print(f"    After type filter {type_filters}: {len(products)} products")

    return products


def process_product(product: dict, page_data: dict | None = None) -> list[dict]:
    title     = product.get("title", "")
    body_html = product.get("body_html", "")
    handle    = product.get("handle", "")
    source_base = f"{BASE_URL}/products/{handle}"
    vendor    = product.get("vendor", "")
    tags      = product.get("tags", [])
    prod_type = product.get("product_type", "")

    image_map   = {img["id"]: img["src"] for img in product.get("images", [])}
    all_images  = [img["src"] for img in product.get("images", [])]
    default_img = all_images[0] if all_images else ""

    # Parse dims from body_html first (always available from Shopify API).
    # page_data (accordion) takes priority when both are present.
    body_dims = parse_inline_dims(body_html)
    page_data = page_data or {}
    merged_dims = {**body_dims, **page_data}  # accordion wins on conflicts

    # Clean description AFTER extracting dims so the strip regex removes inline lines
    desc = clean_html(body_html)

    options  = product.get("options", [])
    variants = product.get("variants", [])
    rows: list[dict] = []

    for variant in variants:
        variant_id  = variant.get("id")
        variant_url = f"{source_base}?variant={variant_id}" if variant_id else source_base

        row: dict = {
            "Source":            variant_url,
            "Product Name":      title,
            "Product Family Id": extract_family_id(title),
            "Description":       desc,
        }
        row.update(merged_dims)

        row["SKU"]   = variant.get("sku") or ""
        row["Price"] = clean_price(str(variant.get("price", "")))

        vi = variant.get("featured_image")
        if vi:
            row["Image URL"] = vi.get("src", default_img)
        elif variant.get("image_id") and variant["image_id"] in image_map:
            row["Image URL"] = image_map[variant["image_id"]]
        else:
            row["Image URL"] = default_img

        # Only use grams if no Weight found from page data or body dims
        grams = variant.get("grams")
        if grams and not merged_dims.get("Weight"):
            row["Weight"] = round(grams / 453.592, 2)

        opt1 = variant.get("option1", "")
        opt2 = variant.get("option2", "")
        opt3 = variant.get("option3", "")
        if opt1 and opt1 not in ("Default Title", ""):
            key = options[0]["name"] if options else "Finish"
            row[key] = opt1
        if opt2 and opt2 not in ("Default Title", ""):
            key = options[1]["name"] if len(options) > 1 else "Size"
            row[key] = opt2
        if opt3 and opt3 not in ("Default Title", ""):
            key = options[2]["name"] if len(options) > 2 else "Option 3"
            row[key] = opt3

        if vendor:
            row["Designer"] = vendor
        if prod_type:
            row["Product Type"] = prod_type
        if tags:
            row["Collection"] = ", ".join(tags)

        rows.append(row)

    if rows:
        return rows
    return [{
        "Source":            source_base,
        "Product Name":      title,
        "Product Family Id": extract_family_id(title),
        "Image URL":         default_img,
        "Description":       desc,
        **merged_dims,
    }]


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

    for cat in categories:
        if not cat["links"]:
            continue

        # Skip external-domain categories (e.g. Table Lamps → theodorealexander.com)
        primary = cat["links"][0]
        if BASE_URL.replace("https://", "") not in primary and "centuryfurniture" not in primary:
            print(f"  [Skip] {cat['name']} — external URL: {primary}")
            continue

        writer.add_sheet(cat["name"], primary, studio_columns=cat["studio_columns"])

        seen_handles: set[str] = set()
        products: list[dict]   = []
        for listing_url in cat["links"]:
            for p in get_products_from_url(listing_url):
                h = p.get("handle", "")
                if h:
                    if h in seen_handles:
                        continue
                    seen_handles.add(h)
                products.append(p)

        if TEST_MODE:
            products = products[:TEST_MAX_PRODUCTS]

        print(f"  {cat['name']}: {len(products)} products across {len(cat['links'])} link(s)")

        global_idx = 1
        for product in products:
            try:
                page_data = fetch_product_page_data(product.get("handle", ""))
                time.sleep(0.5)
                rows = process_product(product, page_data)
                for row in rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
            except Exception as e:
                print(f"    ERROR on {product.get('handle', '?')}: {e}")

        time.sleep(1.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
