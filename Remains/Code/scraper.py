"""
scraper.py  —  Remains
-----------------------
Shopify store: remains.com
Uses the Shopify /products.json API — no Playwright needed.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Remains"
    python orchestrator.py "Remains" --test

Env vars (set by orchestrator):
    HEADLESS             true | false   (default: true, unused here)
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
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Remains")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://remains.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


def clean_html(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return clean_text(soup.get_text(separator=" "))


def _mixed_frac(s: str) -> float:
    """'1-1/4' → 1.25, '1/4' → 0.25, '26' → 26.0"""
    s = s.strip()
    m = re.match(r'^(\d+)-(\d+)/(\d+)$', s)
    if m:
        return int(m.group(1)) + int(m.group(2)) / int(m.group(3))
    m = re.match(r'^(\d+)/(\d+)$', s)
    if m:
        return int(m.group(1)) / int(m.group(2))
    try:
        return float(s)
    except ValueError:
        return 0.0


def _fmt(val: float) -> str:
    s = str(round(val, 4))
    if '.' in s:
        s = s.rstrip('0').rstrip('.')
    return s


def parse_remains_dim_line(line: str) -> dict:
    """
    Parse 'Overall: 26" h. x 20" w. x 1-1/4" d.' or similar.
    Returns dict with Dimensions, Height, Width, Depth, Diameter, Length.
    """
    body = re.sub(r'^Overall\s*[:\.]?\s*', '', line, flags=re.IGNORECASE).strip()
    result: dict = {}

    dim_clean = re.sub(r'"', '', body)
    dim_clean = re.sub(r'\s+', ' ', dim_clean).strip()
    if dim_clean:
        result["Dimensions"] = dim_clean

    label_map = {
        "h": "Height", "w": "Width", "d": "Depth",
        "l": "Length", "dia": "Diameter",
    }
    pat = re.compile(
        r'(\d+(?:-\d+/\d+|/\d+)?(?:\.\d+)?)\s*"?\s*(h|w|d|l|dia(?:m(?:eter)?)?)\.',
        re.IGNORECASE,
    )
    for m in pat.finditer(body):
        val_str = m.group(1)
        lbl = m.group(2).lower()
        if lbl.startswith("dia"):
            lbl = "dia"
        key = label_map.get(lbl)
        if key and key not in result:
            result[key] = _fmt(_mixed_frac(val_str))

    return result


def fetch_product_page_specs(handle: str) -> tuple[dict, str]:
    """
    Fetch the product HTML page and extract the SPECIFICATIONS section.
    Returns (dims_dict, specs_text_string).
    """
    url = f"{BASE_URL}/products/{handle}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        if r.status_code != 200:
            return {}, ""
        soup = BeautifulSoup(r.text, "html.parser")
        page_text = soup.get_text(separator="\n")

        spec_match = re.search(
            r'SPECIFICATIONS\s*\n(.*?)(?=\n[ \t]*(?:SHIPPING|AVAILABILITY|OPTIONS|CONTACT US)\s*\n)',
            page_text, re.DOTALL | re.IGNORECASE,
        )
        if not spec_match:
            return {}, ""

        specs_block = spec_match.group(1)
        lines = [ln.strip() for ln in specs_block.split('\n') if ln.strip()]
        specs_text = "\n".join(lines)

        dims: dict = {}
        for ln in lines:
            if re.match(r'overall\s*[:\.]?', ln, re.IGNORECASE):
                dims = parse_remains_dim_line(ln)
                break

        # Fallback: no "Overall" label — try first dimension-looking line
        if not dims:
            for ln in lines:
                if re.search(r'\d+(?:-\d+/\d+|/\d+)?(?:\.\d+)?\s*"?\s*[hwdl]\.', ln, re.IGNORECASE):
                    dims = parse_remains_dim_line(ln)
                    break

        return dims, specs_text
    except Exception as exc:
        print(f"    Page spec error for {handle}: {exc}")
        return {}, ""


def get_collection_products(collection_handle: str) -> list[dict]:
    """Fetch all products from a Shopify collection via the JSON API."""
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
    path = listing_url.rstrip("/")
    if "/collections/" not in path:
        print(f"    Unknown URL pattern, skipping: {listing_url}")
        return []
    handle = path.split("/collections/")[-1].split("?")[0].split("/")[0]
    return get_collection_products(handle)


def process_product(product: dict) -> list[dict]:
    title     = product.get("title", "")
    body_html = product.get("body_html", "")
    desc      = clean_html(body_html)
    handle    = product.get("handle", "")
    source_base = f"{BASE_URL}/products/{handle}"
    vendor    = product.get("vendor", "")
    tags      = product.get("tags", [])

    # Images
    image_map   = {img["id"]: img["src"] for img in product.get("images", [])}
    all_images  = [img["src"] for img in product.get("images", [])]
    default_img = all_images[0] if all_images else ""

    # Fetch specs & dimensions from the product detail page (not in Shopify API body_html)
    dims, specs_text = fetch_product_page_specs(handle)
    time.sleep(0.4)

    options  = product.get("options", [])
    variants = product.get("variants", [])
    rows: list[dict] = []

    for variant in variants:
        variant_id  = variant.get("id")
        variant_url = f"{source_base}?variant={variant_id}" if variant_id else source_base

        row: dict = {
            "Source":           variant_url,
            "Product Name":     title,
            "Product Family Id": extract_family_id(title),
            "Description":      desc,
        }
        row.update(dims)
        if specs_text:
            row["Specifications"] = specs_text

        # SKU & Price
        row["SKU"]   = variant.get("sku") or ""
        row["Price"] = clean_price(str(variant.get("price", "")))

        # Image: variant-specific → product default
        vi = variant.get("featured_image")
        if vi:
            row["Image URL"] = vi.get("src", default_img)
        elif variant.get("image_id") and variant["image_id"] in image_map:
            row["Image URL"] = image_map[variant["image_id"]]
        else:
            row["Image URL"] = default_img

        # Weight (grams → lbs)
        grams = variant.get("grams")
        if grams:
            row["Weight"] = round(grams / 453.592, 2)

        # Variant options
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

        # Metadata
        if vendor:
            row["Designer"] = vendor
        if tags:
            row["Collection"] = ", ".join(tags)

        rows.append(row)

    return rows if rows else [{
        "Source":           source_base,
        "Product Name":     title,
        "Product Family Id": extract_family_id(title),
        "Image URL":        default_img,
        "Description":      desc,
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

        cat_url = cat["links"][0]
        writer.add_sheet(cat["name"], cat_url, studio_columns=cat["studio_columns"])

        # Collect products from all links, dedup by handle
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
                rows = process_product(product)
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
