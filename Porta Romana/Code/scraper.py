"""
scraper.py  —  Porta Romana
-----------------------------
Shopify store: portaromana.com
Uses the Shopify /products.json API — no Playwright needed.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Porta Romana"
    python orchestrator.py "Porta Romana" --test
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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Porta Romana")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://portaromana.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


def _imperial_frac(s: str) -> str:
    """Parse Porta Romana imperial dimension strings: '9 3/4' → '9.75', '24' → '24'"""
    s = s.strip()
    m = re.match(r'^(\d+)\s+(\d+)/(\d+)$', s)
    if m:
        val = int(m.group(1)) + int(m.group(2)) / int(m.group(3))
        r = str(round(val, 4))
        return r.rstrip('0').rstrip('.')
    m = re.match(r'^(\d+)/(\d+)$', s)
    if m:
        val = int(m.group(1)) / int(m.group(2))
        r = str(round(val, 4))
        return r.rstrip('0').rstrip('.')
    try:
        val = float(s)
        r = str(val)
        return r.rstrip('0').rstrip('.')
    except ValueError:
        return s


_DIM_KEYS = {"height": "Height", "width": "Width", "depth": "Depth",
             "length": "Length", "diameter": "Diameter"}

_SKIP_KEYS = {"cat", "shippingcat", "card-width", "height-range",
              "width-range", "hide-from-search", "main-cat", "maincat",
              "related", "colour", "grouping"}


def parse_product_tags(tags: list[str]) -> dict:
    """
    Parse Porta Romana structured tags (Key:Value) into proper field dict.
    Dimension tags contain both mm and imperial: '615mm | 24 1/4"' — we take inches.
    """
    row: dict = {}
    dim_parts: list[str] = []

    for tag in tags:
        if ":" not in tag:
            continue
        raw_key, _, val = tag.partition(":")
        key = raw_key.strip().lower()
        val = val.strip()

        if key in _SKIP_KEYS:
            continue

        if key == "collection":
            row["Collection"] = val

        elif key in ("material",):
            row["Material"] = val

        elif key == "constructedfrom":
            if not row.get("Material"):
                row["Material"] = val

        elif key == "designedby":
            row["Designer"] = val

        elif key == "craterequired":
            row.setdefault("Specifications", "")
            if val:
                existing = row.get("Specifications", "")
                row["Specifications"] = (existing + "\n" + val).strip()

        elif key == "weight":
            # "2 kg | 4 1/2 lbs" → extract lbs
            lbs_m = re.search(
                r'\|\s*([\d]+(?:\s+\d+/\d+)?(?:\.\d+)?)\s*lbs?', val, re.IGNORECASE
            )
            if lbs_m:
                row["Weight"] = _imperial_frac(lbs_m.group(1))
            else:
                # kg only — convert
                kg_m = re.search(r'([\d]+(?:\.\d+)?)\s*kg', val, re.IGNORECASE)
                if kg_m:
                    row["Weight"] = str(round(float(kg_m.group(1)) * 2.20462, 2))

        elif key in _DIM_KEYS:
            col = _DIM_KEYS[key]
            # Pattern: "615mm | 24 1/4""  or just "245mm"
            imp_m = re.search(
                r'\|\s*([\d]+(?:\s+\d+/\d+)?(?:\.\d+)?)\s*"', val
            )
            if imp_m:
                parsed = _imperial_frac(imp_m.group(1))
                row[col] = parsed
                dim_parts.append(f"{parsed} {col[0].lower()}")
            else:
                # mm only → convert to inches
                mm_m = re.search(r'([\d]+(?:\.\d+)?)\s*mm', val)
                if mm_m:
                    parsed = str(round(float(mm_m.group(1)) / 25.4, 2))
                    row[col] = parsed
                    dim_parts.append(f"{parsed} {col[0].lower()}")

    # Build Dimensions summary string from collected parts
    if dim_parts:
        row["Dimensions"] = " x ".join(dim_parts)

    return row


def clean_html(html: str) -> str:
    if not html:
        return ""
    return clean_text(BeautifulSoup(html, "html.parser").get_text(separator=" "))


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

    image_map   = {img["id"]: img["src"] for img in product.get("images", [])}
    all_images  = [img["src"] for img in product.get("images", [])]
    default_img = all_images[0] if all_images else ""

    # Parse structured data out of Shopify tags
    tag_fields = parse_product_tags(tags)

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
        row.update(tag_fields)

        row["SKU"]   = variant.get("sku") or ""
        row["Price"] = clean_price(str(variant.get("price", "")))

        vi = variant.get("featured_image")
        if vi:
            row["Image URL"] = vi.get("src", default_img)
        elif variant.get("image_id") and variant["image_id"] in image_map:
            row["Image URL"] = image_map[variant["image_id"]]
        else:
            row["Image URL"] = default_img

        # Weight from tags takes priority; fall back to Shopify grams if missing
        if not row.get("Weight"):
            grams = variant.get("grams")
            if grams:
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

        # vendor field on Porta Romana products is always "Porta Romana" — skip Designer
        rows.append(row)

    return rows if rows else [{
        "Source":            source_base,
        "Product Name":      title,
        "Product Family Id": extract_family_id(title),
        "Image URL":         default_img,
        "Description":       desc,
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
