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
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "English Georgian American")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://englishgeorgianamerica.com"
HEADERS = {
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


def get_collection_products(base_url: str, collection_handle: str) -> list[dict]:
    products = []
    page = 1
    while True:
        url = f"{base_url}/collections/{collection_handle}/products.json"
        try:
            r = requests.get(url, params={"limit": 250, "page": page}, headers=HEADERS, timeout=25)
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

    handle = path.split("/collections/")[-1].split("/")[0]
    return get_collection_products(BASE_URL, handle)


def extract_dimensions(description: str) -> dict:
    dims = parse_dimensions(description)
    dims.pop("Dimensions", None)
    if dims:
        return dims

    # Vendor descriptions often use trailing labels like:
    # 'Dimensions: 26" W x 18.75" D x 29.5" H'
    matches = re.findall(
        r'(\d+(?:\.\d+)?)\s*(?:"|in\.?)?\s*(W|D|H|L|Dia|Diam|Diameter)\b',
        description,
        flags=re.IGNORECASE,
    )
    if not matches:
        return {}

    key_map = {
        "w": "Width",
        "d": "Depth",
        "h": "Height",
        "l": "Length",
        "dia": "Diameter",
        "diam": "Diameter",
        "diameter": "Diameter",
    }
    parsed = {}
    for value, label in matches:
        key = key_map.get(label.lower())
        if key:
            parsed[key] = value
    return parsed


def process_product(product: dict) -> list[dict]:
    title = product.get("title", "")
    body_html = product.get("body_html", "")
    description = clean_html(body_html)
    handle = product.get("handle", "")
    source_url = f"{BASE_URL}/products/{handle}"
    vendor = product.get("vendor", "")
    tags = product.get("tags", [])

    # Parse dimensions from description, including trailing-label formats used on this site.
    dims = extract_dimensions(description)

    # Images list
    product_images = [img["src"] for img in product.get("images", [])]
    default_image = product_images[0] if product_images else ""

    # Collection tags
    collection_val = ", ".join(tags) if tags else ""

    rows = []
    variants = product.get("variants", [])
    options = product.get("options", [])

    for variant in variants:
        variant_id = variant.get("id")
        variant_url = f"{source_url}?variant={variant_id}" if variant_id else source_url
        row = {
            "Source": variant_url,
            "Product Name": title,
            "Product Family Id": extract_family_id(title),
            "Description": description,
        }

        # Dimensions
        row.update(dims)

        # SKU and Price
        row["SKU"] = variant.get("sku") or ""
        price_raw = str(variant.get("price", ""))
        row["Price"] = clean_price(price_raw)

        # Image URL
        vi = variant.get("featured_image")
        row["Image URL"] = vi["src"] if vi else default_image

        # Weight (grams → lbs)
        grams = variant.get("grams")
        if grams:
            row["Weight"] = round(grams / 453.592, 2)

        # Variant options → use option names from product.options
        opt1 = variant.get("option1", "")
        opt2 = variant.get("option2", "")
        opt3 = variant.get("option3", "")
        if opt1 and opt1 not in ("Default Title", ""):
            key = options[0]["name"] if options else "Option 1"
            row[key] = opt1
        if opt2 and opt2 not in ("Default Title", ""):
            key = options[1]["name"] if len(options) > 1 else "Option 2"
            row[key] = opt2
        if opt3 and opt3 not in ("Default Title", ""):
            key = options[2]["name"] if len(options) > 2 else "Option 3"
            row[key] = opt3

        # Collection / vendor
        if collection_val:
            row["Collection"] = collection_val
        if vendor:
            row["Designer"] = vendor

        rows.append(row)

    return rows if rows else [{
        "Source": source_url,
        "Product Name": title,
        "Product Family Id": extract_family_id(title),
        "Image URL": default_image,
        "Description": description,
    }]


async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    for cat in categories:
        if not cat["links"]:
            continue

        cat_url = cat["links"][0]
        writer.add_sheet(cat["name"], cat_url, studio_columns=cat["studio_columns"])

        seen_handles: set[str] = set()
        products: list[dict] = []
        for listing_url in cat["links"]:
            for product in get_products_from_url(listing_url):
                handle = product.get("handle", "")
                if handle:
                    if handle in seen_handles:
                        continue
                    seen_handles.add(handle)
                products.append(product)

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
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
