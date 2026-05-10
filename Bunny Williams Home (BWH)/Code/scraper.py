"""
scraper.py  —  Bunny Williams Home (BWH)
-----------------------------------------
Shopify store: bunnywilliamshome.com
Uses the Shopify /products.json API — no Playwright needed.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Bunny Williams Home (BWH)"
    python orchestrator.py "Bunny Williams Home (BWH)" --test
"""

from __future__ import annotations

import asyncio
import json
import os
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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Bunny Williams Home (BWH)")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.bunnywilliamshome.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


def clean_html(html: str) -> str:
    if not html:
        return ""
    return clean_text(BeautifulSoup(html, "html.parser").get_text(separator=" "))


def fetch_product_page_data(handle: str) -> dict:
    """Fetch the product detail page and extract accordion data.

    BWH stores dimensions and tear sheets in accordion items that are NOT
    present in the Shopify /products.json body_html — only on the rendered page.

    Returns dict with keys: 'dim_text' (raw string), 'Tearsheet Link' (URL or '').
    """
    result = {"dim_text": "", "Tearsheet Link": ""}
    try:
        r = requests.get(
            f"{BASE_URL}/products/{handle}",
            headers=HEADERS, timeout=20,
        )
        if r.status_code != 200:
            return result
        soup = BeautifulSoup(r.text, "html.parser")
        for item in soup.find_all("div", class_="accordion__item"):
            label_el = item.find("h4", class_="accordion__trigger")
            if not label_el:
                continue
            label = label_el.get_text(strip=True).lower()
            content = item.find("div", class_="accordion__content")
            if not content:
                continue
            if label == "dimensions":
                result["dim_text"] = content.get_text(separator=" ", strip=True)
            elif "tear" in label:
                # Prefer the "without pricing" tearsheet link
                for a in content.find_all("a", href=True):
                    href = a["href"]
                    if not result["Tearsheet Link"] or "pricing=false" in href:
                        result["Tearsheet Link"] = (
                            href if href.startswith("http") else BASE_URL + href
                        )
    except Exception as e:
        print(f"    Page-data error for {handle}: {e}")
    return result


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


def process_product(product: dict, page_data: dict | None = None) -> list[dict]:
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

    # Parse dimensions from the product detail page (not body_html)
    page_data = page_data or {}
    dim_text  = page_data.get("dim_text", "")
    dims = parse_dimensions(dim_text) if dim_text else {}
    dims.pop("Dimensions", None)

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
        row.update(dims)

        if page_data.get("Tearsheet Link"):
            row["Tearsheet Link"] = page_data["Tearsheet Link"]

        row["SKU"]   = variant.get("sku") or ""
        row["Price"] = clean_price(str(variant.get("price", "")))

        vi = variant.get("featured_image")
        if vi:
            row["Image URL"] = vi.get("src", default_img)
        elif variant.get("image_id") and variant["image_id"] in image_map:
            row["Image URL"] = image_map[variant["image_id"]]
        else:
            row["Image URL"] = default_img

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

        if vendor:
            row["Designer"] = vendor
        if tags:
            row["Collection"] = ", ".join(tags)

        rows.append(row)

    if rows:
        return rows
    fallback: dict = {
        "Source":            source_base,
        "Product Name":      title,
        "Product Family Id": extract_family_id(title),
        "Image URL":         default_img,
        "Description":       desc,
        **dims,
    }
    if page_data.get("Tearsheet Link"):
        fallback["Tearsheet Link"] = page_data["Tearsheet Link"]
    return [fallback]


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
