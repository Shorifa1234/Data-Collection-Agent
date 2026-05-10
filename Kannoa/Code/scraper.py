"""
scraper.py  —  Kannoa
-----------------------
Shopify store: kannoa.com  (outdoor furniture)
Uses the Shopify /products.json API — no Playwright needed.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Kannoa"
    python orchestrator.py "Kannoa" --test
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
    sentence_case,
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Kannoa")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.kannoa.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


# Maps Kannoa spec labels (lowercase) → output column names
_SPEC_LABEL_MAP = {
    "product height": "Height",
    "product width":  "Width",
    "depth":          "Depth",
    "arm height":     "Arm Height",
    "seat height":    "Seat Height",
    "seat depth":     "Seat Depth",
    "seat width":     "Seat Width",
    "diameter":       "Diameter",
    "length":         "Length",
    "weight":         "Weight",
}
_DIMENSION_FIELDS = {
    "Height", "Width", "Depth", "Arm Height", "Seat Height",
    "Seat Depth", "Seat Width", "Diameter", "Length",
}
_SKIP_SPEC_KEYS = {"manufacturer", "brand", "sku"}


def parse_body_html(body_html: str) -> tuple[str, dict]:
    """Return (description_text, spec_dict) parsed from Kannoa body_html.

    Description = paragraph text before the SPECIFICATIONS heading.
    Spec dict   = labeled <li>Key: Value</li> items from the specs table,
                  with dimension values as floats and label remapping applied.
    """
    soup = BeautifulSoup(body_html or "", "html.parser")

    # Description: paragraphs before <h5>SPECIFICATIONS</h5>
    desc_parts: list[str] = []
    for el in soup.find_all(["p", "h5"]):
        if el.name == "h5" and "SPECIFICATIONS" in el.get_text().upper():
            break
        text = el.get_text(separator=" ").strip()
        if text:
            desc_parts.append(text)
    description = clean_text(" ".join(desc_parts))

    specs: dict = {}

    # Parse <li>Label: Value</li> items from the specs table
    spec_table = soup.find("table", {"id": "prod"})
    li_elements = spec_table.find_all("li") if spec_table else []

    # Fallback: scan any <ul> after the SPECIFICATIONS heading
    if not li_elements:
        h5 = soup.find("h5", string=lambda t: t and "SPECIFICATIONS" in t.upper())
        if h5:
            for ul in h5.find_next_siblings("ul"):
                li_elements.extend(ul.find_all("li"))

    for li in li_elements:
        raw = li.get_text().strip()
        if ":" not in raw:
            continue
        label, _, value = raw.partition(":")
        label = label.strip()
        value = value.strip()
        if not label or not value:
            continue
        key_lower = label.lower()
        if key_lower in _SKIP_SPEC_KEYS:
            continue
        col = _SPEC_LABEL_MAP.get(key_lower, sentence_case(label))
        if col in _DIMENSION_FIELDS:
            numeric = safe_float(value)
            if numeric is not None:
                specs[col] = numeric
            # else: leave blank — N/A or non-numeric means field not applicable
        elif value.lower() not in ("n/a", "-", "none", ""):
            specs[col] = value

    # Capture PDF spec sheet as Tearsheet Link
    if spec_table:
        for a in spec_table.find_all("a", href=True):
            href = a["href"]
            if href.endswith(".pdf") or "SPEC SHEET" in a.get_text().upper():
                specs.setdefault("Tearsheet Link", href)
                break

    return description, specs


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
    handle    = product.get("handle", "")
    source_base = f"{BASE_URL}/products/{handle}"
    vendor    = product.get("vendor", "")
    tags      = product.get("tags", [])

    image_map   = {img["id"]: img["src"] for img in product.get("images", [])}
    all_images  = [img["src"] for img in product.get("images", [])]
    default_img = all_images[0] if all_images else ""

    # Parse description text and structured spec fields separately
    description, specs = parse_body_html(body_html)

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
            "Description":       description,
        }
        # Apply all parsed spec fields (Arm Height, Seat Height, Height, Width, etc.)
        row.update(specs)

        row["SKU"]   = variant.get("sku") or specs.get("SKU") or ""
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
        if opt1 and not opt1.startswith("Default Title"):
            key = options[0]["name"] if options else "Finish"
            if key.lower() != "title":
                row[key] = opt1
        if opt2 and not opt2.startswith("Default Title"):
            key = options[1]["name"] if len(options) > 1 else "Size"
            if key.lower() != "title":
                row[key] = opt2
        if opt3 and not opt3.startswith("Default Title"):
            key = options[2]["name"] if len(options) > 2 else "Option 3"
            if key.lower() != "title":
                row[key] = opt3

        if vendor:
            row["Designer"] = vendor
        if tags:
            row["Collection"] = ", ".join(tags)

        rows.append(row)

    return rows if rows else [{
        "Source":            source_base,
        "Product Name":      title,
        "Product Family Id": extract_family_id(title),
        "Image URL":         default_img,
        "Description":       description,
        **specs,
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
