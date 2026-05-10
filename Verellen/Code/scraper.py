"""
Verellen scraper — uses the Magento GraphQL API at magento.verellen.biz/graphql.
No Playwright required (API is public and returns full product data).
"""
import asyncio
import base64
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
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Verellen")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

GRAPHQL_URL = "https://magento.verellen.biz/graphql"
SITE_BASE = "https://verellen.biz"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
}

# Maps tracker URL-key segment → Magento category UID (base64 of numeric ID)
# Built by fetching category IDs via GraphQL beforehand
CATEGORY_URL_KEY_TO_UID: dict[str, str] = {}


def _b64uid(numeric_id: int) -> str:
    return base64.b64encode(str(numeric_id).encode()).decode()


def gql(query: str, variables: dict | None = None) -> dict:
    """Execute a GraphQL query against the Verellen backend."""
    payload: dict = {"query": query}
    if variables:
        payload["variables"] = variables
    try:
        r = requests.post(GRAPHQL_URL, json=payload, headers=HEADERS, timeout=30)
        return r.json()
    except Exception as e:
        print(f"    GraphQL error: {e}")
        return {}


def resolve_category_uid(url_key: str) -> str | None:
    """Return base64 UID for a Magento category given its url_key."""
    if url_key in CATEGORY_URL_KEY_TO_UID:
        return CATEGORY_URL_KEY_TO_UID[url_key]

    data = gql(f'{{ categories(filters: {{url_key: {{eq: "{url_key}"}}}}) {{ items {{ id url_key }} }} }}')
    items = data.get("data", {}).get("categories", {}).get("items", [])
    if items:
        uid = _b64uid(items[0]["id"])
        CATEGORY_URL_KEY_TO_UID[url_key] = uid
        return uid
    return None


def clean_html(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return clean_text(soup.get_text(separator=" "))


PRODUCT_FIELDS = """
  name
  sku
  description { html }
  short_description { html }
  price_range { minimum_price { regular_price { value } } }
  image { url }
  media_gallery { url label }
  url_key
  categories { name url_key }
  ... on SimpleProduct { weight }
"""


def get_products_by_category_uid(uid: str) -> list[dict]:
    """Paginate through all products in a Magento category."""
    all_products = []
    page = 1

    while True:
        data = gql(f"""{{
          products(filter: {{category_uid: {{eq: "{uid}"}}}}, pageSize: 50, currentPage: {page}) {{
            total_count
            items {{ {PRODUCT_FIELDS} }}
          }}
        }}""")
        items = data.get("data", {}).get("products", {}).get("items", [])
        if not items:
            break
        all_products.extend(items)
        total = data["data"]["products"].get("total_count", 0)
        if len(all_products) >= total:
            break
        page += 1
        time.sleep(0.5)

    return all_products


def process_product(product: dict) -> dict:
    """Convert a Verellen GraphQL product to a flat row dict."""
    name = product.get("name", "")
    description = clean_html(product.get("description", {}).get("html", ""))
    url_key = product.get("url_key", "")
    source_url = f"{SITE_BASE}/products/{url_key}" if url_key else SITE_BASE

    row = {
        "Source": source_url,
        "Product Name": name,
        "Product Family Id": extract_family_id(name),
        "SKU": (product.get("sku") or "").strip("()").strip(),
        "Description": description,
    }

    # Price (minimum price = list price)
    price_val = (
        product.get("price_range", {})
        .get("minimum_price", {})
        .get("regular_price", {})
        .get("value")
    )
    if price_val is not None:
        row["Price"] = float(price_val)

    # Image URL — use first media_gallery image for best quality
    gallery = product.get("media_gallery", [])
    main_img = product.get("image", {}).get("url", "")
    if gallery:
        row["Image URL"] = gallery[0]["url"]
    elif main_img:
        row["Image URL"] = main_img

    # Weight
    weight = product.get("weight")
    if weight:
        row["Weight"] = safe_float(str(weight))

    # Dimensions — parse from description (only keep numeric fields)
    if description:
        dims = parse_dimensions(description)
        dims.pop("Dimensions", None)
        row.update(dims)

    return row


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

        # Collect products from all listing URLs for this category
        seen_skus: set[str] = set()
        all_products: list[dict] = []

        for listing_url in cat["links"]:
            # Extract URL key from listing URL (last non-empty path segment)
            url_key = listing_url.rstrip("/").split("/")[-1]
            uid = resolve_category_uid(url_key)
            if not uid:
                print(f"    WARNING: Category '{url_key}' not found in Magento GraphQL")
                continue

            products = get_products_by_category_uid(uid)
            for p in products:
                sku = p.get("sku", "")
                if sku not in seen_skus:
                    seen_skus.add(sku)
                    all_products.append(p)

        if TEST_MODE:
            all_products = all_products[:TEST_MAX_PRODUCTS]

        print(f"  {cat['name']}: {len(all_products)} products")

        global_idx = 1
        for product in all_products:
            try:
                row = process_product(product)
                if not row.get("SKU"):
                    row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                if not row.get("Product Family Id") and row.get("Product Name"):
                    row["Product Family Id"] = extract_family_id(row["Product Name"])
                writer.write_row(row, category_name=cat["name"])
                global_idx += 1
            except Exception as e:
                print(f"    ERROR on {product.get('name', '?')}: {e}")

        time.sleep(0.5)

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
