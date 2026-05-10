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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Hennepin Made")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://hennepinmade.com"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

# Mapping from Technical Specs label → output field name
_SPEC_LABEL_MAP = {
    "DIMENSIONS": "Dimensions",
    "DIMENSIONS (LIGHTS)": "Dimensions",
    "DIMENSIONS (CANOPY)": "Canopy Dimensions",
    "LAMPING": "Lamping",
    "LUMENS": "Lumens",
    "COLOR TEMP": "Color Temperature",
    "COLOR TEMPERATURE": "Color Temperature",
    "INPUT VOLTAGE": "Voltage",
    "MOUNTING": "Mounting",
    "OVERALL HEIGHT": "Overall Height",
    "WEIGHT": "Weight",
    "WEIGHT (CANOPY AND LIGHTS)": "Weight",
    "MOUNT": "Mount",
    "CORD LENGTH": "Cord Length",
    "DROP ROD": "Drop Rod",
    "SOCKET": "Socket",
    "SOCKET TYPE": "Socket Type",
    "WATTAGE": "Wattage",
    "CRI": "CRI",
    "DIMMING": "Dimming",
    "LEAD TIME": "Lead Time",
}


def clean_html(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return clean_text(soup.get_text(separator=" "))


def get_collection_products(collection_handle: str) -> list[dict]:
    products = []
    page = 1
    while True:
        url = f"{BASE_URL}/collections/{collection_handle}/products.json"
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


def scrape_product_page(handle: str) -> dict:
    """
    Fetch the HTML product page and extract:
    - Technical Specs (all label→value pairs from the Collapsible__Content Rte block)
    - Tearsheet Link (first PDF from Downloads collapsible)
    - Lead Time (from product-meta text on page)
    Returns a flat dict of additional fields to merge into variant rows.
    """
    url = f"{BASE_URL}/products/{handle}"
    extra: dict = {}
    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        if r.status_code != 200:
            return extra
        soup = BeautifulSoup(r.text, "html.parser")
    except Exception as e:
        print(f"    HTML fetch error ({handle}): {e}")
        return extra

    # --- Technical Specs collapsible ---
    for btn in soup.find_all("button", class_="Collapsible__Button"):
        if "Technical Specs" not in btn.get_text():
            continue
        content = btn.find_parent(class_="Collapsible").find(class_="Collapsible__Content")
        if not content:
            break
        rte = content.find(class_="Rte") or content

        # Parse each <p>: <strong>LABEL</strong><br/>val1<br/>val2...
        for p in rte.find_all("p"):
            strong = p.find("strong")
            if not strong:
                continue
            label = strong.get_text(strip=True).upper()

            # Collect all text after the <strong>, joined across <br/> and inline elements
            value_parts = []
            for child in p.children:
                if child == strong:
                    continue
                if child.name == "br":
                    continue
                if hasattr(child, "get_text"):
                    t = child.get_text(strip=True)
                    if t:
                        value_parts.append(t)
                else:
                    t = str(child).strip()
                    if t and t != ":":
                        value_parts.append(t)
            value = " ".join(value_parts).strip(" :")

            if not value:
                continue

            field = _SPEC_LABEL_MAP.get(label)
            if field:
                # Don't overwrite Dimensions if we already have a better value
                if field not in extra or not extra[field]:
                    extra[field] = value
            else:
                # Store unknown labels using Title Case key
                key = label.title()
                extra[key] = value

        break

    # --- Parse Weight: strip " lbs" → numeric ---
    if extra.get("Weight"):
        w_match = re.search(r"(\d+(?:\.\d+)?)", extra["Weight"])
        if w_match:
            extra["Weight"] = w_match.group(1)

    # --- Parse Lumens: strip " lm" → numeric string ---
    if extra.get("Lumens"):
        lm_match = re.search(r"(\d+(?:[,]\d+)?)", extra["Lumens"].replace(",", ""))
        if lm_match:
            extra["Lumens"] = lm_match.group(1)

    # --- Downloads collapsible → Tearsheet Link (first PDF) ---
    for btn in soup.find_all("button", class_="Collapsible__Button"):
        if "Downloads" not in btn.get_text():
            continue
        content = btn.find_parent(class_="Collapsible").find(class_="Collapsible__Content")
        if not content:
            break
        for a in content.find_all("a", href=True):
            href = a["href"]
            # Prefer the Product Spec Sheet / Tear Sheet PDF
            label = a.get_text(strip=True).lower()
            if "spec" in label or "tear" in label:
                extra["Tearsheet Link"] = href
                break
        # Fallback: first PDF link
        if not extra.get("Tearsheet Link"):
            for a in content.find_all("a", href=True):
                if ".pdf" in a["href"].lower():
                    extra["Tearsheet Link"] = a["href"]
                    break
        break

    # --- Lead Time ---
    for el in soup.find_all(string=re.compile(r"lead time", re.I)):
        if el.parent.name in ("script", "style"):
            continue
        t = clean_text(str(el))
        if t:
            extra["Lead Time"] = t
            break

    return extra


def process_product(product: dict, page_extra: dict) -> list[dict]:
    title = product.get("title", "")
    body_html = product.get("body_html", "")
    description = clean_html(body_html)
    handle = product.get("handle", "")
    source_url = f"{BASE_URL}/products/{handle}"
    tags = product.get("tags", [])
    product_type = product.get("product_type", "")

    # Product images from API
    product_images = [img["src"] for img in product.get("images", [])]
    image_map = {img["id"]: img["src"] for img in product.get("images", [])}
    default_image = product_images[0] if product_images else ""

    # Build Collection from tags (filter out generic type tags used as collection names)
    collection_val = product_type or ""

    options = product.get("options", [])
    variants = product.get("variants", [])
    rows = []

    for variant in variants:
        variant_id = variant.get("id")
        variant_url = f"{source_url}?variant={variant_id}" if variant_id else source_url
        row = {
            "Source": variant_url,
            "Product Name": title,
            "Product Family Id": extract_family_id(title),
            "Description": description,
        }

        # Merge Technical Specs + Tearsheet from HTML page
        row.update(page_extra)

        # SKU and Price from API
        row["SKU"] = variant.get("sku") or ""
        row["Price"] = clean_price(str(variant.get("price", "")))

        # Image URL: prefer variant-specific image
        vi = variant.get("featured_image")
        if vi and vi.get("src"):
            row["Image URL"] = vi["src"]
        elif variant.get("image_id") and variant["image_id"] in image_map:
            row["Image URL"] = image_map[variant["image_id"]]
        else:
            row["Image URL"] = default_image

        # Weight from API (grams → lbs) — only if not already set from Tech Specs
        if not row.get("Weight"):
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
            key = options[1]["name"] if len(options) > 1 else "Option 2"
            row[key] = opt2
        if opt3 and opt3 not in ("Default Title", ""):
            key = options[2]["name"] if len(options) > 2 else "Option 3"
            row[key] = opt3

        if collection_val:
            row["Collection"] = collection_val

        rows.append(row)

    return rows if rows else [{
        "Source": source_url,
        "Product Name": title,
        "Product Family Id": extract_family_id(title),
        "Image URL": default_image,
        "Description": description,
        **page_extra,
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

        # Collect products from all links, deduplicating by product handle
        seen_handles: set[str] = set()
        all_products: list[dict] = []
        for link_url in cat["links"]:
            handle = link_url.rstrip("/").split("/")[-1]
            for product in get_collection_products(handle):
                ph = product.get("handle", "")
                if ph not in seen_handles:
                    seen_handles.add(ph)
                    all_products.append(product)

        products = all_products
        if TEST_MODE:
            products = products[:TEST_MAX_PRODUCTS]

        print(f"  {cat['name']}: {len(products)} products (links={len(cat['links'])})")

        global_idx = 1
        for product in products:
            try:
                handle = product.get("handle", "")
                # Fetch HTML page for Technical Specs + Tearsheet
                page_extra = scrape_product_page(handle)
                rows = process_product(product, page_extra)
                for row in rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
            except Exception as e:
                print(f"    ERROR on {product.get('handle', '?')}: {e}")
            time.sleep(0.8)

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
