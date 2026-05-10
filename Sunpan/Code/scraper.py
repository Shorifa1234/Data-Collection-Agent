"""
scraper.py — Sunpan
--------------------
Platform: sunpan.com (Shopify)

Listing  : /collections/{slug}/products.json?limit=250&page=N  (Shopify JSON API)
Product  : /products/{handle}  (requests + BeautifulSoup for spec table + dimensions)

Product fields:
  Product Name   : product title (per variant title if distinct)
  SKU            : variant sku
  Price          : variant price
  Finish / Color : variant option1
  Image URL      : variant featured_image.src  (or first product image)
  Description    : product body_html (stripped)
  Material       : spec table row "Material"
  Shape          : spec table row "Shape"
  Base           : spec table row "Base / Legs"
  Finish         : spec table row "Base Finish" + "Material Finish"
  Designer       : spec table row "Designed by"
  Dimensions     : Overall Dimensions row → parsed into Width/Height/Depth
  Weight         : Carton Weight / Net Weight from spec table
  + any other spec table rows found

Run directly:
    python scraper.py
Or via orchestrator:
    python orchestrator.py "Sunpan"
    python orchestrator.py "Sunpan" --test
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text,
    clean_price,
    sentence_case,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Sunpan")
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://sunpan.com"
SESSION  = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
})

# Spec table label → output field name
_SPEC_LABEL_MAP = {
    "sku no":             "SKU",
    "material":           "Materials",
    "shape":              "Shape",
    "base / legs":        "Base",
    "base/legs":          "Base",
    "base finish":        "Finish",
    "material finish":    "Color",
    "designed by":        "Designer",
    "contract viable":    "Contract Viable",
    "assembly required":  "Assembly Required",
    "warranty":           "Warranty",
    "carb compliant":     "CARB Compliant",
    "additional features": "Additional Features",
    "weight":             "Weight",
    "net weight":         "Net Weight",
    "collection":         "Collection",
    "style":              "Style",
    "color":              "Color",
    "fabric":             "Fabric",
    "seat height":        "Seat Height",
    "seat depth":         "Seat Depth",
    "seat width":         "Seat Width",
    "arm height":         "Arm Height",
    "socket":             "Socket",
    "wattage":            "Wattage",
}


# ---------------------------------------------------------------------------
# Listing — Shopify products.json API
# ---------------------------------------------------------------------------

def get_collection_slug(listing_url: str) -> str:
    """Extract collection slug from a Sunpan collection URL."""
    m = re.search(r'/collections/([^/?#]+)', listing_url)
    return m.group(1) if m else ""


def get_product_handles(listing_url: str, max_products: int = 0) -> list[dict]:
    """
    Return all Shopify products for a collection via the JSON API.
    Each item is the full Shopify product dict (with variants + images).
    """
    slug = get_collection_slug(listing_url)
    if not slug:
        return []

    # Build query params from the listing URL (e.g. filter params)
    # For filtered listings like ?filter.p.product_type=Ottomans we pass them too
    extra_params = ""
    if "?" in listing_url:
        qs = listing_url.split("?", 1)[1]
        extra_params = "&" + qs if qs else ""

    products = []
    page = 1
    while True:
        url = f"{BASE_URL}/collections/{slug}/products.json?limit=250&page={page}{extra_params}"
        try:
            resp = SESSION.get(url, timeout=20)
            resp.raise_for_status()
            data = resp.json().get("products", [])
        except Exception as e:
            print(f"  [WARN] API error {url}: {e}")
            break
        if not data:
            break
        products.extend(data)
        if max_products and len(products) >= max_products:
            products = products[:max_products]
            break
        page += 1
        time.sleep(0.5)

    return products


# ---------------------------------------------------------------------------
# Product detail — requests + BeautifulSoup for spec table
# ---------------------------------------------------------------------------

def _strip_html(html: str) -> str:
    return clean_text(re.sub(r"<[^>]+>", " ", html or ""))


def _upgrade_cdn_image(url: str) -> str:
    """Remove Shopify CDN resize params and request large version."""
    if not url:
        return ""
    url = re.sub(r"_\d+x\d+(\.\w+)(\?|$)", r"\1\2", url)
    url = url.split("?")[0]
    return url


def _scrape_product_page(handle: str) -> dict:
    """
    Fetch the product HTML page and extract the spec table.
    Returns a dict of extra fields (Material, Dimensions, Designer, etc.).
    """
    url = f"{BASE_URL}/products/{handle}"
    try:
        resp = SESSION.get(url, timeout=20)
        resp.raise_for_status()
    except Exception as e:
        print(f"    [WARN] detail page {handle}: {e}")
        return {}

    soup = BeautifulSoup(resp.text, "html.parser")
    page_text = soup.get_text(separator="\n")
    extras: dict = {}

    # Parse all spec table rows (2-column key→value rows)
    for tbl in soup.find_all("table"):
        rows = tbl.find_all("tr")
        in_dims_section = False
        for row in rows:
            cells = [c.get_text(strip=True) for c in row.find_all(["td", "th"])]
            if not cells:
                continue

            if len(cells) == 1:
                label = cells[0].lower().strip()
                in_dims_section = (label == "dimensions")
                continue

            label = cells[0].lower().strip()
            # Prefer Imperial value (cells[1]); cells[2] is Metric if present
            value = clean_text(cells[1]) if len(cells) > 1 and cells[1] else ""

            if in_dims_section:
                if label in ("metric", "imperial", ""):
                    continue
                if label in ("carton weight", "net weight", "weight"):
                    in_dims_section = False
                    wt = re.sub(r"\s*(kg|cm|mm)\b.*", "", value, flags=re.I).strip()
                    extras.setdefault("Weight", wt)
                continue  # skip other dim sub-rows — handled via text regex below

            field = _SPEC_LABEL_MAP.get(label)
            if field and value:
                extras.setdefault(field, value)

    # Dimensions — more reliable via text search than table cell parsing
    # Sunpan format: "Overall Dimensions\n30.50W x 18.00D x 26.00H in\n77.47W..."
    dim_match = re.search(
        r"Overall\s+Dimensions\s*\n\s*([\d.]+W\s*x\s*[\d.]+D?\s*x?\s*[\d.]*H?[^\n]*in)",
        page_text, re.I
    )
    if not dim_match:
        # Broader fallback — look for "NNW x NND x NNH" pattern in the spec area
        dim_match = re.search(
            r"([\d.]+W\s*x\s*[\d.]+D\s*x\s*[\d.]+H)",
            page_text
        )
    if dim_match:
        raw = dim_match.group(1)
        dim_str = re.sub(r"\s*(in|cm)\s*$", "", raw, flags=re.I).strip()
        extras.setdefault("Dimensions", dim_str)
        parsed = parse_dimensions(dim_str)
        for k, v in parsed.items():
            if k != "Dimensions":
                extras.setdefault(k, v)

    return extras


# ---------------------------------------------------------------------------
# Build rows from Shopify product dict + HTML extras
# ---------------------------------------------------------------------------

def _build_rows(shopify_prod: dict, html_extras: dict) -> list[dict]:
    """Return one row dict per variant."""
    title     = sentence_case(shopify_prod.get("title", ""))
    body_html = shopify_prod.get("body_html", "")
    description = _strip_html(body_html)
    handle    = shopify_prod.get("handle", "")
    product_url = f"{BASE_URL}/products/{handle}"
    images    = shopify_prod.get("images", [])

    rows = []
    for variant in shopify_prod.get("variants", []):
        sku      = variant.get("sku", "")
        price    = clean_price(str(variant.get("price") or ""))
        option1  = variant.get("option1") or ""
        option2  = variant.get("option2") or ""

        # Variant display name
        v_title = variant.get("title", "Default Title")
        if v_title and v_title != "Default Title":
            name = f"{title} - {v_title}"
        else:
            name = title

        # Image: prefer variant image → first product image
        v_img_data = variant.get("featured_image")
        if v_img_data and v_img_data.get("src"):
            img = _upgrade_cdn_image(v_img_data["src"])
        elif images:
            img = _upgrade_cdn_image(images[0]["src"])
        else:
            img = ""

        row: dict = {
            "Source":            f"{product_url}?variant={variant.get('id', '')}",
            "Product Name":      name,
            "Product Family Id": title,
            "SKU":               sku,
            "Price":             price,
            "Description":       description,
            "Image URL":         img,
            "Manufacturer":      VENDOR_NAME,
        }

        # Finish from variant option
        if option1 and option1 != "Default Title":
            row["Finish"] = option1
        if option2 and option2 != "Default Title":
            row["Color"] = option2

        # Merge html_extras (Material, Dimensions, etc.) but don't overwrite variant fields
        for k, v in html_extras.items():
            if k not in row:
                row[k] = v

        rows.append(row)

    return rows if rows else [{"Source": product_url, "Product Name": title,
                               "Product Family Id": title, "Manufacturer": VENDOR_NAME,
                               **html_extras}]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
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

        writer.add_sheet(
            cat["name"],
            cat["links"][0],
            studio_columns=cat["studio_columns"],
        )

        # Collect all Shopify products from all listing URLs for this category
        seen_handles: set[str] = set()
        all_products: list[dict] = []
        max_p = TEST_MAX_PRODUCTS if TEST_MODE else 0

        for listing_url in cat["links"]:
            for prod in get_product_handles(listing_url, max_products=max_p):
                handle = prod.get("handle", "")
                if handle and handle not in seen_handles:
                    seen_handles.add(handle)
                    all_products.append(prod)

        if TEST_MODE:
            all_products = all_products[:TEST_MAX_PRODUCTS]

        print(f"\n[Category] {cat['name']}: {len(all_products)} products")

        global_idx = 1
        for shopify_prod in all_products:
            handle = shopify_prod.get("handle", "")
            try:
                html_extras = _scrape_product_page(handle)
                rows = _build_rows(shopify_prod, html_extras)
                for row in rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
                print(f"  [{global_idx - 1}] {handle}")
            except Exception as e:
                print(f"  [ERROR] {handle}: {e}")
            time.sleep(0.6)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
