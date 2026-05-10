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
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    sentence_case,
    clean_price,
    generate_sku,
    extract_family_id,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Woud")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get(
    "OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://wouddesign.com"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

CM_TO_IN = 0.393701  # 1 cm in inches


def cm_to_in(cm_str: str) -> str:
    """Convert a cm numeric string to inches, rounded to 2 decimal places."""
    try:
        return str(round(float(cm_str) * CM_TO_IN, 2))
    except (ValueError, TypeError):
        return cm_str


# Mapping from HTML label → output field name
_DETAIL_MAP = {
    "colour":            "Color",
    "color":             "Color",
    "materials":         "Materials",
    "material":          "Material",
    "cord":              "Cord",
    "socket":            "Socket",
    "lightsource":       "Lamp Type",
    "light source":      "Lamp Type",
    "bulb":              "Bulb Type",
    "specifications":    "Specifications",
    "country of origin": "Origin",
    "finishing":         "Finish",
    "finish":            "Finish",
    "legs":              "Legs",
    "frame":             "Frame",
    "shade":             "Shade Details",
    "canopy":            "Canopy",
    "construction":      "Construction",
    "pile height":       "Pile Height",
    "thickness":         "Thickness",
    "shown here":        "Upholstery",
    "composition":       "Composition",
    "care instructions": "Care Instructions",
}

# Measurement labels → field names  (all values extracted in cm then converted to inches)
_MEASUREMENT_MAP = [
    ("length",      "Length"),
    ("width",       "Width"),
    ("depth",       "Depth"),
    ("height",      "Height"),
    ("diameter",    "Diameter"),
    ("seat height", "Seat Height"),
    ("seat depth",  "Seat Depth"),
    ("seat width",  "Seat Width"),
    ("arm height",  "Arm Height"),
    ("cord",        "Cord Length"),   # numeric cm value in Measurements section
]


# ---------------------------------------------------------------------------
# Shopify collection API — fast path for /collections/ URLs
# ---------------------------------------------------------------------------
def get_collection_products(collection_handle: str) -> list[dict]:
    """Return all products from a Shopify collection (250 per page, auto-paginated)."""
    products: list[dict] = []
    page = 1
    while True:
        url = f"{BASE_URL}/collections/{collection_handle}/products.json"
        try:
            r = requests.get(url, params={"limit": 250, "page": page},
                             headers=HEADERS, timeout=25)
            if r.status_code != 200:
                break
            batch = r.json().get("products", [])
            if not batch:
                break
            products.extend(batch)
            if len(batch) < 250:
                break
            page += 1
            time.sleep(0.5)
        except Exception as e:
            print(f"    [API] collection error page {page}: {e}")
            break
    return products


def get_product_by_handle(handle: str) -> dict | None:
    """Fetch one product via Shopify product JSON API."""
    for url in [
        f"{BASE_URL}/products/{handle}.js",
        f"{BASE_URL}/en-us/products/{handle}.js",
    ]:
        try:
            r = requests.get(url, headers=HEADERS, timeout=25)
            if r.status_code == 200:
                return r.json()
        except Exception:
            pass
    return None


# ---------------------------------------------------------------------------
# Product HTML parsing — details, measurements, designer, files
# ---------------------------------------------------------------------------
def parse_product_html(handle: str) -> dict:
    """
    Fetch product detail page HTML and extract details, measurements, designer, files.
    Tries /en-us/products/ first; falls back to /products/ if 404.
    Stores the working URL in extra["_source_url"] for process_product to use.
    """
    extra: dict = {}
    soup = None
    working_url = f"{BASE_URL}/en-us/products/{handle}"

    for try_url in [
        f"{BASE_URL}/en-us/products/{handle}",
        f"{BASE_URL}/products/{handle}",
    ]:
        try:
            r = requests.get(try_url, headers=HEADERS, timeout=25)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "html.parser")
                working_url = try_url
                break
        except Exception as e:
            print(f"    [HTML] error ({try_url}): {e}")

    extra["_source_url"] = working_url   # consumed by process_product
    if soup is None:
        return extra

    page_text = soup.get_text(separator="\n")


    # --- Key: value pairs from Details accordion ---
    for label, field in _DETAIL_MAP.items():
        if field in extra:
            continue
        m = re.search(
            rf'(?:^|\n)\s*{re.escape(label)}\s*:\s*([^\n]+)',
            page_text, re.IGNORECASE,
        )
        if m:
            val = clean_text(m.group(1))
            if val:
                extra[field] = val

    # --- Measurements from Measurements accordion (site shows cm → convert to inches) ---
    for label, field in _MEASUREMENT_MAP:
        m = re.search(
            rf'(?:^|\n)\s*{re.escape(label)}\s*:\s*([\d.]+)\s*(?:cm|mm)?',
            page_text, re.IGNORECASE,
        )
        if m:
            extra[field] = cm_to_in(m.group(1))

    # Shade: Diameter: X cm  (special nested format found in lighting products)
    shade_dia = re.search(
        r'shade\s*:\s*diameter\s*:\s*([\d.]+)\s*(?:cm|mm)?',
        page_text, re.IGNORECASE,
    )
    if shade_dia:
        extra["Shade Diameter"] = cm_to_in(shade_dia.group(1))

    # Build Dimensions string from primary dimension sub-fields (now in inches)
    dim_parts = []
    for abbr, field in [("W", "Width"), ("D", "Depth"), ("H", "Height"),
                         ("L", "Length"), ("Dia", "Diameter")]:
        if extra.get(field):
            dim_parts.append(f"{abbr} {extra[field]}")
    if dim_parts:
        extra["Dimensions"] = " x ".join(dim_parts)

    # --- Designer: "Designed by [Name]" ---
    dm = re.search(r'Designed by\s+([A-Z][^\n]+)', page_text)
    if dm:
        designer = clean_text(dm.group(1))
        designer = re.split(r'\s*Read more', designer, flags=re.IGNORECASE)[0].strip()
        if designer:
            extra["Designer"] = designer

    # --- File links ---
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not href:
            continue
        label = clean_text(a.get_text()).lower()
        if "assembly" in label and "Assembly Instructions" not in extra:
            extra["Assembly Instructions"] = href
        elif ("care" in label or "maintenance" in label) and "Care Instructions" not in extra:
            extra["Care Instructions"] = href
        elif ("2d" in label or "3d" in label) and "2D/3D Files" not in extra:
            extra["2D/3D Files"] = href

    return extra


# ---------------------------------------------------------------------------
# Convert Shopify product dict → row dicts (one per variant)
# ---------------------------------------------------------------------------
def _img_src(img_item) -> str:
    """Return https CDN URL from either a dict-format or string-format image entry."""
    src = img_item.get("src", "") if isinstance(img_item, dict) else str(img_item)
    return ("https:" + src) if src.startswith("//") else src


def process_product(api_product: dict, html_extra: dict) -> list[dict]:
    """
    One Shopify product may have multiple variants (e.g. different sizes).
    Returns one row dict per variant, sharing all non-variant fields.
    On Woud, each color/finish is a separate product URL, so most products
    have just one or a few size variants.
    """
    title      = api_product.get("title", "")
    body_html  = api_product.get("body_html", "")
    handle     = api_product.get("handle", "")

    # Use the URL verified by parse_product_html (may be non-locale if en-us was 404)
    source_url = html_extra.pop("_source_url", f"{BASE_URL}/en-us/products/{handle}")

    description = clean_text(
        BeautifulSoup(body_html, "html.parser").get_text(" ")
    ) if body_html else ""

    # Handle two image formats: list-of-dicts (collection API) vs list-of-strings (product.js API)
    images_raw  = api_product.get("images", [])
    default_img = _img_src(images_raw[0]) if images_raw else ""
    image_map: dict = {
        img["id"]: _img_src(img)
        for img in images_raw
        if isinstance(img, dict) and img.get("id")
    }

    # Normalise options: collection API returns list of dicts; product.js returns list of strings
    options_raw  = api_product.get("options", [])
    option_names = [
        (opt.get("name") if isinstance(opt, dict) else str(opt))
        for opt in options_raw
    ]
    variants = api_product.get("variants", [])

    base_row = {
        "Source URL":       source_url,
        "Product Name":     sentence_case(title),
        "Product Family Id": extract_family_id(title),
        "Description":      description,
        "Manufacturer":     VENDOR_NAME,
    }
    base_row.update(html_extra)

    rows: list[dict] = []
    for variant in variants:
        if not isinstance(variant, dict):
            continue

        row = dict(base_row)

        # SKU
        row["SKU"] = variant.get("sku") or ""

        # Price: Shopify returns cents as int (49900) or str ("49900")
        try:
            row["Price"] = round(int(variant.get("price", 0)) / 100, 2)
        except (TypeError, ValueError):
            row["Price"] = clean_price(str(variant.get("price", "")))

        # Weight: API grams → kg (only if HTML didn't provide it)
        grams = variant.get("grams")
        if grams and not row.get("Weight"):
            row["Weight"] = round(grams / 1000, 2)

        # Image: prefer variant-specific, then product default
        vi = variant.get("featured_image")
        if vi and isinstance(vi, dict) and vi.get("src"):
            row["Image URL"] = _img_src(vi)
        elif variant.get("image_id") and variant["image_id"] in image_map:
            row["Image URL"] = image_map[variant["image_id"]]
        else:
            row["Image URL"] = default_img

        # Shopify variant option values (Finish, Size, Color, etc.)
        for i, opt_val in enumerate([
            variant.get("option1"),
            variant.get("option2"),
            variant.get("option3"),
        ]):
            if opt_val and opt_val not in ("Default Title", ""):
                opt_name = option_names[i] if i < len(option_names) else f"Option {i + 1}"
                row[opt_name] = opt_val

        rows.append(row)

    return rows or [{
        "Source URL":        source_url,
        "Product Name":      sentence_case(title),
        "Product Family Id": extract_family_id(title),
        "Image URL":         default_img,
        "Description":       description,
        "Manufacturer":      VENDOR_NAME,
        **html_extra,
    }]


# ---------------------------------------------------------------------------
# Playwright helper — extract product handles from /pages/ landing pages
# ---------------------------------------------------------------------------
async def get_handles_from_page(browser_page, page_url: str) -> list[str]:
    """
    Woud uses /pages/ landing pages for some categories (e.g. lounge-chairs,
    nakki-series-overview). These need JS rendering to show product cards.
    Returns deduplicated product handles in order of appearance.
    """
    handles: list[str] = []
    seen: set[str] = set()
    try:
        await browser_page.goto(page_url, timeout=45_000, wait_until="domcontentloaded")
        await browser_page.wait_for_timeout(3000)

        anchors = await browser_page.query_selector_all('a[href*="/products/"]')
        for a in anchors:
            href = await a.get_attribute("href")
            if not href or "/products/" not in href:
                continue
            handle = href.split("/products/")[-1].split("?")[0].rstrip("/")
            if handle and handle not in seen:
                seen.add(handle)
                handles.append(handle)
    except Exception as e:
        print(f"    [Playwright] page error ({page_url}): {e}")
    return handles


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST MODE: max {TEST_MAX_CATEGORIES} cats, {TEST_MAX_PRODUCTS} products each]")

    async with PlaywrightBrowser(headless=HEADLESS) as browser_page:
        for cat in categories:
            if not cat["links"]:
                continue

            print(f"\n{'=' * 60}")
            print(f"Category: {cat['name']}")
            print(f"{'=' * 60}")

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            # Collect all Shopify product dicts, deduplicated by handle
            seen_handles: set[str] = set()
            all_products: list[dict] = []

            for listing_url in cat["links"]:
                if "/en-us/collections/" in listing_url:
                    # Fast path: Shopify collection API
                    coll_handle = listing_url.rstrip("/").split("/")[-1]
                    api_batch = get_collection_products(coll_handle)
                    print(f"  [API] {listing_url} -> {len(api_batch)} products")
                    for p in api_batch:
                        h = p.get("handle", "")
                        if h and h not in seen_handles:
                            seen_handles.add(h)
                            all_products.append(p)
                else:
                    # /pages/ URL: extract handles via Playwright, then API per handle
                    page_handles = await get_handles_from_page(browser_page, listing_url)
                    print(f"  [Page] {listing_url} -> {len(page_handles)} links")
                    for h in page_handles:
                        if h and h not in seen_handles:
                            seen_handles.add(h)
                            p = get_product_by_handle(h)
                            if p:
                                all_products.append(p)
                            time.sleep(0.3)

            print(f"  Total unique products: {len(all_products)}")
            if TEST_MODE:
                all_products = all_products[:TEST_MAX_PRODUCTS]

            global_idx = 1
            for api_product in all_products:
                handle = api_product.get("handle", "")
                print(f"  [{global_idx:>3}] {handle}")
                try:
                    html_extra = parse_product_html(handle)
                    rows = process_product(api_product, html_extra)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(
                                info["vendor_name"], cat["name"], global_idx
                            )
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(
                                row["Product Name"]
                            )
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"    ERROR on {handle}: {e}")

                time.sleep(0.8)

    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
