import asyncio
import json
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

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

VENDOR_NAME = os.environ.get("VENDOR_NAME", "SkLO")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.sklo.com"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


# Unicode prime characters used in SkLO dimension strings
_PRIME_RE = re.compile(r'[\u2033\u2032\u201d\u2019"\']')


def _parse_sklo_dimensions(dims_str: str) -> dict:
    """
    Parse SkLO dimension strings like:
      "30″H x 18.5″Dia (760x470mm)"
      "8″H x 7.5″Dia (200x190mm)"
      "9″H (230mm)"
      "12″W x 6″D x 4″H"
    Returns a dict with Dimensions, and any of Height/Width/Depth/Diameter found.
    """
    if not dims_str:
        return {}

    # Strip Unicode prime / inch marks and the metric part in parentheses
    cleaned = _PRIME_RE.sub("", dims_str)
    cleaned = re.sub(r"\s*\([^)]+\)", "", cleaned).strip()

    result = {}
    dim_map = {"H": "Height", "W": "Width", "D": "Depth", "Dia": "Diameter", "L": "Length"}
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*(Dia|H|W|D|L)\b", cleaned, re.IGNORECASE):
        num, key = m.group(1), m.group(2).capitalize()
        if key == "Dia":
            result["Diameter"] = num
        else:
            full = dim_map.get(key.upper(), key)
            result[full] = num

    if cleaned:
        result["Dimensions"] = cleaned

    return result


def _extract_all_variants(soup: BeautifulSoup) -> list[dict] | None:
    """
    Extract the allVariants JS array from the page's inline <script>.
    Returns a list of variant dicts, or None if not found.
    """
    for script in soup.find_all("script"):
        text = script.string or ""
        if "allVariants" not in text:
            continue
        m = re.search(r"var allVariants\s*=\s*(\[.*?\]);", text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass
    return None


def get_product_links(listing_url: str) -> list[str]:
    """
    Collect all product URLs from a SkLO listing page.
    Filters to only links whose path matches /category/product-slug/.
    """
    parsed = urlparse(listing_url)
    cat_path = "/" + parsed.path.strip("/").split("/")[0] + "/"  # e.g. "/light/"

    links: list[str] = []
    seen: set[str] = set()

    try:
        r = requests.get(listing_url, headers=HEADERS, timeout=25)
        if r.status_code != 200:
            print(f"    Listing {listing_url}: HTTP {r.status_code}")
            return links
        soup = BeautifulSoup(r.text, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            # Normalise relative URLs
            if href.startswith("/"):
                href = BASE_URL + href
            if not href.startswith(BASE_URL):
                continue
            path = urlparse(href).path
            # Must be /category/product-slug/ (exactly two path segments)
            if (
                path.startswith(cat_path)
                and re.match(r"^/[a-z]+/[a-z0-9-]+/$", path)
                and href not in seen
            ):
                seen.add(href)
                links.append(href)
    except Exception as e:
        print(f"    Listing error ({listing_url}): {e}")

    return links


def scrape_product(url: str) -> list[dict]:
    """
    Scrape a SkLO product page.
    Extracts shared fields from HTML and reads allVariants JS for per-variant rows.
    Returns one dict per variant (or one dict if no variant data found).
    """
    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        if r.status_code != 200:
            return [{"Source": url}]
    except Exception as e:
        print(f"    Fetch error ({url}): {e}")
        return [{"Source": url}]

    soup = BeautifulSoup(r.text, "html.parser")
    base: dict = {"Source": url}

    # --- Product Name: <h1 class="single-product__title"><strong>…</strong><span>…</span> ---
    h1 = soup.find("h1", class_="single-product__title")
    if h1:
        strong = h1.find("strong")
        span = h1.find("span")
        parts = [clean_text(el.get_text()) for el in [strong, span] if el and clean_text(el.get_text())]
        if parts:
            base["Product Name"] = " ".join(parts)
    if not base.get("Product Name"):
        h1_any = soup.find("h1")
        if h1_any:
            base["Product Name"] = clean_text(h1_any.get_text())

    # --- Collection ---
    coll_el = soup.find(class_="single-product__collection")
    if coll_el:
        h3 = coll_el.find("h3")
        if h3:
            base["Collection"] = clean_text(h3.get_text())

    # --- Description ---
    desc_el = soup.find(class_="single-product__description")
    if desc_el:
        base["Description"] = clean_text(desc_el.get_text(separator=" "))

    # --- Image URL: og:image is the canonical main product image ---
    og_img = soup.find("meta", property="og:image")
    if og_img and og_img.get("content"):
        base["Image URL"] = og_img["content"]

    # --- Product Family Id ---
    if base.get("Product Name"):
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    # --- Variants from allVariants JS ---
    all_variants = _extract_all_variants(soup)
    if not all_variants:
        # No variant data — single-row fallback
        # Try base price
        price_el = soup.find(id="basePrice")
        if price_el:
            base["Price"] = clean_price(price_el.get_text())
        return [base]

    rows: list[dict] = []
    for v in all_variants:
        row = base.copy()

        sku = v.get("sku", "")
        if sku:
            row["SKU"] = sku

        price_usd = v.get("price_usd") or v.get("price_msrp")
        if price_usd is not None:
            row["Price"] = clean_price(str(price_usd))

        # Dimensions — per-variant (may differ for products with size options)
        dims_str = v.get("dimensions", "")
        if dims_str:
            parsed = _parse_sklo_dimensions(dims_str)
            row.update(parsed)

        # Finish fields
        glass_color = v.get("glass_color", "")
        if glass_color:
            row["Finish"] = glass_color

        metal_finish = v.get("metal_finish", "")
        if metal_finish:
            row["Metal Finish"] = metal_finish

        # Size option (e.g. "small", "large")
        option = v.get("option", "")
        if option:
            row["Size"] = option

        # Lighting extras
        if v.get("cord_quantity"):
            row["Cord Quantity"] = v["cord_quantity"]
        if v.get("drop_quantity"):
            row["Drop Quantity"] = v["drop_quantity"]

        if row.get("Product Name"):
            row["Product Family Id"] = extract_family_id(row["Product Name"])

        rows.append(row)

    return rows if rows else [base]


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

        seen_urls: set[str] = set()
        all_product_urls: list[str] = []
        for listing_url in cat["links"]:
            for u in get_product_links(listing_url):
                if u not in seen_urls:
                    seen_urls.add(u)
                    all_product_urls.append(u)

        if TEST_MODE:
            all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

        print(f"  {cat['name']}: {len(all_product_urls)} products")

        global_idx = 1
        for url in all_product_urls:
            try:
                variant_rows = scrape_product(url)
                for row in variant_rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
            except Exception as e:
                print(f"    ERROR on {url}: {e}")
            time.sleep(0.8)

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
