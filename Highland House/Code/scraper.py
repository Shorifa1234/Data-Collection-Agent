"""
scraper.py  —  Highland House
--------------------------------
Platform: highlandhousefurniture.com (custom ASP.NET)
Uses requests + BeautifulSoup — no Playwright needed.

Site structure:
  - Listing: /Consumer/ShowItems.aspx?TypeID=NN
    All products appear on a single page (no pagination, confirmed).
    Product links: /Consumer/ShowItemDetail.aspx?SKU=HH27-120B
  - Product detail: /Consumer/ShowItemDetail.aspx?SKU=XXXX
    - SKU in #hSKU (hidden input value)
    - Product name in h1 (page title)
    - Dimensions in #dimensionDiv → #width, #depth, #height spans
    - Description in static paragraphs
    - Images: /ProductCatalog/prod-images/[filename]_hires.jpg
    - Price: NOT displayed (trade-only site)
    - Finishes: AJAX-loaded (not in static HTML) — captured as text list if present

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Highland House"
    python orchestrator.py "Highland House" --test
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
    parse_dimensions,
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Highland House")
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://highlandhousefurniture.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}


def get_html(url: str) -> BeautifulSoup | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        if r.status_code != 200:
            print(f"    HTTP {r.status_code}: {url}")
            return None
        return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        print(f"    Fetch error {url}: {e}")
        return None


def get_product_links_from_listing(listing_url: str) -> list[str]:
    """
    Extract all ShowItemDetail links from a category listing page.
    All products are on one page — no pagination on Highland House listings.
    """
    soup = get_html(listing_url)
    if not soup:
        return []

    seen: set[str] = set()
    links: list[str] = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "ShowItemDetail.aspx" in href:
            if href.startswith("/"):
                href = BASE_URL + href
            elif not href.startswith("http"):
                href = BASE_URL + "/" + href.lstrip("/")
            href = href.split("?")[0] + "?" + href.split("?")[1] if "?" in href else href
            if href not in seen:
                seen.add(href)
                links.append(href)
    return links


def scrape_product(product_url: str) -> dict:
    """
    Scrape one Highland House product detail page.
    Extracts: Name, SKU, Dimensions, Description, Image URL, Materials.
    Price is not displayed on the site (trade-only).
    """
    data: dict = {"Source": product_url}
    soup = get_html(product_url)
    if not soup:
        return data

    # ── SKU ──────────────────────────────────────────────────────────────────
    sku_input = soup.find("input", id="hSKU") or soup.find("input", {"id": re.compile(r"sku", re.I)})
    if sku_input:
        data["SKU"] = clean_text(sku_input.get("value", ""))
    # Fallback: parse SKU from URL query param
    if not data.get("SKU"):
        m = re.search(r"[?&]SKU=([^&]+)", product_url, re.I)
        if m:
            data["SKU"] = m.group(1)

    # ── Product Name ─────────────────────────────────────────────────────────
    h1 = soup.find("h1")
    if h1:
        name = clean_text(h1.get_text())
        # Strip leading SKU prefix if present (e.g. "HH27-120B - Palma Dining Table Base")
        if " - " in name:
            name = name.split(" - ", 1)[1].strip()
        data["Product Name"] = name
        data["Product Family Id"] = extract_family_id(name)

    # ── Dimensions ───────────────────────────────────────────────────────────
    # Rendered in #dimensionDiv with child spans #width, #depth, #height
    dim_div = soup.find(id="dimensionDiv") or soup.find("div", class_=re.compile(r"dimension", re.I))
    if dim_div:
        def _dim(el_id: str) -> str:
            el = dim_div.find(id=el_id) or dim_div.find(class_=el_id)
            return clean_text(el.get_text()) if el else ""

        w = _dim("width")
        d = _dim("depth")
        h = _dim("height")
        parts = [f"W {w}" if w else "", f"D {d}" if d else "", f"H {h}" if h else ""]
        dim_str = " x ".join(p for p in parts if p)
        if dim_str:
            dims = parse_dimensions(dim_str)
            data.update(dims)

    # ── Description ──────────────────────────────────────────────────────────
    # Description is usually in a paragraph near the product title or a .description div
    desc_candidates = [
        soup.find("div", class_=re.compile(r"product[-_]?desc|description|detail", re.I)),
        soup.find("div", id=re.compile(r"desc|detail", re.I)),
    ]
    for el in desc_candidates:
        if el:
            text = clean_text(el.get_text(separator=" "))
            if len(text) > 20:
                data["Description"] = text
                break

    # Fallback: grab paragraphs near the product heading
    if not data.get("Description"):
        for p in soup.find_all("p"):
            text = clean_text(p.get_text())
            if len(text) > 40 and not text.startswith("Copyright"):
                data["Description"] = text
                break

    # ── Materials ─────────────────────────────────────────────────────────────
    mat_el = soup.find(string=re.compile(r"material", re.I))
    if mat_el and mat_el.parent:
        sib = mat_el.parent.find_next_sibling()
        if sib:
            data["Materials"] = clean_text(sib.get_text())

    # ── Image URL (high-res) ──────────────────────────────────────────────────
    # Highland House images: /ProductCatalog/prod-images/[filename]_hires.jpg
    # Main carousel first
    for img in soup.find_all("img"):
        src = img.get("data-zoom-image") or img.get("data-src") or img.get("src", "")
        if src and "prod-images" in src and "_hires" in src:
            data["Image URL"] = (BASE_URL + src) if src.startswith("/") else src
            break
    # Fallback: any product catalog image
    if not data.get("Image URL"):
        for img in soup.find_all("img"):
            src = img.get("src", "")
            if src and "prod-images" in src:
                # Upgrade to hires version
                hires = re.sub(r"\.(jpg|jpeg|png|webp)$", r"_hires.\1", src, flags=re.I)
                data["Image URL"] = (BASE_URL + hires) if hires.startswith("/") else hires
                break

    # ── Collection / tags ─────────────────────────────────────────────────────
    # Try meta tags or breadcrumbs for collection context
    meta_cat = soup.find("meta", property="product:category") or soup.find("meta", attrs={"name": "category"})
    if meta_cat:
        data.setdefault("Collection", meta_cat.get("content", ""))

    return data


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

    global_idx = 1

    for cat in categories:
        if not cat["links"]:
            continue

        cat_url = cat["links"][0]
        writer.add_sheet(cat["name"], cat_url, studio_columns=cat["studio_columns"])

        # Collect product URLs from all links, dedup by SKU query param
        seen_skus: set[str] = set()
        product_urls: list[str] = []
        for listing_url in cat["links"]:
            for u in get_product_links_from_listing(listing_url):
                m = re.search(r"[?&]SKU=([^&]+)", u, re.I)
                key = m.group(1) if m else u
                if key not in seen_skus:
                    seen_skus.add(key)
                    product_urls.append(u)

        if TEST_MODE:
            product_urls = product_urls[:TEST_MAX_PRODUCTS]

        print(f"  {cat['name']}: {len(product_urls)} products across {len(cat['links'])} link(s)")

        for product_url in product_urls:
            try:
                data = scrape_product(product_url)
                if not data.get("SKU"):
                    data["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                if not data.get("Product Family Id") and data.get("Product Name"):
                    data["Product Family Id"] = extract_family_id(data["Product Name"])
                writer.write_row(data, category_name=cat["name"])
                global_idx += 1
            except Exception as e:
                print(f"    ERROR {product_url}: {e}")
            time.sleep(0.8)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
