"""
scraper.py  —  Hooker
----------------------
Platform: hookerfurniture.com (listing) + hookerfurnishings.com (products + listings)

Site structure:
  vendor_info.json contains hookerfurniture.com URLs, but the actual product
  listings and detail pages are on hookerfurnishings.com.

  URL mapping (hookerfurniture.com → hookerfurnishings.com):
    /bedroom/nightstands/room-type.aspx  →  /bedroom/nightstands
    /dining-room/tables/room-type.aspx   →  /dining-room/tables
    /office-furniture/desks/category-type.aspx → /office-furniture/desks
    /itembrowser.aspx?room=living-room&type=ottomans → /living-room/ottomans

  Listing   : hookerfurnishings.com/{path}?page=N
              Product hrefs: a[href] with single slug containing digits
  Product   : hookerfurnishings.com/{slug}

Product page fields (SPA page with JSON-LD + data block):
  Product Name    : h1 or JSON-LD name
  SKU             : data block "SKU: {value}"
  Price           : data block "MSRP Price: {value}"
  Image URL       : first img[itemprop="image"] or og:image
  Description     : p[itemprop="description"] or meta[name="description"]
  Width/Height/Depth/Diameter/Length: data block
  Finish          : data block "Finish Description: {value}"
  Materials       : data block "Material Description: {value}"
  Collection      : data block "Marketing Collection Name: {value}"
  Features        : data block "Feature Bullets: {value}"
  UPC             : data block "UPC: {value}"
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

import requests as _requests

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Hooker")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

LISTING_BASE = "https://hookerfurnishings.com"
PRODUCT_BASE = "https://hookerfurnishings.com"
TIMEOUT_MS   = 60_000


def _to_furnishings_url(hooker_url: str) -> str:
    """
    Convert hookerfurniture.com listing URL to hookerfurnishings.com category URL.
    e.g. https://www.hookerfurniture.com/bedroom/nightstands/room-type.aspx
         → https://hookerfurnishings.com/bedroom/nightstands
    """
    parsed = urlparse(hooker_url)

    # Handle itembrowser.aspx URLs (query-param based categories)
    if "itembrowser.aspx" in hooker_url:
        params = parse_qs(parsed.query)
        room  = params.get("room", [""])[0]   # e.g. "living-room"
        types = params.get("type", [])         # may appear multiple times
        cat   = types[-1] if types else ""     # take the last unique type value
        if room and cat:
            return f"{LISTING_BASE}/{room}/{cat}"
        return ""

    # Standard path: remove /{room|category|department}-type.aspx and query params
    path = re.sub(r"/(room|category|department)-type\.aspx.*$", "", parsed.path)
    path = path.rstrip("/")
    return f"{LISTING_BASE}{path}"


def _parse_data_block(html: str) -> dict[str, str]:
    """Parse the product data div: 'Key: Value<br>Key2: Value2<br>...' """
    result: dict[str, str] = {}
    cleaned = re.sub(r"<(?!br\s*/?>)[^>]+>", "", html, flags=re.IGNORECASE)
    for segment in re.split(r"<br\s*/?>", cleaned, flags=re.IGNORECASE):
        segment = re.sub(r"&[a-z]+;", " ", segment).strip()
        if ": " in segment:
            key, _, val = segment.partition(": ")
            key = key.strip().lower()
            val = val.strip()
            if key and val and len(key) < 80:
                result[key] = val
    return result


_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}
_SLUG_RE = re.compile(
    r'href="https://hookerfurnishings\.com/([a-z0-9][a-z0-9-]*[0-9][a-z0-9-]*)"'
)


def _fetch_listing_links(url: str) -> list[str]:
    """Fetch a listing page via requests and extract product slugs from SSR HTML."""
    try:
        r = _requests.get(url, headers=_HEADERS, timeout=30)
        slugs = list(dict.fromkeys(_SLUG_RE.findall(r.text)))
        return [f"{LISTING_BASE}/{s}" for s in slugs]
    except Exception as e:
        print(f"  [WARN] requests fetch failed for {url}: {e}")
        return []


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect product URLs from hookerfurnishings.com category listing via SSR HTML."""
    furnishings_url = _to_furnishings_url(listing_url)
    if not furnishings_url:
        print(f"  [WARN] Cannot convert URL: {listing_url}")
        return []

    print(f"  [Listing] {furnishings_url}")
    all_links: list[str] = []
    seen: set[str] = set()
    page_num = 1

    while True:
        url = furnishings_url if page_num == 1 else f"{furnishings_url}?page={page_num}"
        links = _fetch_listing_links(url)

        new_links = [l for l in links if l not in seen]
        for l in new_links:
            seen.add(l)
        all_links.extend(new_links)

        if not new_links:
            break

        page_num += 1

    print(f"  [Listing] {len(all_links)} products across {page_num} pages")
    return all_links


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Hooker Furnishings product detail page."""
    row: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(4000)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [row]

    # ── 1. Product Name ────────────────────────────────────────────────────────
    for sel in ["h1[itemprop='name']", "h1.page-title-wrapper h1", "h1"]:
        el = await page.query_selector(sel)
        if el:
            text = clean_text(await el.inner_text())
            if text and len(text) < 150:
                row["Product Name"] = text
                break

    # ── 2. Description ─────────────────────────────────────────────────────────
    desc_el = await page.query_selector("p[itemprop='description']")
    if desc_el:
        text = clean_text(await desc_el.inner_text())
        if text:
            row["Description"] = text

    # ── 3. Image URL ───────────────────────────────────────────────────────────
    img_el = await page.query_selector("img[itemprop='image']")
    if img_el:
        src = (
            await img_el.get_attribute("src")
            or await img_el.get_attribute("data-src")
            or ""
        )
        if src and not src.startswith("data:"):
            if not src.startswith("http"):
                src = urljoin(PRODUCT_BASE, src)
            row["Image URL"] = src
    if not row.get("Image URL"):
        og = await page.query_selector("meta[property='og:image']")
        if og:
            content = await og.get_attribute("content")
            if content:
                row["Image URL"] = content

    # ── 4. Parse data block (Key: Value<br> pairs) ─────────────────────────────
    # Find the SMALLEST div containing "SKU: " to avoid matching the page wrapper
    data_block_html = await page.evaluate("""
        () => {
            const divs = document.querySelectorAll('div, article, section');
            let best = null, bestLen = Infinity;
            for (const d of divs) {
                const h = d.innerHTML || '';
                if (h.includes('SKU: ') && h.includes('<br')) {
                    if (h.length < bestLen) {
                        bestLen = h.length;
                        best = h;
                    }
                }
            }
            return best || '';
        }
    """)

    if data_block_html:
        data = _parse_data_block(data_block_html)

        if data.get("sku") and not row.get("SKU"):
            row["SKU"] = data["sku"]

        msrp = data.get("msrp price", data.get("msrp", ""))
        if msrp:
            row["Price"] = clean_price(msrp)

        if data.get("weight"):
            m = re.search(r"([\d.]+)", data["weight"])
            if m:
                row["Weight"] = m.group(1)

        # Individual dimension fields
        for key, col in [("height", "Height"), ("width", "Width"), ("depth", "Depth"),
                         ("diameter", "Diameter"), ("length", "Length")]:
            if data.get(key) and not row.get(col):
                m = re.search(r"([\d.]+)", data[key])
                if m:
                    row[col] = m.group(1)

        # Build combined Dimensions string
        dim_parts = []
        for col, label in [("Width", "W"), ("Depth", "D"), ("Height", "H"),
                            ("Diameter", "Dia"), ("Length", "L")]:
            if row.get(col):
                dim_parts.append(f"{label} {row[col]}")
        if dim_parts:
            row.setdefault("Dimensions", " x ".join(dim_parts))

        if data.get("finish description"):
            row["Finish"] = data["finish description"]
        if data.get("material description"):
            row["Materials"] = data["material description"]

        collection = (
            data.get("marketing collection name")
            or data.get("collection filter")
            or data.get("suite", "")
        )
        if collection:
            row["Collection"] = collection

        if data.get("style"):
            row["Style"] = data["style"]

        features = data.get("feature bullets", data.get("features", ""))
        if features:
            row["Features"] = re.sub(r"\|", ", ", features)

        if data.get("upc"):
            row["UPC"] = data["upc"]

        if data.get("carton height") or data.get("carton width") or data.get("carton length"):
            parts = []
            for k in ("carton length", "carton width", "carton height"):
                if data.get(k):
                    parts.append(data[k])
            if parts:
                row["Carton Size"] = " x ".join(parts)

        if data.get("carton weight"):
            row["Carton Weight"] = data["carton weight"]

        skip = {
            "product name", "sku", "msrp price", "msrp", "weight",
            "height", "width", "depth", "diameter", "length",
            "finish description", "material description",
            "marketing collection name", "collection filter", "style",
            "feature bullets", "features", "upc",
            "carton height", "carton width", "carton length", "carton weight",
            "modular items", "modular parent", "ac downloadable product", "ac gift card",
            "name", "description", "status date", "item web rank", "item cover rank",
            "brand (code)", "line disc", "erp status", "image role data",
            "vendor", "vendor number", "consolidated warehouses", "view type",
            "is fabric available", "is leather available", "in stock",
            "intro date", "parent sku (namedconfig)", "sub type",
            "alternate finish items", "alternate cover items", "alternate bed sizes",
        }
        for k, v in data.items():
            if k not in skip and v and v not in ("-", "No", ""):
                col_name = k.title()
                row.setdefault(col_name, v)

    # ── 5. JSON-LD fallback ────────────────────────────────────────────────────
    if not row.get("SKU") or not row.get("Image URL"):
        ld_scripts = await page.evaluate("""
            () => Array.from(document.querySelectorAll('script[type="application/ld+json"]'))
                       .map(s => s.textContent)
        """)
        for raw in ld_scripts:
            try:
                d = json.loads(raw)
                candidates = (
                    d.get("@graph", [d]) if isinstance(d, dict)
                    else (d if isinstance(d, list) else [d])
                )
                for obj in candidates:
                    if obj.get("@type") == "Product":
                        if not row.get("SKU") and obj.get("sku"):
                            row["SKU"] = str(obj["sku"])
                        if not row.get("Image URL") and obj.get("image"):
                            img = obj["image"]
                            row["Image URL"] = img if isinstance(img, str) else img[0]
                        break
            except Exception:
                pass

    # ── 6. Product Family Id ───────────────────────────────────────────────────
    if not row.get("Product Family Id") and row.get("Product Name"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    return [row]


async def main() -> None:
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: all {len(categories)} categories, max {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor : {info['vendor_name']}")
    print(f"[Scraper] Mode   : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output : {OUTPUT_PATH}")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(cat["name"], cat["links"][0], studio_columns=cat["studio_columns"])

            seen_urls: set[str] = set()
            all_urls:  list[str] = []

            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_urls.append(u)

            if TEST_MODE:
                all_urls = all_urls[:TEST_MAX_PRODUCTS]

            print(f"\n[Category] {cat['name']}: {len(all_urls)} products")

            for idx, url in enumerate(all_urls, 1):
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        row["Manufacturer"] = info["vendor_name"]
                        writer.write_row(row, category_name=cat["name"])
                    print(f"  [{idx}] {url.rstrip('/').split('/')[-1]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay(0.8, 2.0)

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
