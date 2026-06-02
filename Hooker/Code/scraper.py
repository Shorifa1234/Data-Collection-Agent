"""
scraper.py  —  Hooker Furniture
---------------------------------
Platform: hookerfurniture.com (listing URLs) + hookerfurnishings.com (actual pages)

Site structure:
  Listing   : Magento 2 GraphQL API at hookerfurnishings.com/graphql
              Uses category_uid filter (base64-encoded IDs in vendor_info.json)
  Product   : hookerfurnishings.com/{slug}  (loaded via Playwright)

Page selectors discovered:
  SKU       : p[class*='productSku']  → "SKU: 6820-90116-99"
  Dimensions: div[class*='dimensionsAndDesign-sectionValue']
              → first p contains <b>Label:</b>, second p[class*='unit'] = value
  Specs     : p[class*='productSpecification-itemLabel'] → label
              next sibling → value
  Description: p[itemprop='description'] or meta[name='description']
  Image     : img[itemprop='image'] or og:image
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

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
from bs4 import BeautifulSoup

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Hooker Furniture")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / "Hooker" / "Data" / "Hooker Furniture.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

PRODUCT_BASE = "https://hookerfurnishings.com"
GQL_URL      = "https://hookerfurnishings.com/graphql"
TIMEOUT_MS   = 60_000

# Sub-brands under Hooker Furnishings — these go in Brand column, not Manufacturer
_HOOKER_BRANDS = {"hooker furniture", "hooker furnishings"}

_GQL_HEADERS = {
    "Content-Type": "application/json",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
}

_GQL_QUERY = """
query GetCategoryProducts($uid: String!, $page: Int!) {
  products(
    filter: { category_uid: { eq: $uid } }
    pageSize: 48
    currentPage: $page
    sort: { name: ASC }
  ) {
    total_count
    page_info { total_pages current_page }
    items {
      name
      sku
      url_key
      url_suffix
    }
  }
}
"""


def _gql_fetch_category(uid: str) -> list[str]:
    """Fetch all product URLs for a GraphQL category UID via Magento 2 API."""
    urls: list[str] = []
    page = 1
    total_pages = 1

    while page <= total_pages:
        try:
            r = _requests.post(
                GQL_URL,
                headers=_GQL_HEADERS,
                json={"query": _GQL_QUERY, "variables": {"uid": uid, "page": page}},
                timeout=30,
            )
            data = r.json()
            prods = data.get("data", {}).get("products", {})
            if page == 1:
                total_pages = prods.get("page_info", {}).get("total_pages", 1)
                print(f"    UID {uid}: {prods.get('total_count', 0)} products, {total_pages} pages")
            for item in prods.get("items", []):
                url_key    = item.get("url_key", "")
                url_suffix = item.get("url_suffix", "")
                if url_key:
                    urls.append(f"{PRODUCT_BASE}/{url_key}{url_suffix}")
        except Exception as e:
            print(f"    [WARN] GQL fetch failed uid={uid} page={page}: {e}")
            break
        page += 1

    return urls


async def get_product_links(page, cat: dict) -> list[str]:
    """Return all product URLs for a category using GraphQL UIDs."""
    gql_uids = cat.get("gql_uids", [])
    if not gql_uids:
        print(f"  [WARN] No gql_uids for category — skipping")
        return []

    seen: set[str] = set()
    all_urls: list[str] = []

    for uid in gql_uids:
        for url in _gql_fetch_category(uid):
            if url not in seen:
                seen.add(url)
                all_urls.append(url)

    print(f"  [Listing] {len(all_urls)} unique products (across {len(gql_uids)} UID(s))")
    return all_urls


# ── Column name mapping (spec label → Excel column) ───────────────────────────
_LABEL_MAP = {
    "alternate finish items": "Alternate Finish Items",
    "alternate cover items": "Alternate Cover Items",
    "alternate bed sizes": "Alternate Bed Sizes",
    "arm height": "Arm Height",
    "back panel description": "Back Panel Description",
    "back panel material": "Back Panel Material",
    "brand": "Brand",
    "carton weight": "Carton Weight",
    "clearance": "Clearance",
    "collection": "Collection",
    "collection features": "Collection Features",
    "color family": "Color Family",
    "consolidated warehouses": "Consolidated Warehouses",
    "custom": "Custom",
    "depth": "Depth",
    "depth range": "Depth Range",
    "diameter": "Diameter",
    "distressing": "Distressing",
    "div": "Div",
    "distance from wall to recline": "Distance from Wall to Recline",
    "drawer box construction method": "Drawer Box Construction Method",
    "drawer construction": "Drawer Construction",
    "drawers": "Drawers",
    "feature filter": "Feature Filter",
    "finish filter": "Finish Filter",
    "finish construction": "Finish Construction",
    "frame construction": "Frame Construction",
    "full recline length": "Full Recline Length",
    "height": "Height",
    "height range": "Height Range",
    "in stock": "In Stock",
    "leather": "Leather",
    "leather type": "Leather Type",
    "levelers": "Levelers",
    "marketing collection name": "Collection",
    "material filter": "Material Filter",
    "number of drawers": "Number of Drawers",
    "padding": "Padding",
    "product care": "Product Care",
    "seat": "Seat",
    "seat back": "Seat Back",
    "seat depth": "Seat Depth",
    "seat height": "Seat Height",
    "seat width": "Seat Width",
    "sub type": "Sub Type",
    "suite": "Suite",
    "tipover restraint included": "Tipover Restraint Included",
    "top coat material": "Top Coat Material",
    "top load capacity": "Top Load Capacity",
    "upc": "UPC",
    "visible in 3d": "Visible In 3D",
    "volume": "Volume",
    "weight": "Weight",
    "weight capacity": "Weight Capacity",
    "width": "Width",
    "width range": "Width Range",
    "wood distressing type": "Wood Distressing Type",
    "wood joinery type": "Wood Joinery Type",
    "1st row drawer dimensions": "1st Row Drawer Dimensions",
    "1st row drawer weight capacity": "1st Row Drawer Weight Capacity",
    "2nd and 3rd row drawer dimensions": "2nd and 3rd Row Drawer Dimensions",
    "2nd and 3rd row drawer weight capacity": "2nd and 3rd Row Drawer Weight Capacity",
}

# Labels to skip (internal/nav/irrelevant)
_SKIP_LABELS = {
    "brand", "category", "sub category", "subcategory", "type",
    "name", "status date", "item web rank", "item cover rank",
    "line disc", "erp status", "image role data", "vendor", "vendor number",
    "view type", "is fabric available", "is leather available", "intro date",
    "parent sku (namedconfig)", "modular items", "modular parent",
    "ac downloadable product", "ac gift card", "brand (code)",
}

# Dimension/measurement labels — strip trailing unit suffixes
_DIM_LABELS = {"width", "depth", "height", "diameter", "length", "weight",
               "arm height", "seat width", "seat height", "seat depth",
               "full recline length", "distance from wall to recline", "clearance"}


def _store_spec(label_raw: str, value: str, row: dict) -> None:
    """Store label/value spec into row dict using canonical column name."""
    label = label_raw.strip().rstrip(":").strip()
    if not label or not value or len(label) > 80:
        return
    label_lower = label.lower()
    if label_lower in _SKIP_LABELS:
        return

    col = _LABEL_MAP.get(label_lower, label.title())

    # Strip units from dimension fields
    if label_lower in _DIM_LABELS:
        val_clean = re.sub(r"\s*(in\.?|ft\.?|cm\.?|lbs?\.?)$", "", value, flags=re.I).strip()
        row.setdefault(col, val_clean)
    elif col == "Brand":
        # Only store Brand if it's a sub-brand (not Hooker itself)
        if value.lower() not in _HOOKER_BRANDS:
            row.setdefault("Brand", value)
    elif col == "Collection" and row.get("Collection"):
        pass  # don't overwrite
    else:
        row.setdefault(col, value)


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Hooker Furnishings product detail page."""
    row: dict = {"Source URL": url, "Manufacturer": VENDOR_NAME}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(3500)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [row]

    # ── 1. Product Name (h1) ──────────────────────────────────────────────────
    for sel in ["h1[itemprop='name']", "h1[aria-live='polite']", "h1"]:
        el = await page.query_selector(sel)
        if el:
            text = clean_text(await el.inner_text())
            if text and len(text) < 200:
                row["Product Name"] = text
                row["Product Family Id"] = extract_family_id(text)
                break

    # ── 2. SKU ────────────────────────────────────────────────────────────────
    sku_el = await page.query_selector("[class*='productSku']")
    if sku_el:
        sku_text = clean_text(await sku_el.inner_text())
        m = re.search(r"SKU:\s*([A-Za-z0-9\-]+)", sku_text)
        if m:
            row["SKU"] = m.group(1)

    # ── 3. Description ────────────────────────────────────────────────────────
    for sel in [
        "p[itemprop='description']",
        "[class*='description'] p",
        "meta[name='description']",
    ]:
        el = await page.query_selector(sel)
        if el:
            if sel == "meta[name='description']":
                text = clean_text(await el.get_attribute("content") or "")
            else:
                text = clean_text(await el.inner_text())
            if text and len(text) > 20:
                row["Description"] = text
                break

    # ── 4. Image URL ──────────────────────────────────────────────────────────
    for sel in [
        "img[itemprop='image']",
        "[class*='productImage'] img",
        "[class*='gallery'] img",
        "picture img",
    ]:
        img_el = await page.query_selector(sel)
        if img_el:
            src = (
                await img_el.get_attribute("data-zoom-image")
                or await img_el.get_attribute("data-src")
                or await img_el.get_attribute("src")
                or ""
            )
            if src and not src.startswith("data:"):
                if src.startswith("//"):
                    src = "https:" + src
                elif not src.startswith("http"):
                    src = urljoin(PRODUCT_BASE, src)
                row["Image URL"] = src
                break

    if not row.get("Image URL"):
        og = await page.query_selector("meta[property='og:image']")
        if og:
            content = await og.get_attribute("content")
            if content:
                row["Image URL"] = content

    # ── 5. Click Overall Dimensions tab and parse ────────────────────────────
    try:
        dim_btn = await page.query_selector("button:has-text('Overall Dimensions')")
        if dim_btn:
            await dim_btn.click()
            await page.wait_for_timeout(800)
    except Exception:
        pass

    html = await page.content()
    soup = BeautifulSoup(html, "lxml")

    # Dimension rows: div[class*='dimensionsAndDesign-sectionValue']
    # Structure: <div><p><b>Label:</b></p><p class="*unit*">Value</p></div>
    for div in soup.select("[class*='dimensionsAndDesign-sectionValue']"):
        label_p = div.select_one("b")
        value_p = div.select_one("[class*='unit']")
        if label_p and value_p:
            label = clean_text(label_p.get_text())
            value = clean_text(value_p.get_text())
            _store_spec(label, value, row)

    # ── 6. Click Design Elements & Features tab ───────────────────────────────
    try:
        feat_btn = await page.query_selector("button:has-text('Design Elements & Features')")
        if feat_btn:
            await feat_btn.click()
            await page.wait_for_timeout(800)
            html2 = await page.content()
            soup = BeautifulSoup(html2, "lxml")
    except Exception:
        pass

    # ── 7. Parse all productSpecification-itemLabel rows ─────────────────────
    for label_el in soup.select("[class*='productSpecification-itemLabel']"):
        label = clean_text(label_el.get_text())
        sib = label_el.find_next_sibling()
        if sib:
            value = clean_text(sib.get_text())
            _store_spec(label, value, row)

    # ── 8. Click Product Details, Finish/Frame Construction, Product Care ─────
    for section_label in ["Product Details", "Finish Construction", "Frame Construction", "Product Care"]:
        try:
            sec_btn = await page.query_selector(f"button:has-text('{section_label}')")
            if sec_btn:
                await sec_btn.click()
                await page.wait_for_timeout(500)
        except Exception:
            pass

    html3 = await page.content()
    soup3 = BeautifulSoup(html3, "lxml")

    # Parse spec items again (new sections may now be expanded)
    for label_el in soup3.select("[class*='productSpecification-itemLabel']"):
        label = clean_text(label_el.get_text())
        sib = label_el.find_next_sibling()
        if sib:
            value = clean_text(sib.get_text())
            _store_spec(label, value, row)

    # Also parse section content paragraphs (Finish Construction, Frame Construction text)
    for section_title in soup3.select("[class*='productSpecification-sectionTitle']"):
        section_name = clean_text(section_title.get_text())
        # Find next sibling paragraph with the description
        sib = section_title.find_next_sibling()
        if sib:
            val = clean_text(sib.get_text())
            if val and section_name:
                col = _LABEL_MAP.get(section_name.lower(), section_name.title())
                row.setdefault(col, val)

    # ── 9. Seating options (Arm Height, Seat dims visible from View Summary) ──
    # These also appear in productSpecification-itemLabel — already parsed above

    # ── 10. Build Dimensions string ───────────────────────────────────────────
    dim_parts = []
    for col, label in [("Width", "W"), ("Depth", "D"), ("Height", "H"),
                        ("Diameter", "Dia"), ("Length", "L")]:
        if row.get(col):
            dim_parts.append(f"{label} {row[col]}")
    if dim_parts:
        row.setdefault("Dimensions", " x ".join(dim_parts))

    # ── 11. JSON-LD fallback ──────────────────────────────────────────────────
    for script in soup3.find_all("script", type="application/ld+json"):
        try:
            d = json.loads(script.string or "{}")
            items = d.get("@graph", [d]) if isinstance(d, dict) else (d if isinstance(d, list) else [d])
            for obj in items:
                if obj.get("@type") == "Product":
                    if not row.get("SKU") and obj.get("sku"):
                        row["SKU"] = str(obj["sku"])
                    if not row.get("Image URL") and obj.get("image"):
                        img = obj["image"]
                        row["Image URL"] = img if isinstance(img, str) else img[0]
                    if not row.get("Product Name") and obj.get("name"):
                        row["Product Name"] = clean_text(obj["name"])
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    if not row.get("Price"):
                        offers = obj.get("offers", {})
                        if isinstance(offers, list):
                            offers = offers[0]
                        price = offers.get("price") or offers.get("lowPrice", "")
                        if price:
                            row["Price"] = clean_price(str(price))
                    break
        except Exception:
            pass

    return [row]


async def main() -> None:
    info_path = Path(__file__).parent / "vendor_info.json"
    info      = json.loads(info_path.read_text(encoding="utf-8"))

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: {len(categories)} categories, max {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor  : {VENDOR_NAME}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")
    print(f"[Scraper] Headless: {HEADLESS}")

    writer = ExcelWriter(OUTPUT_PATH, VENDOR_NAME)

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat.get("gql_uids") and not cat.get("links"):
                print(f"[Skip] {cat['name']} — no gql_uids or links")
                continue

            print(f"\n[Category] {cat['name']}")
            source_url = cat["links"][0] if cat.get("links") else ""
            writer.add_sheet(
                cat["name"], source_url,
                studio_columns=cat.get("studio_columns", []),
            )

            all_urls = await get_product_links(page, cat)

            if TEST_MODE:
                all_urls = all_urls[:TEST_MAX_PRODUCTS]

            print(f"  {len(all_urls)} products to scrape")

            global_idx = 1
            for url in all_urls:
                slug = url.rstrip("/").split("/")[-1][:60]
                print(f"  [{global_idx}/{len(all_urls)}] {slug}")
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
                            print(f"    [SKU generated] {row['SKU']}")
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                    global_idx += 1

                await async_polite_delay(0.8, 2.0)

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
