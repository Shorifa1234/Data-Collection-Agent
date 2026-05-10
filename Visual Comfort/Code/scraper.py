"""
scraper.py  -  Visual Comfort
------------------------------
Scrapes all product data from visualcomfort.com for every category
defined in the SD tracker, with fully dynamic column output.

KEY BEHAVIOUR: One row per finish/size/color variant.
- JSON-LD offers array is the primary source of all variants.
- If only one offer is present, falls back to clicking UI swatch buttons.
- Fields that change per variant  : SKU, Price, Finish, Image URL
- Fields that stay the same       : Product Name, Description, dims, specs…

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Visual Comfort"        # full run
    python orchestrator.py "Visual Comfort" --test # test run

Env vars (set by orchestrator):
    HEADLESS             true | false   (default: true)
    OUTPUT_PATH          absolute path to output .xlsx
    VENDOR_NAME          vendor name string
    TEST_MODE            true | false   (default: false)
    TEST_MAX_CATEGORIES  max categories in test mode (default: 2)
    TEST_MAX_PRODUCTS    max products per category in test mode (default: 5)
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from bs4 import BeautifulSoup

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
    safe_float,
)

# ── config ───────────────────────────────────────────────────────────────────
VENDOR_NAME = os.environ.get("VENDOR_NAME", "Visual Comfort")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)

TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://www.visualcomfort.com"
TIMEOUT_MS = 60_000


# ── listing helpers ───────────────────────────────────────────────────────────
async def get_product_links(page, listing_url: str, max_products: int | None = None) -> list[str]:
    """Collect product URLs from a category page, handling pagination."""
    links: list[str] = []
    seen:  set[str]  = set()

    await page.goto(listing_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
    await page.wait_for_timeout(5000)

    page_num = 1

    while True:
        try:
            await page.wait_for_selector("li.product-card", timeout=10000)
        except Exception:
            print(f"    [WARN] No product cards on page {page_num}")
            break

        html = await page.content()
        soup = BeautifulSoup(html, "lxml")

        cards = soup.select("li.product-card")
        added = 0
        for card in cards:
            a = card.find("a", href=lambda h: h and h.startswith("/us/p/"))
            if a:
                href = a["href"].rstrip("/")
                full_url = BASE_URL + href
                if full_url not in seen:
                    links.append(full_url)
                    seen.add(full_url)
                    added += 1

        print(f"    [Listing p{page_num}] +{added} new products (total {len(links)})")

        if max_products and len(links) >= max_products:
            break

        try:
            active_btn = await page.query_selector("div.pagination button.active")
            if not active_btn:
                break

            all_btns = await page.query_selector_all("div.pagination li button")
            active_idx = None
            for i, btn in enumerate(all_btns):
                cls = await btn.get_attribute("class") or ""
                if "active" in cls:
                    active_idx = i
                    break

            if active_idx is None or active_idx >= len(all_btns) - 1:
                break

            next_btn = all_btns[active_idx + 1]
            next_text = (await next_btn.inner_text()).strip()
            if not next_text.isdigit():
                break
            await next_btn.click()
            await page.wait_for_timeout(4000)
            page_num += 1
        except Exception as e:
            print(f"    [WARN] Pagination failed: {e}")
            break

    return links[:max_products] if max_products else links


# ── spec parsing helpers ──────────────────────────────────────────────────────
_SPEC_LABEL_MAP = {
    "height": "Height",
    "width": "Width",
    "length": "Length",
    "depth": "Depth",
    "diameter": "Diameter",
    "weight": "Weight",
    "canopy": "Canopy",
    "socket": "Socket",
    "wattage": "Wattage",
    "chain length": "Chain Length",
    "shade details": "Shade Details",
    "shade": "Shade Details",
    "fixture height": "Fixture Height",
    "o/a height": "O/A Height",
    "overall height": "Overall Height",
    "min. custom height": "Min. Custom Height",
    "minimum height": "Min. Custom Height",
    "extension": "Extension",
    "backplate": "Backplate",
    "base": "Base",
    "finish": "Finish",
    "rating": "Rating",
    "lightsource": "Lightsource",
    "light source": "Lightsource",
    "bulb type": "Lightsource",
    "color temperature": "Color Temperature",
    "lumens": "Lumens",
    "cri": "CRI",
    "designer": "Designer",
    "collection": "Collection",
    "series": "Collection",
}


def _clean_dim(value: str) -> str:
    """Strip inch marks and trailing whitespace from dimension values."""
    return re.sub(r'["\u201c\u201d]', "", value).strip()


def _parse_specs_from_text(spec_text: str) -> tuple[dict, list[str]]:
    """
    Parse label: value pairs from a spec block text.
    Returns (data_dict, spec_parts_list).
    """
    data: dict       = {}
    spec_parts: list = []

    for line in spec_text.splitlines():
        line = line.strip()
        if ":" not in line:
            continue
        key, _, val = line.partition(":")
        key = key.strip()
        val = val.strip()
        if not key or not val:
            continue

        canonical = _SPEC_LABEL_MAP.get(key.lower())
        if canonical:
            clean_val = _clean_dim(val)
            data[canonical] = clean_val
            spec_parts.append(f"{canonical}: {clean_val}")
        else:
            data[key] = val
            spec_parts.append(f"{key}: {val}")

    return data, spec_parts


# ── variant extraction ────────────────────────────────────────────────────────
def _extract_finish_from_offer_name(offer_name: str) -> str | None:
    """
    Visual Comfort offer names look like:
      "Talia Large Chandelier in Burnished Silver Leaf and Clear Swirled Glass"
    Extract the finish part after " in ".
    """
    if offer_name and " in " in offer_name:
        return offer_name.split(" in ", 1)[1].strip()
    return None


def _variants_from_jsonld(soup: BeautifulSoup, base_data: dict) -> list[dict]:
    """
    Parse the JSON-LD Product block and return one dict per offer (variant).
    Returns [] if no multi-offer Product block is found.
    """
    variants: list[dict] = []

    for script in soup.find_all("script", type="application/ld+json"):
        try:
            d = json.loads(script.string or "{}")
            candidates = d if isinstance(d, list) else [d]
            for obj in candidates:
                if obj.get("@type") != "Product":
                    continue

                # Base product name (shared across all variants)
                name = clean_text(obj.get("name", ""))
                if name and not base_data.get("Product Name"):
                    base_data["Product Name"] = name

                offers = obj.get("offers", [])
                if isinstance(offers, dict):
                    offers = [offers]

                for offer in offers:
                    row = dict(base_data)

                    # Variant-specific: SKU
                    sku = str(offer.get("sku") or offer.get("productID") or "").strip()
                    if sku:
                        row["SKU"] = sku

                    # Variant-specific: Price
                    price = offer.get("price") or offer.get("lowPrice")
                    if price is not None:
                        parsed_price = clean_price(str(price))
                        if parsed_price is not None:
                            row["Price"] = parsed_price

                    # Variant-specific: Image URL
                    img = offer.get("image", "")
                    if img:
                        row["Image URL"] = img
                    elif not row.get("Image URL"):
                        row["Image URL"] = base_data.get("Image URL", "")

                    # Variant-specific: Finish (from offer name "Product in Finish Name")
                    offer_name = clean_text(offer.get("name", ""))
                    finish = _extract_finish_from_offer_name(offer_name)
                    if finish:
                        row["Finish"] = finish

                    variants.append(row)

                if variants:
                    return variants
        except Exception:
            pass

    return variants


async def _variants_from_ui_swatches(page, base_data: dict) -> list[dict]:
    """
    Fallback: click each finish/color swatch on the page and capture the
    resulting SKU, price, image, and finish name.

    Tries multiple selector patterns used by visualcomfort.com.
    """
    variants: list[dict] = []

    # Possible swatch container selectors (try in order)
    swatch_selectors = [
        ".product-options-wrapper .swatch-option",
        "[data-role='swatch-options'] .swatch-option",
        ".swatch-attribute .swatch-option",
        "[class*='finish'] button",
        "[class*='swatch'] button",
        "[class*='variant'] button",
        "[class*='color'] button",
    ]

    swatch_els = []
    for sel in swatch_selectors:
        els = await page.query_selector_all(sel)
        if els:
            swatch_els = els
            print(f"    [Variants] Found {len(els)} swatches via '{sel}'")
            break

    if not swatch_els:
        print("    [Variants] No swatch elements found — returning single row")
        return []

    for swatch in swatch_els:
        try:
            # Get the finish label from the swatch
            finish_label = (
                await swatch.get_attribute("data-option-label")
                or await swatch.get_attribute("title")
                or await swatch.get_attribute("aria-label")
                or clean_text(await swatch.inner_text())
            )

            await swatch.click()
            await page.wait_for_timeout(2000)

            html = await page.content()
            soup = BeautifulSoup(html, "lxml")

            row = dict(base_data)

            # Capture updated SKU
            sku_el = soup.select_one(
                ".product-sku, [class*='sku'], [itemprop='sku']"
            )
            if sku_el:
                sku_text = re.sub(
                    r'^(sku|item\s*#|model)\s*[:#]?\s*', '',
                    clean_text(sku_el.get_text()), flags=re.I
                ).strip()
                if sku_text:
                    row["SKU"] = sku_text

            # Capture updated price
            price_el = soup.select_one(
                "[class*='price'] .price, .price-box .price, [itemprop='price']"
            )
            if price_el:
                price_val = clean_price(clean_text(price_el.get_text()))
                if price_val is not None:
                    row["Price"] = price_val

            # Capture updated image — prefer high-res zoom source
            img_el = soup.select_one(
                ".gallery-placeholder__image, .fotorama__img, "
                ".product-image-photo, [class*='product-image'] img"
            )
            if img_el:
                src = (
                    img_el.get("data-zoom-image")
                    or img_el.get("data-src")
                    or img_el.get("src", "")
                )
                if src and not src.startswith("data:"):
                    row["Image URL"] = src

            if finish_label:
                row["Finish"] = finish_label

            variants.append(row)

        except Exception as e:
            print(f"    [WARN] Swatch click failed: {e}")

    return variants


# ── base product data extraction ──────────────────────────────────────────────
def _extract_base_data(soup: BeautifulSoup, url: str) -> dict:
    """
    Extract all fields that are SHARED across every variant of a product:
    name, description, specs/dimensions, designer, collection, etc.
    Does NOT include variant-specific fields (SKU, price, finish, image).
    """
    data: dict = {"Source": url}

    # ── 1. Product Name from page heading ───────────────────────────────────
    for sel in ["h1.page-title span", "h1[itemprop='name']", "h1.product-name", "h1"]:
        el = soup.select_one(sel)
        if el:
            name = clean_text(el.get_text())
            if name:
                data["Product Name"] = name
                break

    # ── 2. Description ───────────────────────────────────────────────────────
    for sel in [
        ".product.attribute.description .value",
        "[class*='description'] .value",
        ".product-description",
        "[itemprop='description']",
        "[class*='description']",
    ]:
        el = soup.select_one(sel)
        if el:
            desc = clean_text(el.get_text())
            desc = re.sub(r'^DESCRIPTION\s*', '', desc).strip()
            if len(desc) > 30:
                data["Description"] = desc
                # Designer: "by [Name] for" or "by [Name],"
                m = re.search(
                    r'\bby\s+([A-Z][a-zA-Z\s\.]+?)(?:\s+for\b|\s+in\s+our|\s*\.\s|\s*,)',
                    desc
                )
                if m:
                    data.setdefault("Designer", m.group(1).strip())
                # Collection/series
                m2 = re.search(
                    r'\bthe\s+([A-Z][a-zA-Z\s]+?)\s+(?:series|collection)\b',
                    desc, re.I
                )
                if m2:
                    data.setdefault("Collection", m2.group(1).strip().title())
                break

    # ── 3. Spec section ──────────────────────────────────────────────────────
    spec_text = ""
    for sel in [
        ".product-attributes", ".product-specs",
        "[class*='specifications']", ".additional-attributes",
        "table.data.table.additional-attributes",
    ]:
        el = soup.select_one(sel)
        if el:
            spec_text = el.get_text(separator="\n")
            break

    if not spec_text:
        for el in soup.find_all(["div", "section", "ul"], limit=200):
            t = el.get_text()
            if "Height:" in t and ("Width:" in t or "Diameter:" in t or "Weight:" in t):
                if 50 < len(t) < 3000:
                    spec_text = t
                    break

    if spec_text:
        parsed, spec_parts = _parse_specs_from_text(spec_text)
        # Specs are shared — they don't vary by finish
        for k, v in parsed.items():
            # Don't overwrite Finish here; it's variant-specific
            if k.lower() != "finish":
                data.setdefault(k, v)
        if spec_parts:
            data["Specifications"] = " | ".join(spec_parts)

    # ── 4. Normalise dimension fields ────────────────────────────────────────
    for dim_key in ["Height", "Width", "Length", "Depth", "Diameter"]:
        raw = data.get(dim_key, "")
        if raw:
            clean = re.sub(r'["\u201c\u201d]', "", str(raw)).strip()
            m = re.match(r'^(\d+(?:\.\d+)?)', clean)
            if m:
                data[dim_key] = m.group(1)

    # ── 5. Weight — numeric only ─────────────────────────────────────────────
    if data.get("Weight"):
        m = re.match(r'^(\d+(?:\.\d+)?)', str(data["Weight"]).replace(",", ""))
        if m:
            data["Weight"] = m.group(1)

    # ── 6. Tearsheet ─────────────────────────────────────────────────────────
    ts = soup.find("a", string=re.compile(r"tear\s*sheet|spec\s*sheet", re.I))
    if not ts:
        ts = soup.find("a", href=re.compile(r"tearsheet|spec.sheet", re.I))
    if ts:
        href = ts.get("href", "")
        data["Tearsheet Link"] = (BASE_URL + href) if href.startswith("/") else href

    return data


# ── product detail scraper — returns one dict per variant ────────────────────
async def scrape_product(page, url: str) -> list[dict]:
    """
    Navigate to a product detail page and return one dict per variant
    (finish / size / color combination).

    Strategy:
      1. Extract base/shared data from the page (name, desc, specs, dims).
      2. Try JSON-LD offers array — one offer = one variant row.
         This is fast (no clicking) and covers 99% of VC products.
      3. If JSON-LD has only 1 offer (or none), fall back to clicking
         UI swatch buttons to enumerate each finish variant.
      4. If UI swatches also yield nothing, return a single-row list
         with whatever data was collected.
    """
    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
    except Exception as e:
        print(f"    [WARN] Failed to load: {url} — {e}")
        return [{"Source": url}]

    html = await page.content()
    soup = BeautifulSoup(html, "lxml")

    # Step 1: shared/base fields
    base = _extract_base_data(soup, url)

    # Step 2: variants from JSON-LD offers
    variants = _variants_from_jsonld(soup, base)

    if len(variants) > 1:
        print(f"    [Variants] {len(variants)} variants from JSON-LD")
        return variants

    # Step 3: fall back to UI swatch clicking
    print("    [Variants] JSON-LD single/no offer — trying UI swatches")
    variants = await _variants_from_ui_swatches(page, base)

    if variants:
        print(f"    [Variants] {len(variants)} variants from UI swatches")
        return variants

    # Step 4: single-row fallback — use the one JSON-LD offer if present
    if len(variants) == 0 and base:
        # Try to at least get SKU/price/image from the single JSON-LD offer
        single_variants = _variants_from_jsonld(soup, base)
        if single_variants:
            print("    [Variants] 1 variant (no finish selector found)")
            return single_variants

    print("    [Variants] No variants found — returning base data only")
    return [base]


# ── category scraper ──────────────────────────────────────────────────────────
async def scrape_category(
    page, writer: ExcelWriter, cat: dict, max_products: int | None = None
):
    name           = cat["name"]
    links          = cat["links"]
    studio_columns = cat.get("studio_columns", [])
    primary_link   = links[0] if links else ""

    mode_tag = f" [TEST: max {max_products} products]" if max_products else ""
    print(f"\n[Category] {name}{mode_tag}")
    writer.add_sheet(name, primary_link, studio_columns=studio_columns)

    all_urls: list[str] = []
    seen_urls: set[str] = set()
    for link in links:
        page_urls = await get_product_links(page, link, max_products=max_products)
        for u in page_urls:
            if u not in seen_urls:
                all_urls.append(u)
                seen_urls.add(u)
        if max_products and len(all_urls) >= max_products:
            break

    if max_products:
        all_urls = all_urls[:max_products]

    print(f"  [Category] {len(all_urls)} product pages to scrape")

    vendor_name_env = os.environ.get("VENDOR_NAME", VENDOR_NAME)
    global_idx = 1   # continuous index across all variant rows in this category

    for i, product_url in enumerate(all_urls, 1):
        print(f"  [{i}/{len(all_urls)}] {product_url}")
        try:
            variant_rows = await scrape_product(page, product_url)

            for variant in variant_rows:
                # Ensure Product Family Id is set
                if not variant.get("Product Family Id") and variant.get("Product Name"):
                    variant["Product Family Id"] = extract_family_id(
                        variant["Product Name"]
                    )

                # Generate SKU if missing
                if not variant.get("SKU"):
                    variant["SKU"] = generate_sku(vendor_name_env, name, global_idx)
                    print(f"    [SKU generated] {variant['SKU']}")

                writer.write_row(variant, category_name=name)
                global_idx += 1

            print(f"    → {len(variant_rows)} variant row(s) written")

        except Exception as e:
            print(f"  [ERROR] {product_url}: {e}")

        await async_polite_delay(1.0, 2.5)

    print(f"  [Category] Done — {global_idx - 1} total rows buffered")


# ── main ──────────────────────────────────────────────────────────────────────
async def main():
    info_path = Path(__file__).parent / "vendor_info.json"
    if info_path.exists():
        vendor_info = json.loads(info_path.read_text(encoding="utf-8"))
    else:
        sys.path.insert(0, str(PROJECT_ROOT))
        from vendor_parser import parse_vendor
        vendor_info = parse_vendor(VENDOR_NAME)

    categories = vendor_info["categories"]

    max_products: int | None = None
    if TEST_MODE:
        categories = [c for c in categories if c.get("links")][:TEST_MAX_CATEGORIES]
        max_products = TEST_MAX_PRODUCTS

    print(f"\n[Scraper] Vendor  : {vendor_info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")
    print(f"[Scraper] Headless: {HEADLESS}")
    print(f"[Scraper] Cats    : {len(categories)}"
          + (f" (capped at {TEST_MAX_CATEGORIES})" if TEST_MODE else ""))
    if TEST_MODE:
        print(f"[Scraper] Max products/cat: {TEST_MAX_PRODUCTS}")

    writer = ExcelWriter(OUTPUT_PATH, vendor_info["vendor_name"])

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        page.set_default_timeout(TIMEOUT_MS)

        for cat in categories:
            if not cat.get("links"):
                print(f"[Skip] {cat['name']} — no links")
                continue
            await scrape_category(page, writer, cat, max_products=max_products)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
