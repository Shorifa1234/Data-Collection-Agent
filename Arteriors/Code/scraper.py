"""
scraper.py  -  Arteriors
------------------------
Scrapes ALL product data from arteriorshome.com (Magento 2 / Klevu) for every
category defined in the SD tracker, with fully dynamic column output.

Products are rendered by Klevu JS — static HTML contains no product cards.
Pagination is handled by clicking Klevu's .klevuPaginate buttons.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Arteriors"           # full run
    python orchestrator.py "Arteriors" --test    # test run

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

# ── project root so base_scraper is importable ──────────────────────────────
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
VENDOR_NAME = os.environ.get("VENDOR_NAME", "Arteriors")
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

BASE_URL   = "https://www.arteriorshome.com"
TIMEOUT_MS = 45_000

# ── listing helpers ───────────────────────────────────────────────────────────
async def _accept_cookies(page) -> None:
    """Dismiss the cookie/consent banner if present."""
    try:
        for selector in [
            "button:has-text('Accept All')",
            "button:has-text('Accept')",
            "[id*='cookie'] button",
            ".cookie-notice button",
        ]:
            btn = await page.query_selector(selector)
            if btn:
                await btn.click()
                await page.wait_for_timeout(800)
                break
    except Exception:
        pass


async def _extract_links_from_page(page, seen: set[str], listing_base: str) -> list[str]:
    """
    Parse the current rendered page for product detail links.
    Appends found hrefs to `seen` in-place; returns list of NEW hrefs found.
    """
    html = await page.content()
    soup = BeautifulSoup(html, "lxml")
    new_hrefs: list[str] = []

    # Strategy 1: Hyva/Magento 2 canonical product link class
    for a in soup.select("a.product-item-link"):
        href = a.get("href", "").strip().rstrip("/")
        if href and href not in seen:
            new_hrefs.append(href)
            seen.add(href)

    # Strategy 2: first <a> inside each product-item li/div
    if not new_hrefs:
        for item in soup.select("li.product-item, li.item.product, div.product-item"):
            for a in item.find_all("a", href=True):
                href = a["href"].strip().rstrip("/")
                if (
                    href
                    and href not in seen
                    and href.startswith(BASE_URL)
                    and href.rstrip("/") != listing_base.rstrip("/")
                ):
                    new_hrefs.append(href)
                    seen.add(href)
                    break   # one URL per card

    # Strategy 3: URL-pattern fallback — any link deeper than listing URL
    if not new_hrefs:
        base_depth = listing_base.rstrip("/").count("/")
        for a in soup.find_all("a", href=True):
            href = a["href"].strip().rstrip("/")
            if (
                href
                and href not in seen
                and href.startswith(BASE_URL)
                and "?" not in href
                and href.count("/") > base_depth
            ):
                new_hrefs.append(href)
                seen.add(href)

    return new_hrefs


async def get_product_links(page, listing_url: str, max_products: int | None = None) -> list[str]:
    """
    Collect all product URLs from an Arteriors category page.
    The site uses Hyva (Magento 2 + Alpine.js + Tailwind).
    Pagination is URL-based: ?p=2, ?p=3 …
    """
    links: list[str] = []
    seen:  set[str]  = set()

    base_url = listing_url.split("?")[0].rstrip("/")
    page_num = 1

    while True:
        paginated_url = listing_url if page_num == 1 else f"{base_url}?p={page_num}"
        print(f"    [Listing p{page_num}] {paginated_url}")

        await page.goto(paginated_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        if page_num == 1:
            await _accept_cookies(page)

        # Wait for at least one product item to appear
        try:
            await page.wait_for_selector(
                "a.product-item-link, li.product-item, li.item.product",
                timeout=20_000,
            )
        except Exception:
            print(f"    [WARN] No product items found on page {page_num} — stopping pagination")
            break

        await page.wait_for_timeout(1500)

        new_hrefs = await _extract_links_from_page(page, seen, base_url)
        links.extend(new_hrefs)
        print(f"    [Listing p{page_num}] +{len(new_hrefs)} new products (total {len(links)})")

        if not new_hrefs:
            print(f"    [Listing] No new links on page {page_num} — done")
            break

        if max_products and len(links) >= max_products:
            break

        # Check if a next-page link exists on the rendered page
        next_btn = await page.query_selector(
            "a.action.next, li.pages-item-next a, .pages .next, a[aria-label='Next']"
        )
        if not next_btn:
            break

        page_num += 1

    return links[:max_products] if max_products else links


# ── product detail scraper ────────────────────────────────────────────────────
async def scrape_product(page, url: str) -> list[dict]:
    """
    Navigate to a product detail page and collect every field available.
    Returns a list of flat dicts — one per variant (finish/size/color).
    If no selectable variants exist, returns a single-element list.
    """
    data: dict = {"Source": url, "Manufacturer": VENDOR_NAME}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await _accept_cookies(page)
        await page.wait_for_timeout(2500)
    except Exception as e:
        print(f"    [WARN] Failed to load: {url} — {e}")
        return [data]

    html = await page.content()
    soup = BeautifulSoup(html, "lxml")

    # ── 1. JSON-LD Product schema ────────────────────────────────────────────
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            d = json.loads(script.string or "{}")
            candidates = d.get("@graph", [d]) if isinstance(d, dict) else d
            if isinstance(candidates, dict):
                candidates = [candidates]
            for obj in candidates:
                if obj.get("@type") == "Product":
                    name = clean_text(obj.get("name", ""))
                    if name:
                        data["Product Name"] = name.upper()
                        data["Product Family Id"] = extract_family_id(data["Product Name"])
                    if obj.get("description"):
                        data["Description"] = clean_text(obj["description"])
                    if obj.get("sku"):
                        data["SKU"] = str(obj["sku"])
                    offers = obj.get("offers", {})
                    if isinstance(offers, list):
                        offers = offers[0]
                    price = offers.get("price") or offers.get("lowPrice")
                    if price is not None:
                        data["Price"] = clean_price(str(price))
                    break
        except Exception:
            pass

    # ── 2. Image URL — from product detail page gallery ──────────────────────
    for selector in [
        ".product-image-gallery .fotorama__img",
        ".gallery-placeholder img",
        ".product-media img",
        ".product-image-photo",
        "[class*='product-image'] img",
    ]:
        img_tag = soup.select_one(selector)
        if img_tag:
            src = (
                img_tag.get("data-zoom-image")
                or img_tag.get("data-src")
                or img_tag.get("src", "")
            )
            if src and not src.startswith("data:"):
                data["Image URL"] = src
                break

    # ── 2b. JSON-LD image fallback ────────────────────────────────────────────
    if not data.get("Image URL"):
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                d = json.loads(script.string or "{}")
                candidates = d.get("@graph", [d]) if isinstance(d, dict) else d
                if isinstance(candidates, dict):
                    candidates = [candidates]
                for obj in candidates:
                    if obj.get("@type") == "Product":
                        img = obj.get("image")
                        if img:
                            data["Image URL"] = img if isinstance(img, str) else img[0]
                        break
            except Exception:
                pass
            if data.get("Image URL"):
                break

    # ── 2. SKU fallback (h2.pro-sku) ────────────────────────────────────────
    if not data.get("SKU"):
        sku_el = soup.select_one("h2.pro-sku, .product-sku h2, .product-sku")
        if sku_el:
            data["SKU"] = clean_text(sku_el.get_text())

    # ── 3. Finish / colour variant ───────────────────────────────────────────
    # In .config_custom_options: the <ul> after #dimensions holds the finish name
    config = soup.find(class_="config_custom_options")
    if config:
        dim_div = config.find(id="dimensions")
        if dim_div:
            finish_ul = dim_div.find_next_sibling("ul")
            if finish_ul:
                span = finish_ul.find("span", class_="data")
                if span:
                    finish_text = clean_text(span.get_text())
                    if finish_text:
                        data["Finish"] = finish_text

    # ── 4. Dimensions from span.product-metric ───────────────────────────────
    dim_w = soup.select_one("span.product-metric.dimension_w")
    dim_d = soup.select_one("span.product-metric.dimension_d")
    dim_h = soup.select_one("span.product-metric.dimension_h")
    dim_dia = soup.select_one("span.product-metric.dimension_dia, span.product-metric.dimension_diameter")

    def _extract_inches(span_el) -> str | None:
        """Extract the numeric value from data-in attribute, stripping ' in'."""
        if not span_el:
            return None
        raw = span_el.get("data-in", "") or span_el.get_text()
        # Remove label prefix like "Width: "
        raw = re.sub(r'^[^0-9]*', '', raw)
        # Remove " in" and " cm"
        raw = re.sub(r'\s*(in|cm).*$', '', raw, flags=re.I).strip()
        return raw if raw else None

    if dim_w or dim_d or dim_h or dim_dia:
        parts = []
        if dim_w:
            v = _extract_inches(dim_w)
            if v:
                data["Width"] = v
                parts.append(f"W {v}")
        if dim_d:
            v = _extract_inches(dim_d)
            if v:
                data["Depth"] = v
                parts.append(f"D {v}")
        if dim_h:
            v = _extract_inches(dim_h)
            if v:
                data["Height"] = v
                parts.append(f"H {v}")
        if dim_dia:
            v = _extract_inches(dim_dia)
            if v:
                data["Diameter"] = v
                parts.append(f"Dia {v}")
        if parts:
            data["Dimensions"] = " x ".join(parts)

    # ── 5. Weight ────────────────────────────────────────────────────────────
    weight_el = soup.select_one("span.product-metric-weight")
    if weight_el:
        raw_w = weight_el.get("data-in", "") or weight_el.get_text()
        # Strip " lbs" etc.
        raw_w = re.sub(r'\s*(lbs?|kg).*$', '', raw_w, flags=re.I).strip()
        if raw_w:
            data["Weight"] = raw_w

    # ── 6. Attributes table (.table-row with .label-attr pairs) ──────────────
    # Maps label text to canonical field names
    _ATTR_MAP = {
        "primary material": "Materials",
        "material": "Materials",
        "finish will vary": None,   # skip boolean flags
        "top coat/sealant": None,
        "environment suitability": None,
        "overall": None,            # dimension (duplicate)
        "contract suitability": None,
        "shipping requirements": None,
        "weight": None,             # already captured above
    }

    spec_parts: list[str] = []
    for row in soup.select("div.table-row"):
        label_el = row.select_one(".label-attr")
        val_el   = row.select_one(".table-cell:not(.label-attr)")
        if not label_el or not val_el:
            continue
        label = clean_text(label_el.get_text())
        value = clean_text(val_el.get_text())
        if not label or not value:
            continue

        canonical = _ATTR_MAP.get(label.lower())
        if canonical is None and label.lower() in _ATTR_MAP:
            continue    # explicitly skipped
        if canonical:
            data.setdefault(canonical, value)
        else:
            # Store as-is using the label as key (title-case)
            key = label.title()
            data.setdefault(key, value)

        spec_parts.append(f"{label}: {value}")

    if spec_parts:
        data.setdefault("Specifications", " | ".join(spec_parts))

    # ── 7. Tearsheet link ────────────────────────────────────────────────────
    # Arteriors tearsheet is AJAX-generated; the SKU anchors it
    sku = data.get("SKU", "")
    if sku:
        data["Tearsheet Link"] = f"{BASE_URL}/tearsheet/download/?sku={sku}"

    return [data]


# ── category scraper ──────────────────────────────────────────────────────────
async def scrape_category(
    page,
    writer: ExcelWriter,
    cat: dict,
    max_products: int | None = None,
):
    name           = cat["name"]
    links          = cat["links"]
    studio_columns = cat.get("studio_columns", [])
    primary_link   = links[0] if links else ""

    mode_tag = f" [TEST: max {max_products} products]" if max_products else ""
    print(f"\n[Category] {name}{mode_tag}")
    writer.add_sheet(name, primary_link, studio_columns=studio_columns)

    # Collect product URLs across all category links
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

    print(f"  [Category] {len(all_urls)} products to scrape")

    vendor_name_env = os.environ.get("VENDOR_NAME", VENDOR_NAME)

    global_idx = 1
    for i, product_url in enumerate(all_urls, 1):
        print(f"  [{i}/{len(all_urls)}] {product_url}")
        try:
            variant_rows = await scrape_product(page, product_url)
            for variant in variant_rows:
                if not variant.get("SKU"):
                    variant["SKU"] = generate_sku(vendor_name_env, name, global_idx)
                    print(f"    [SKU generated] {variant['SKU']}")

                if not variant.get("Product Family Id") and variant.get("Product Name"):
                    variant["Product Family Id"] = extract_family_id(variant["Product Name"])

                writer.write_row(variant, category_name=name)
                global_idx += 1
        except Exception as e:
            print(f"  [ERROR] {product_url}: {e}")

        await async_polite_delay(1.0, 2.5)

    print(f"  [Category] Done — {len(all_urls)} rows buffered")


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
