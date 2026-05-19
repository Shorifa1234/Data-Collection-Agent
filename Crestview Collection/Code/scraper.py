import asyncio
import json
import os
import re
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Crestview Collection")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.crestviewcollection.com"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def sku_from_url(url: str) -> str:
    """Extract SKU from product URL slug (last hyphen-segment, uppercased).
    e.g. .../virentia-cocktail-table-cvfnr5335 → CVFNR5335
    """
    slug = url.rstrip("/").split("/")[-1].split("?")[0]
    parts = slug.split("-")
    if parts:
        candidate = parts[-1].upper()
        if re.match(r'^CV[A-Z0-9]+$', candidate):
            return candidate
    return ""


def strip_inches(val: str) -> str:
    """Remove inch marks and 'in' suffix: '18"' → '18'."""
    return re.sub(r'["\']', '', val).replace(" in", "").strip()


def clean_dim_summary(val: str) -> str:
    """'36 x 36 x 18 (in)' → '36 x 36 x 18'."""
    return re.sub(r'\s*\(in\)\s*', '', val).strip()


# ---------------------------------------------------------------------------
# GraphQL response parser — uses direct cv_* fields (NOT custom_attributes)
# ---------------------------------------------------------------------------

def _label_val(v) -> str:
    """Extract string value from scalar, list, or label dict."""
    if v is None:
        return ""
    if isinstance(v, list):
        return ", ".join(str(x) for x in v if x is not None)
    return str(v)


def _parse_gql_item(item: dict) -> dict:
    """Parse Crestview's Magento 2 GraphQL product item.

    All meaningful data is in direct cv_* named fields.
    custom_attributes are all null for this site.
    """
    data = {}

    # Core identity
    name = (item.get("name") or "").strip()
    if name:
        data["Product Name"] = sentence_case(name)

    sku = (item.get("sku") or "").strip().upper()
    if sku:
        data["SKU"] = sku

    # Description (strip HTML)
    desc_html = (item.get("description") or {}).get("html", "")
    if desc_html:
        data["Description"] = clean_text(re.sub(r'<[^>]+>', ' ', desc_html))

    # Image — prefer media_gallery position-1 (higher res); fall back to small_image
    gallery = sorted(
        [g for g in (item.get("media_gallery") or [])
         if g.get("url") and "/catalog/category/" not in g.get("url", "")],
        key=lambda g: g.get("position") or 99,
    )
    if gallery:
        data["Image URL"] = gallery[0]["url"]
    else:
        sm_img = (item.get("small_image") or {}).get("url", "")
        if sm_img and "/catalog/category/" not in sm_img:
            data["Image URL"] = sm_img

    # Price — available via older price field (not price_range on this site)
    try:
        price_val = (
            item.get("price", {})
                .get("regularPrice", {})
                .get("amount", {})
                .get("value")
        )
        if price_val and float(price_val) > 0:
            data["Price"] = float(price_val)
    except Exception:
        pass

    # cv_* direct fields → human-readable field names
    def _set(field: str, raw):
        v = _label_val(raw)
        if v and v not in ("0", "None", ""):
            data[field] = v

    _set("Collection",  item.get("cv_collection_label"))
    _set("Styles",      item.get("cv_style_label"))
    _set("Type",        item.get("cv_product_category_label"))
    _set("Origin",      item.get("country_of_manufacture"))
    _set("UPC Code",    item.get("cv_upc"))
    _set("Finish",      item.get("cv_primary_finish"))
    _set("Stock Status", item.get("stock_status"))

    # Dimensions — numeric, convert to string without units
    def _dim(field: str, raw):
        if raw is not None:
            try:
                v = float(raw)
                # Store as integer string if whole number, else float string
                data[field] = str(int(v)) if v == int(v) else str(v)
            except Exception:
                pass

    _dim("Height", item.get("cv_furniture_height"))
    _dim("Width",  item.get("cv_furniture_width"))
    _dim("Length", item.get("cv_furniture_length"))
    _dim("Depth",  item.get("cv_furniture_depth"))
    _dim("Weight", item.get("cv_furniture_weight"))

    # Construct Dimensions summary from individual values
    h = item.get("cv_furniture_height")
    w = item.get("cv_furniture_width")
    l = item.get("cv_furniture_length")
    d = item.get("cv_furniture_depth")
    dim_parts = []
    if l: dim_parts.append(f"{l}")
    if w: dim_parts.append(f"{w}")
    if d: dim_parts.append(f"{d}")
    if h: dim_parts.append(f"{h}")
    if dim_parts:
        data["Dimensions"] = " x ".join(str(int(float(p)) if float(p) == int(float(p)) else float(p)) for p in dim_parts)

    # Lead time / ETA fields
    eta1 = item.get("cv_in_transit_eta_1")
    eta2 = item.get("cv_future_eta_1")
    if eta1:
        data["Lead Time"] = _label_val(eta1)
    elif eta2:
        data["Lead Time"] = _label_val(eta2)

    return data


# ---------------------------------------------------------------------------
# Direct GraphQL fetch — bypasses React rendering entirely
# ---------------------------------------------------------------------------

_GQL_QUERY = """
{
  products(filter: {sku: {eq: "%s"}}) {
    items {
      name sku
      description { html }
      small_image { url }
      media_gallery { url position }
      price { regularPrice { amount { value } } }
      stock_status
      country_of_manufacture
      cv_collection_label
      cv_style_label
      cv_product_category_label
      cv_upc
      cv_primary_finish
      cv_furniture_height
      cv_furniture_width
      cv_furniture_length
      cv_furniture_depth
      cv_furniture_weight
      cv_in_transit_eta_1
      cv_future_eta_1
    }
  }
}
"""


async def _direct_gql_product(page, sku: str) -> dict:
    """POST to /graphql directly — works even when React crashes on the page."""
    query = _GQL_QUERY % sku
    try:
        result = await page.evaluate(
            """
            async (query) => {
                try {
                    const r = await fetch('/graphql', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'Accept': 'application/json'
                        },
                        body: JSON.stringify({ query })
                    });
                    if (!r.ok) return null;
                    return await r.json();
                } catch (e) {
                    return { _fetchError: e.message };
                }
            }
            """,
            query,
        )
        if not result or result.get("_fetchError"):
            return {}
        items = (result.get("data") or {}).get("products", {}).get("items") or []
        if items:
            return _parse_gql_item(items[0])
    except Exception as exc:
        print(f"    Direct GQL error for {sku}: {exc}")
    return {}


# ---------------------------------------------------------------------------
# DOM body-text parser (fallback)
# ---------------------------------------------------------------------------

_LABEL_MAP = {
    "SKU":               "SKU",
    "Collection":        "Collection",
    "Dimensions":        "Dimensions",
    "Material":          "Material",
    "Finish":            "Finish",
    "Color":             "Color",
    "UPC Code":          "UPC Code",
    "Weight":            "Weight",
    "Item Height":       "Height",
    "Item Length":       "Length",
    "Item Width":        "Width",
    "Item Depth":        "Depth",
    "Item Diameter":     "Diameter",
    "Type":              "Type",
    "Styles":            "Styles",
    "Socket":            "Socket",
    "Rating":            "Rating",
    "Color Temperature": "Color Temperature",
    "Wattage":           "Wattage",
    "Bulb Type":         "Bulb Type",
    "Designer":          "Designer",
    "Seat Height":       "Seat Height",
    "Seat Depth":        "Seat Depth",
    "Seat Width":        "Seat Width",
    "Arm Height":        "Arm Height",
}

_DIM_FIELDS = {"Height", "Width", "Length", "Depth", "Diameter"}


def _parse_body_text(body: str, data: dict):
    """Extract label-value pairs from rendered page body text."""
    lines = [l.strip() for l in body.split("\n") if l.strip()]

    for i, line in enumerate(lines):
        field = _LABEL_MAP.get(line)
        if not field:
            continue
        if i + 1 >= len(lines):
            continue
        value = lines[i + 1].strip()
        # Skip if value is another label, a nav item, or login prompt
        if (value in _LABEL_MAP
                or value.upper().startswith("LOGIN")
                or len(value) > 200):
            continue
        if not value:
            continue

        # Clean dimension values
        if field in _DIM_FIELDS:
            value = strip_inches(value)
        elif field == "Dimensions":
            value = clean_dim_summary(value)

        # Only set if not already captured (GraphQL takes precedence)
        if field not in data or not data[field]:
            data[field] = value


# ---------------------------------------------------------------------------
# Product scraper
# ---------------------------------------------------------------------------

async def scrape_product(page, url: str) -> list[dict]:
    """Return a list with one dict per product (no variant splitting on this site)."""
    data = {"Source URL": url}

    url_sku = sku_from_url(url)
    if url_sku:
        data["SKU"] = url_sku

    # Navigate to product page — we don't need JS to render, just domcontentloaded
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    # Brief wait so the browser context is ready for a same-origin fetch
    await page.wait_for_timeout(800)

    # --- Primary: direct GraphQL POST (bypasses React crash entirely) --------
    if url_sku:
        gql_result = await _direct_gql_product(page, url_sku)
        if gql_result:
            data.update(gql_result)

    # --- Image fallbacks (in order of quality) ------------------------------
    # 1. og:image meta tag (often a high-res version)
    if not data.get("Image URL"):
        try:
            og = await page.get_attribute('meta[property="og:image"]', "content")
            if og and "/catalog/category/" not in og:
                data["Image URL"] = og
        except Exception:
            pass

    # 2. First catalog/product image in DOM
    if not data.get("Image URL"):
        try:
            img = await page.query_selector('img[src*="media/catalog/product"]')
            if img:
                src = (await img.get_attribute("src") or
                       await img.get_attribute("data-src") or "")
                if src:
                    data["Image URL"] = src
        except Exception:
            pass

    # --- Product name fallback from h1 (if React actually rendered) ---------
    if not data.get("Product Name"):
        try:
            h1 = await page.query_selector("h1")
            if h1:
                text = clean_text(await h1.inner_text())
                if text:
                    data["Product Name"] = sentence_case(text)
        except Exception:
            pass

    # --- Body text for any remaining fields (Material, etc.) ----------------
    try:
        body_text = await page.inner_text("body")
        _parse_body_text(body_text, data)
    except Exception:
        pass

    # --- Post-processing ---------------------------------------------------
    if data.get("Dimensions"):
        data["Dimensions"] = clean_dim_summary(str(data["Dimensions"]))

    if data.get("Product Name") and not data.get("Product Family Id"):
        data["Product Family Id"] = extract_family_id(data["Product Name"])

    return [data]


# ---------------------------------------------------------------------------
# Listing page crawler
# ---------------------------------------------------------------------------

async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from a listing (all pages, deduplicated)."""
    # Strip any existing ?page= param to get the base URL
    base_url = re.sub(r'\?page=\d+', '', listing_url).rstrip('?').rstrip('&')

    all_links: list[str] = []
    seen: set[str] = set()
    page_num = 1

    while True:
        url = f"{base_url}?page={page_num}"
        print(f"    Listing p{page_num}: {url}")

        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        # Wait for product detail links (slugs ending in -cv...) to appear in DOM
        try:
            await page.wait_for_selector('a[href*="-cv"]', timeout=15_000)
        except Exception:
            print(f"    p{page_num}: no product links — stopping")
            break
        await page.wait_for_timeout(1000)  # brief extra settle

        hrefs: list[str] = await page.eval_on_selector_all(
            "a[href]",
            "els => [...new Set(els.map(e => e.href))]"
        )

        # Keep only product detail pages: URL slug ends with -CV<alphanum>
        new_links = []
        for h in hrefs:
            clean = h.split("?")[0].rstrip("/")
            if re.search(r'-cv[a-z0-9]+$', clean.lower()) and clean not in seen:
                seen.add(clean)
                new_links.append(clean)

        if not new_links:
            print(f"    p{page_num}: no new products — end of pagination")
            break

        all_links.extend(new_links)
        print(f"    p{page_num}: +{len(new_links)} (total {len(all_links)})")

        # Check whether a next-page control exists
        has_next: bool = await page.evaluate("""
        () => {
            // Numbered pagination: look for a link/button after the active page
            const els = Array.from(document.querySelectorAll('a, button'));
            return els.some(el => {
                const txt  = (el.textContent || '').trim();
                const aria = (el.getAttribute('aria-label') || '').toLowerCase();
                return txt === '>' || txt === '›' || txt === 'Next'
                    || aria.includes('next') || aria.includes('page suivante');
            });
        }
        """)

        if not has_next:
            print(f"    p{page_num}: no next button — end of pagination")
            break

        page_num += 1
        await async_polite_delay()

    return all_links


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main():
    t0 = time.time()
    info   = json.loads(
        (Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        print(f"[TEST MODE: all {len(categories)} categories, "
              f"max {TEST_MAX_PRODUCTS} products each]")
        if TEST_MAX_CATEGORIES < len(categories):
            categories = categories[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        # Extra stealth patches — this site's React bundle checks for browser APIs
        # that headless Chromium doesn't expose by default, causing a JS crash.
        await page.context.add_init_script("""
            if (!window.chrome) { window.chrome = { runtime: {} }; }
            if (!window.matchMedia) {
                window.matchMedia = q => ({
                    matches: false, media: q, onchange: null,
                    addListener: () => {}, removeListener: () => {},
                    addEventListener: () => {}, removeEventListener: () => {},
                    dispatchEvent: () => false
                });
            }
            Object.defineProperty(navigator, 'plugins',
                { get: () => [1, 2, 3, 4, 5] });
            Object.defineProperty(navigator, 'languages',
                { get: () => ['en-US', 'en'] });
        """)

        for cat in categories:
            if not cat["links"]:
                print(f"\nSkipping {cat['name']} — no links")
                continue

            cat_name = cat["name"]
            print(f"\n{'='*60}")
            print(f"[{cat['group']}] {cat_name}  ({len(cat['links'])} link(s))")

            writer.add_sheet(
                cat_name,
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            # Collect URLs from ALL listing links for this category
            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            print(f"  Found {len(all_product_urls)} products")

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]
                print(f"  [TEST] Capped at {len(all_product_urls)}")

            global_idx = 1
            for prod_url in all_product_urls:
                try:
                    rows = await scrape_product(page, prod_url)
                    for row in rows:
                        row["Manufacturer"] = info["vendor_name"]
                        row["Category"]     = cat_name
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(
                                info["vendor_name"], cat_name, global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(
                                row["Product Name"])
                        writer.write_row(row, category_name=cat_name)
                        global_idx += 1
                        print(f"    [{global_idx-1}] {row.get('Product Name','?')} "
                              f"| SKU: {row.get('SKU','?')} "
                              f"| Img: {'Y' if row.get('Image URL') else 'N'}")
                except Exception as exc:
                    print(f"  ERROR {prod_url}: {exc}")

                await async_polite_delay()

    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")
    print(f"Total time: {time.time()-t0:.1f}s")


if __name__ == "__main__":
    asyncio.run(main())
