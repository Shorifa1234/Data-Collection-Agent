import asyncio, json, math, os, re, sys, time
import requests
from pathlib import Path
from urllib.parse import parse_qs, urlparse

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    clean_text, clean_price, generate_sku, extract_family_id,
    async_polite_delay, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "A.R.T Furniture")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://arthomefurnishings.com"
API_BASE = "https://api.houseofmarkor.com/mibd"
IMAGE_CDN = "https://lemon.houseofmarkor.com/page/lemon/NEWreadFile/read/blockid/234/width/1200/height/9999999999/filename"

_SESSION = requests.Session()
_SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://arthomefurnishings.com/",
    "Content-Type": "application/json",
})
_COOKIE = "scraper_session_001"  # API accepts any non-empty value


# ── API helpers ───────────────────────────────────────────────────────────────

def _api_params() -> dict:
    return {
        "timestamp": int(time.time() * 1000),
        "obtain": "1",
        "cookie": _COOKIE,
        "lang_type": "en",
        "os": "Win10",
        "browser": "Chrome",
    }


def _get_detail(sku: str) -> dict:
    """Call the product detail API and return the val dict."""
    r = _SESSION.post(
        f"{API_BASE}/productDetail/detail",
        params=_api_params(),
        json={"sku_number": sku},
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    if data.get("err") != 200:
        return {}
    return data.get("val", {})


def _get_listing_products(room: str, category: str, page: int = 1, limit: int = 48) -> dict:
    """Call the product search API for one page."""
    r = _SESSION.post(
        f"{API_BASE}/art/product/search",
        params=_api_params(),
        json={
            "page": page,
            "limit": limit,
            "search": "",
            "Room": room,
            "Category": category,
            "type": "category",
            "type_filter": category,
        },
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    if data.get("err") != 200:
        return {"list": [], "total": 0}
    return data.get("val", {"list": [], "total": 0})


def _collect_all_skus(room: str, category: str) -> list[dict]:
    """Paginate through all pages and return every product entry."""
    first = _get_listing_products(room, category, page=1)
    total = first.get("total", 0)
    products = list(first.get("list", []))
    total_pages = math.ceil(total / 48)
    for pg in range(2, total_pages + 1):
        result = _get_listing_products(room, category, page=pg)
        products.extend(result.get("list", []))
        time.sleep(0.2)
    return products


def _parse_listing_url(url: str) -> tuple[str, str]:
    """Extract Room and Category from a listing URL's query string."""
    qs = parse_qs(urlparse(url).query)
    room = qs.get("Room", [""])[0]
    category = qs.get("Category", [""])[0]
    return room, category


# ── Dimension parser (Inches format: w-30.00 x d-18.00 x h-26.00) ────────────

def _parse_art_dimensions(dim_str: str) -> dict:
    """Parse 'w-30.00 x d-18.00 x h-26.00' into Width/Depth/Height numeric strings."""
    result = {}
    if not dim_str:
        return result
    for code, field in [("w", "Width"), ("d", "Depth"), ("h", "Height"),
                        ("l", "Length"), ("dia", "Diameter")]:
        m = re.search(rf'\b{code}-(\d+\.?\d*)', dim_str, re.IGNORECASE)
        if m:
            result[field] = m.group(1)
    result["Dimensions"] = dim_str.strip()
    return result


# ── Spec-item parser (from .specifications_item_title elements) ───────────────

def _parse_spec_items(items: list[str]) -> dict:
    """
    Parse the spec items from the product page.
    Items are raw text strings like:
      'SKU : 299140-2349'
      'Dimensions in Inches: w-30.00 x d-18.00 x h-26.00'
      'Weight: 80.4687249 lbs'
      '- Style: Contemporary'
      '- Material: Parawood Solids, ...'
      '- Dovetail drawer construction...'   ← feature bullet
    """
    data: dict = {}
    feature_bullets: list[str] = []

    for raw in items:
        text = raw.strip().lstrip("- ").strip()

        # Skip the SKU line (already from API)
        if text.lower().startswith("sku"):
            continue

        # Dimensions in Inches — primary source for W/D/H
        if text.lower().startswith("dimensions in inches"):
            dim_val = text.split(":", 1)[-1].strip()
            dims = _parse_art_dimensions(dim_val)
            data.update(dims)
            continue

        # Dimensions in Centimeters — skip (we use inches)
        if text.lower().startswith("dimensions in centimeters"):
            continue

        # Weight: 80.4687249 lbs
        if text.lower().startswith("weight:"):
            wt_raw = text.split(":", 1)[-1].strip()
            wt_num = re.search(r"([\d.]+)", wt_raw)
            if wt_num:
                data["Weight"] = wt_num.group(1)
            continue

        # Key: Value lines like "Style: Contemporary" or "Finish(es): Linen"
        kv = re.match(r'^([A-Za-z][A-Za-z &/()-]{1,40}):\s*(.+)$', text)
        if kv:
            key = kv.group(1).strip().title()
            val = kv.group(2).strip()
            # Normalise plural/variant spellings → canonical column names
            key = re.sub(r'\(e?s\)$', '', key, flags=re.IGNORECASE).strip()
            key_map = {
                "Style": "Style",
                "Finish": "Finish",
                "Material": "Material",
                "Furniture Piece": "Furniture Piece",
                "Seating": "Seating Type",
            }
            mapped = key_map.get(key, key)
            data[mapped] = val
            continue

        # Everything else is a feature bullet point
        if text:
            feature_bullets.append(text)

    if feature_bullets:
        data["Features"] = " | ".join(feature_bullets)

    return data


# ── Product page scraper (Playwright) ────────────────────────────────────────

async def scrape_product(page, sku: str, listing_data: dict) -> dict:
    """
    Navigate to the product detail page, intercept the detail API,
    and extract all fields from the rendered HTML.
    Returns a flat dict with all available data.
    """
    url = f"{BASE_URL}/page/products/product-{sku}.html"

    # Fetch detail API directly via requests (no Playwright interception needed)
    detail_api: dict = await asyncio.to_thread(_get_detail, sku)

    # Navigate the product page for the rendered spec items (Features, Style, Finish, Material, Weight)
    await page.goto(url, timeout=60_000, wait_until="domcontentloaded")
    try:
        await page.wait_for_selector(".specifications_item", timeout=20_000)
    except Exception:
        await page.wait_for_timeout(5000)

    row: dict = {
        "Manufacturer": VENDOR_NAME,
        "Source": url,
    }

    # ── Image URL (Salsify CDN from detail API, fallback to internal CDN) ────
    image_url = None
    if detail_api.get("picture_all"):
        # Prefer first "silos" type (clean product shot)
        for pic in detail_api["picture_all"]:
            if pic.get("type") == "silos" and pic.get("url"):
                image_url = pic["url"]
                break
        if not image_url:
            image_url = detail_api["picture_all"][0].get("url")
    if not image_url and listing_data.get("image"):
        image_url = f"{IMAGE_CDN}/{listing_data['image']}"
    if image_url:
        row["Image URL"] = image_url

    # ── Core fields ────────────────────────────────────────────────────────
    name = (detail_api.get("name") or listing_data.get("name") or "").strip()
    row["Product Name"] = name
    row["Product Family Id"] = extract_family_id(name)
    row["SKU"] = sku

    price_raw = listing_data.get("msrp_price", "")
    if price_raw and str(price_raw) not in ("0", ""):
        row["Price"] = clean_price(str(price_raw))

    desc = clean_text(detail_api.get("description") or "")
    if desc:
        row["Description"] = desc

    collection = (detail_api.get("collection") or detail_api.get("hom_portfolio")
                  or listing_data.get("collection") or "").strip()
    if collection:
        row["Collection"] = collection

    # ── Technical data from detail API ────────────────────────────────────
    if detail_api.get("upc_code"):
        row["UPC Code"] = detail_api["upc_code"]


    # ── Spec items from rendered HTML ──────────────────────────────────────
    spec_els = await page.query_selector_all(
        ".specifications_item .specifications_item_title"
    )
    spec_texts = []
    for el in spec_els:
        t = await el.inner_text()
        t = t.strip()
        if t:
            spec_texts.append(t)

    spec_data = _parse_spec_items(spec_texts)
    row.update(spec_data)

    # If Dimensions not captured from spec items, fall back to detail API
    if "Width" not in row and detail_api.get("meta_wpcf_product_dimensions"):
        dims = _parse_art_dimensions(detail_api["meta_wpcf_product_dimensions"])
        row.update(dims)

    # ── About the Collection ───────────────────────────────────────────────
    about_el = await page.query_selector(".specifications_item .specifications_item_info")
    if about_el:
        about_text = await about_el.inner_text()
        about_text = about_text.strip()
        if about_text and len(about_text) > 20:
            row["About Collection"] = about_text

    # Remove None / empty
    return {k: v for k, v in row.items() if v is not None and v != ""}


# ── Category scraper ──────────────────────────────────────────────────────────

async def scrape_category(page, cat: dict, writer: ExcelWriter) -> None:
    if TEST_MODE:
        print(f"  [TEST: max {TEST_MAX_PRODUCTS} products per category]")

    seen_skus: set[str] = set()
    all_listing_data: list[dict] = []

    for listing_url in cat["links"]:
        room, category_param = _parse_listing_url(listing_url)
        if not room or not category_param:
            print(f"  [WARN] Could not parse Room/Category from {listing_url}")
            continue

        products = _collect_all_skus(room, category_param)
        for p in products:
            sku = p.get("sku_number", "")
            if sku and sku not in seen_skus:
                seen_skus.add(sku)
                all_listing_data.append(p)

            if TEST_MODE and len(all_listing_data) >= TEST_MAX_PRODUCTS:
                break
        if TEST_MODE and len(all_listing_data) >= TEST_MAX_PRODUCTS:
            break

    label = " [TEST]" if TEST_MODE else ""
    print(f"  {cat['name']}: {len(all_listing_data)} products{label}")

    global_idx = 1
    for listing_item in all_listing_data:
        sku = listing_item.get("sku_number", "")
        if not sku:
            continue
        try:
            row = await scrape_product(page, sku, listing_item)
            if not row.get("SKU"):
                row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
            writer.write_row(row, category_name=cat["name"])
            global_idx += 1
        except Exception as e:
            print(f"  [ERROR] {sku}: {e}")

        await async_polite_delay()

    print(f"  {cat['name']}: wrote {global_idx - 1} rows")


# ── Main ──────────────────────────────────────────────────────────────────────

async def main() -> None:
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE and TEST_MAX_CATEGORIES < len(categories):
        categories = categories[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue
            writer.add_sheet(cat["name"], cat["links"][0],
                             studio_columns=cat["studio_columns"])
            await scrape_category(page, cat, writer)

    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
