import asyncio, json, os, re, sys
from pathlib import Path
from urllib.parse import urlparse, parse_qs, urlencode

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    clean_text, clean_price, generate_sku, extract_family_id,
    parse_dimensions, async_polite_delay,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Bramble")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://www.brambleco.com"
IMAGE_BASE = "https://s3.amazonaws.com/emuncloud-staticassets/productImages/bram019/large"

def _skip_datum(label: str) -> bool:
    """Return True for Package-related shipping fields the user wants excluded."""
    l = label.lower()
    if "number of packages" in l:
        return True
    # Skip any "Package N ..." label (dimensions, weight)
    if re.match(r"^package\s*\d+", l):
        return True
    if l == "package weight total (lbs)":
        return True
    return False


# ── URL helpers ───────────────────────────────────────────────────────────────

def _nav_params_from_listing(listing_url: str) -> str:
    """Extract only nav_level_* params (not $MultiView) from a listing URL."""
    if "#?" in listing_url:
        hash_part = listing_url.split("#?")[1]
    elif "#" in listing_url:
        hash_part = listing_url.split("#")[1].lstrip("?")
    else:
        hash_part = ""
    params = [p for p in hash_part.split("&") if p.startswith("nav_level_")]
    return "&".join(params)


def _product_url(product_id: str, nav_params: str) -> str:
    """Build product URL with $MultiView=Yes (exactly once) + nav params."""
    base = f"{BASE_URL}/shop/{product_id}"
    parts = ["$MultiView=Yes"]
    if nav_params:
        parts.append(nav_params)
    parts += [f"productId={product_id}", "position=-1",
              "orderBy=Featured,Id", "context=shop", "page=1"]
    return f"{base}?{'&'.join(parts)}"


def _listing_page_url(listing_url: str, page_num: int) -> str:
    """Replace page=N in the hash query string."""
    return re.sub(r"page=\d+", f"page={page_num}", listing_url)


# ── Listing scraper ───────────────────────────────────────────────────────────

async def get_product_entries(page, listing_url: str,
                              max_products: int | None = None) -> list[dict]:
    """
    Paginate through the listing and return dicts with {id, name}.
    Uses data-product-id / data-name attributes from yotpo divs.
    """
    seen_ids: set[str] = set()
    results: list[dict] = []
    page_num = 1

    while True:
        url = _listing_page_url(listing_url, page_num)
        try:
            await page.goto(url, timeout=120_000, wait_until="domcontentloaded")
        except Exception as e:
            print(f"  [WARN] listing page {page_num} timeout/err: {str(e)[:60]}")
            break
        await page.wait_for_timeout(8000)

        html = await page.content()

        # Extract product IDs + names from yotpo data attributes
        entries = re.findall(
            r'data-product-id="(\d+)"[^>]*data-name="([^"]+)"',
            html
        )
        # Also try reversed attribute order
        entries += re.findall(
            r'data-name="([^"]+)"[^>]*data-product-id="(\d+)"',
            html
        )
        # Normalise to (id, name) pairs
        normalized = []
        for e in entries:
            if e[0].isdigit():
                normalized.append((e[0], e[1]))
            else:
                normalized.append((e[1], e[0]))

        # Fallback: just extract IDs from href="#/shop/{id}"
        if not normalized:
            raw_ids = re.findall(r'href="#/shop/(\d+)', html)
            normalized = [(rid, "") for rid in raw_ids]

        new_count = 0
        for pid, pname in normalized:
            if pid not in seen_ids:
                seen_ids.add(pid)
                results.append({"id": pid, "name": pname})
                new_count += 1
            if max_products and len(results) >= max_products:
                break

        if max_products and len(results) >= max_products:
            break
        if new_count == 0:
            break  # no new products — pagination exhausted

        page_num += 1

    return results


# ── Product scraper ───────────────────────────────────────────────────────────

async def scrape_product(page, product_id: str, product_name: str,
                         nav_params: str) -> dict:
    url = _product_url(product_id, nav_params)

    try:
        await page.goto(url, timeout=120_000, wait_until="domcontentloaded")
    except Exception as e:
        print(f"  [WARN] product {product_id} goto err: {str(e)[:60]}")

    # Wait for the product detail panel (ensures DIMENSIONS/SKU p-elements render)
    try:
        await page.wait_for_selector(".one-up-details p", timeout=30_000)
    except Exception:
        pass
    # Additional wait for datum elements (shipping/lighting section)
    try:
        await page.wait_for_selector("datum", timeout=15_000)
    except Exception:
        pass
    await page.wait_for_timeout(3000)

    # ── Product Name ──────────────────────────────────────────────────────────
    if not product_name:
        # Try yotpo div first
        yotpo = await page.query_selector("div[data-product-id][data-name]")
        if yotpo:
            product_name = (await yotpo.get_attribute("data-name")) or ""
    if not product_name:
        # Try h1 inside one-up section
        h_el = await page.query_selector(".one-up-header h1, [class*='one-up'] h1, h1.product-name")
        if h_el:
            product_name = (await h_el.inner_text()).strip()

    # ── Parse p-text lines in detail panel ───────────────────────────────────
    # Try the primary selector, fall back to any p inside a one-up container
    p_els = await page.query_selector_all(
        ".one-up-details p, [class*='one-up-detail'] p, "
        "[class*='one-up'] p, shopping-item-details p"
    )
    shown_in, item_type, raw_dim = "", "", ""

    # Collect all non-empty text lines first so we can do look-ahead
    p_texts: list[str] = []
    for p_el in p_els:
        t = (await p_el.inner_text()).strip()
        if t:
            p_texts.append(t)

    pending_dim = False  # True when we saw a standalone "DIMENSIONS" label
    for i, text in enumerate(p_texts):
        upper = text.upper()

        if pending_dim:
            # The previous line was just "DIMENSIONS" — this line is the value
            raw_dim = text
            pending_dim = False
            continue

        if text.lower().startswith("shown in:"):
            shown_in = text.split(":", 1)[1].strip()
        elif upper.startswith("DIMENSIONS"):
            val = re.sub(r"(?i)^dimensions\s*:?\s*", "", text).strip()
            if val:
                raw_dim = val
            else:
                # Value might be on the next line
                pending_dim = True
        elif upper in ("CUSTOM", "QUICK SHIP", "SALE", "QUICK SHIP/CUSTOM EXPRESS"):
            item_type = text.strip()
        elif upper.startswith("SKU"):
            product_id = re.sub(r"(?i)sku\s*:?\s*", "", text).strip() or product_id

    # Fallback: scan full page text for DIMENSIONS line if still missing
    if not raw_dim:
        try:
            body_text = await page.inner_text("body")
            m = re.search(r"DIMENSIONS\s*:?\s*([WwDdHhLl\d][^\n]{3,60}(?:in|cm)?)", body_text)
            if m:
                raw_dim = m.group(1).strip()
        except Exception:
            pass

    # ── Datum elements (Shipping + Lighting details) ──────────────────────────
    datum_data: dict = {}
    datum_els = await page.query_selector_all("datum")
    for d_el in datum_els:
        label = (await d_el.get_attribute("label") or "").strip()
        if not label or _skip_datum(label):
            continue
        inner = (await d_el.inner_text()).strip()
        # Remove "label: " prefix from inner text
        if inner.lower().startswith(label.lower()):
            value = inner[len(label):].lstrip(": ").strip()
        else:
            value = inner
        # Skip placeholder / empty values
        if value and value not in ("-", "—", "N/A", "n/a"):
            datum_data[label] = value

    # ── Image ─────────────────────────────────────────────────────────────────
    image_url = f"{IMAGE_BASE}/{product_id}.jpg"
    # Try to get actual image src from page (catches alternate CDN paths)
    img_el = await page.query_selector(
        "shopping-item-image img[src*='productImages'], "
        ".one-up img[src*='s3.amazonaws'], "
        "[class*='image'] img[src*='productImages']"
    )
    if img_el:
        src = await img_el.get_attribute("src") or ""
        if "large" in src and ".jpg" in src:
            image_url = src.split("?")[0]  # strip cache-busting query

    # ── Tearsheet ─────────────────────────────────────────────────────────────
    tearsheet = ""
    ts_el = await page.query_selector(
        "a[href*='ProductDetailSheet'], a[href*='tearsheet']"
    )
    if ts_el:
        tearsheet = (await ts_el.get_attribute("href") or "").strip()

    # ── Downloads (product images zip) ───────────────────────────────────────
    download_url = ""
    dl_el = await page.query_selector("a[href*='PRODUCT IMAGES'], a:text-is('PRODUCT IMAGES')")
    if dl_el:
        download_url = (await dl_el.get_attribute("href") or "").strip()

    # ── Description / Product Story ───────────────────────────────────────────
    desc = ""
    desc_el = await page.query_selector(".one-up-long-desc")
    if desc_el:
        desc_text = (await desc_el.inner_text()).strip()
        # Only use as description if it's different from the product name
        if desc_text and desc_text.lower() != product_name.lower():
            desc = desc_text

    # ── Build row ─────────────────────────────────────────────────────────────
    row: dict = {
        "Manufacturer":      VENDOR_NAME,
        "Source":            url,
        "Image URL":         image_url,
        "Product Name":      clean_text(product_name) or "",
        "Product Family Id": extract_family_id(product_name) if product_name else "",
        "SKU":               product_id,
    }

    if shown_in:
        row["Shown In"] = shown_in
    if item_type:
        row["Item Type"] = item_type
    if desc:
        row["Description"] = desc
    if tearsheet:
        row["Tearsheet Link"] = tearsheet
    if download_url:
        row["Product Images Download"] = download_url

    # Dimensions (parse from the rendered "W 30 x D 18 x H 30 in" string)
    if raw_dim:
        row["Dimensions"] = raw_dim.replace(" in", "").strip()
        dims = parse_dimensions(raw_dim)
        for k in ("Width", "Depth", "Height", "Length", "Diameter"):
            if dims.get(k):
                row[k] = dims[k]

    # Shipping / Lighting datum fields
    # Net Weight → Weight; everything else by its label
    for label, value in datum_data.items():
        if label.lower() == "net weight (lbs)":
            row["Weight"] = value
        else:
            row[label] = value

    # Remove None / empty
    return {k: v for k, v in row.items() if v is not None and v != ""}


# ── Category scraper ──────────────────────────────────────────────────────────

async def scrape_category(page, cat: dict, writer: ExcelWriter) -> None:
    if TEST_MODE:
        print(f"  [TEST: max {TEST_MAX_PRODUCTS} products per category]")

    seen_ids: set[str] = set()
    all_entries: list[dict] = []

    for listing_url in cat["links"]:
        max_p = TEST_MAX_PRODUCTS if TEST_MODE else None
        entries = await get_product_entries(
            page, listing_url,
            max_products=(max_p - len(all_entries)) if max_p else None,
        )
        for e in entries:
            if e["id"] not in seen_ids:
                seen_ids.add(e["id"])
                all_entries.append(e)
        if TEST_MODE and len(all_entries) >= TEST_MAX_PRODUCTS:
            break

    label = " [TEST]" if TEST_MODE else ""
    print(f"  {cat['name']}: {len(all_entries)} products{label}")

    nav_params = _nav_params_from_listing(cat["links"][0])
    global_idx = 1

    for entry in all_entries:
        pid   = entry["id"]
        pname = entry.get("name", "")
        try:
            row = await scrape_product(page, pid, pname, nav_params)
            if not row.get("SKU"):
                row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
            writer.write_row(row, category_name=cat["name"])
            global_idx += 1
        except Exception as e:
            print(f"  [ERROR] product {pid}: {e}")

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
