import asyncio, json, os, sys, re
from pathlib import Path
from urllib.parse import urlparse, parse_qs
import requests as _requests

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

FAST_SIMON_UUID     = "1382dbe2-7b41-41ac-8cfe-bfbf7e3e3dd5"
FAST_SIMON_STORE_ID = 1

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, parse_spec_block, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "P Kaufman")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))


def _clean_html(text: str) -> str:
    """Strip HTML tags and decode basic entities."""
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&amp;", "&").replace("&nbsp;", " ").replace("&#x2F;", "/")
    return " ".join(text.split()).strip()


async def scrape_product(page, url: str) -> list[dict]:
    """Return a list with one dict per product page (each colorway is its own URL)."""
    row = {"Source URL": url}

    try:
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
    except Exception as e:
        print(f"  [WARN] Failed to load {url}: {e}")
        return [row]

    html = await page.content()

    # --- Product Name ---
    name_m = re.search(
        r'<h1[^>]*class="[^"]*productView-title[^"]*"[^>]*>(.*?)</h1>', html, re.DOTALL
    )
    if name_m:
        row["Product Name"] = clean_text(name_m.group(1))

    # --- SKU from BCData JS variable ---
    bc_m = re.search(r'"sku":\s*"([^"]+)"', html)
    if bc_m:
        row["SKU"] = bc_m.group(1).strip()

    # --- Image URL (highest resolution via data-zoom-image) ---
    zoom_m = re.search(r'data-zoom-image="([^"]+)"', html)
    if zoom_m:
        row["Image URL"] = zoom_m.group(1)
    else:
        # fallback: main product image src
        img_m = re.search(r'class="productView-img[^"]*"[^>]*src="([^"]+)"', html)
        if img_m:
            row["Image URL"] = img_m.group(1)

    # --- All specifications from the spec list ---
    specs_section_m = re.search(
        r'product-specifications-content[^>]+>(.*?)(?:</div>|<div class="product-description)',
        html, re.DOTALL
    )
    seen_keys: set = set()
    if specs_section_m:
        pairs = re.findall(
            r'productView-info-name[^>]*>\s*(.*?)\s*</span>.*?productView-info-value[^>]*>\s*(.*?)\s*</span>',
            specs_section_m.group(1), re.DOTALL
        )
        for raw_key, raw_val in pairs:
            key = _clean_html(raw_key).strip()
            val = _clean_html(raw_val).strip()
            if not key or not val:
                continue
            # Map spec keys to our column names
            col = _map_spec_key(key)
            if col and col not in seen_keys:
                row[col] = val
                seen_keys.add(col)
            elif col and col in seen_keys:
                # Already have this column - skip duplicate (e.g. Brand appears twice)
                pass
            else:
                # Store unmapped specs verbatim
                if key not in seen_keys:
                    row[key] = val
                    seen_keys.add(key)

    # --- Collection (shown outside spec table as a link) ---
    if "Collection" not in row:
        coll_m = re.search(
            r'<span[^>]*>\s*Collection:\s*</span>\s*<a[^>]*>(.*?)</a>', html, re.DOTALL
        )
        if not coll_m:
            coll_m = re.search(
                r'Collection:\s*</[^>]+>\s*<[^>]+>(.*?)</[^>]+>', html, re.DOTALL
            )
        if coll_m:
            row["Collection"] = _clean_html(coll_m.group(1)).strip()

    # --- Clean fabric Width to numeric only (strip inch marks) ---
    if row.get("Width"):
        w_clean = re.sub(r'["\s]', "", str(row["Width"]))
        try:
            row["Width"] = str(safe_float(w_clean) or w_clean)
        except Exception:
            pass

    # --- Manufacturer always vendor name ---
    row["Manufacturer"] = VENDOR_NAME

    # --- Product Family Id ---
    if row.get("Product Name") and not row.get("Product Family Id"):
        # Pattern Name from specs is the best family id
        if row.get("Pattern Name"):
            row["Product Family Id"] = row["Pattern Name"]
        else:
            row["Product Family Id"] = extract_family_id(row["Product Name"])

    return [row]


def _map_spec_key(key: str) -> str | None:
    """Map raw spec label from the page to our output column name."""
    mapping = {
        "Collection Name":       "Collection",
        "Pattern Name":          "Pattern Name",
        "Color Name":            "Color",
        "Division Name":         "Division Name",
        "Design Status":         "Design Status",
        "New Product":           "New Product",
        "Discontinued":          "Discontinued",
        "Fabric Content":        "Material",
        "Fabric Width":          "Width",
        "Vertical Repeat":       "Vertical Repeat",
        "Horizontal Repeat":     "Horizontal Repeat",
        "Repeat":                "Repeat",
        "Pattern Match":         "Pattern Match",
        "Fabric Weight":         "Fabric Weight",
        "Country of Origin":     "Origin",
        "Finish":                "Finish",
        "Fabric Care":           "Care Instructions",
        "End Use":               "End Use",
        "Print or Woven":        "Construction",
        "Pattern design":        "Pattern",
        "Abrasion-Wyzenbeek":    "Abrasion-Wyzenbeek",
        "Abrasion-Martindale":   "Abrasion-Martindale",
        "Flammability":          "Flammability",
        "Sustainability":        "Sustainability",
        "Brand":                 "Brand",
        "Fabric Type":           "Fabric Type",
        "UPC Code":              "UPC Code",
        "Pile Height":           "Pile Height",
        "Backing":               "Backing",
        "Color Name":            "Color",
    }
    return mapping.get(key)


def _fastsimon_search_links(query: str) -> list[str]:
    """Call FastSimon full_text_search API and return all product URLs."""
    all_links: list[str] = []
    page_num = 1
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://pkaufmann.com/"}

    while True:
        api_url = (
            f"https://api.fastsimon.com/full_text_search"
            f"?request_source=v-next&src=v-next"
            f"&UUID={FAST_SIMON_UUID}&uuid={FAST_SIMON_UUID}"
            f"&store_id={FAST_SIMON_STORE_ID}&api_type=json"
            f"&facets_required=1&products_per_page=50&narrow=%5B%5D"
            f"&q={query}&page_num={page_num}&sort_by=creation_date"
            f"&with_product_attributes=true&qs=false"
        )
        try:
            r = _requests.get(api_url, headers=headers, timeout=20)
            if r.status_code != 200:
                print(f"    [WARN] FastSimon API {r.status_code} for q={query} page={page_num}")
                break
            data = r.json()
        except Exception as e:
            print(f"    [WARN] FastSimon API error: {e}")
            break

        items = data.get("items", [])
        if not items:
            break

        page_links = []
        for item in items:
            rel = item.get("u", "")
            if rel:
                full = "https://pkaufmann.com" + rel if rel.startswith("/") else rel
                page_links.append(full)

        all_links.extend(page_links)
        total = data.get("total_results", len(all_links))
        print(f"    FastSimon q={query} page {page_num}: {len(page_links)} products (total={total})")

        if len(all_links) >= total:
            break
        page_num += 1

    return all_links


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from a listing page, handling ?page=N pagination.
    For /search-results/?q=... URLs, calls the FastSimon API directly.
    """
    # --- FastSimon search results URL → use API directly ---
    parsed = urlparse(listing_url)
    if "/search-results/" in parsed.path or "search-results" in parsed.path:
        qs = parse_qs(parsed.query)
        query = qs.get("q", [None])[0]
        if query:
            return _fastsimon_search_links(query)
        else:
            print(f"    [WARN] No query param in search URL: {listing_url}")
            return []

    # --- Regular category / view-all page → Playwright scraping ---
    all_links: list[str] = []
    page_num = 1

    while True:
        url = listing_url if page_num == 1 else f"{listing_url.rstrip('/')}?page={page_num}"
        print(f"    Listing page {page_num}: {url}")

        try:
            await page.goto(url, timeout=45_000, wait_until="networkidle")
        except Exception:
            await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(4000)

        # Find product cards
        cards = await page.query_selector_all("[data-entity-id]")
        if not cards:
            cards = await page.query_selector_all("li.isp_grid_product")

        if not cards:
            if page_num == 1:
                print(f"    [WARN] No product cards found on {listing_url}")
            break

        page_links: list[str] = []
        for card in cards:
            link_el = await card.query_selector("a.card-figure__link, a[href]")
            if link_el:
                href = await link_el.get_attribute("href")
                if href:
                    if href.startswith("/"):
                        href = "https://pkaufmann.com" + href
                    if "pkaufmann.com/" in href and href not in all_links:
                        page_links.append(href)

        if not page_links:
            break

        all_links.extend(page_links)
        print(f"      Found {len(page_links)} products (total so far: {len(all_links)})")

        # Check if there's a next page
        html = await page.content()
        if f"page={page_num + 1}" not in html:
            break

        page_num += 1
        await async_polite_delay()

    return all_links


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    if TEST_MODE:
        print(f"[TEST MODE] max {TEST_MAX_PRODUCTS} products per category")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        cats = info["categories"]
        if TEST_MODE:
            cats = cats[:TEST_MAX_CATEGORIES]

        for cat in cats:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls: set[str] = set()
            all_product_urls: list[str] = []

            for listing_url in cat["links"]:
                print(f"\n  Collecting links from: {listing_url}")
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"\n  Scraping {len(all_product_urls)} products for category: {cat['name']}")

            global_idx = 1
            for url in all_product_urls:
                print(f"    [{global_idx}] {url}")
                try:
                    variant_rows = await scrape_product(page, url)
                    for variant in variant_rows:
                        if not variant.get("SKU"):
                            variant["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not variant.get("Product Family Id") and variant.get("Product Name"):
                            variant["Product Family Id"] = extract_family_id(variant["Product Name"])
                        writer.write_row(variant, category_name=cat["name"])
                    global_idx += 1
                except Exception as e:
                    print(f"    [ERROR] {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
