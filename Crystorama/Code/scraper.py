import asyncio, json, os, sys, re
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, parse_spec_block, safe_float,
)

VENDOR_NAME  = os.environ.get("VENDOR_NAME", "Crystorama")
HEADLESS     = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH  = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE    = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://crystoramalightinglights.com"
AJAX_URL = f"{BASE_URL}/on/demandware.store/Sites-lny_us-Site/en_US/Search-ShowAjax"


def _extract_cgid(html: str) -> str | None:
    """Pull cgid= value from a sort-button data-url attribute."""
    m = re.search(r'cgid=([^&"\']+)', html)
    return m.group(1) if m else None


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from a listing, using AJAX pagination."""
    await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)
    html = await page.content()

    cgid = _extract_cgid(html)
    if not cgid:
        # Fallback: collect links from the rendered page only
        anchors = await page.query_selector_all(".product-tile a.product-tile-image-container")
        links = []
        for a in anchors:
            href = await a.get_attribute("href")
            if href:
                links.append(urljoin(BASE_URL, href))
        return list(dict.fromkeys(links))

    links: list[str] = []
    start = 0
    sz = 48
    while True:
        ajax = f"{AJAX_URL}?cgid={cgid}&start={start}&sz={sz}"
        await page.goto(ajax, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)
        anchors = await page.query_selector_all("a.product-tile-image-container")
        if not anchors:
            break
        for a in anchors:
            href = await a.get_attribute("href")
            if href:
                links.append(urljoin(BASE_URL, href))
        if len(anchors) < sz:
            break
        start += sz
        await async_polite_delay()

    return list(dict.fromkeys(links))


async def scrape_product(page, url: str) -> list[dict]:
    """Return a single-element list with all fields from the product page."""
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)
    html = await page.content()

    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")

    data: dict = {"Source": url}

    # ── JSON-LD ──────────────────────────────────────────────────────────
    ld = {}
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            obj = json.loads(script.string or "")
            if obj.get("@type") == "Product":
                ld = obj
                break
        except Exception:
            pass

    data["Product Name"] = clean_text(ld.get("name", ""))
    data["Description"]  = clean_text(ld.get("description", ""))
    data["SKU"]          = clean_text(ld.get("mpn") or ld.get("sku", ""))
    data["Manufacturer"] = VENDOR_NAME

    offers = ld.get("offers", {})
    if isinstance(offers, list):
        offers = offers[0] if offers else {}
    raw_price = offers.get("price", "")
    data["Price"] = clean_price(str(raw_price)) if raw_price else None

    imgs = ld.get("image", [])
    if isinstance(imgs, str):
        imgs = [imgs]
    data["Image URL"] = imgs[0] if imgs else ""

    # High-res image from carousel as fallback
    if not data["Image URL"]:
        img_el = soup.select_one(".primary-images img, .carousel-item img")
        if img_el:
            data["Image URL"] = img_el.get("src", "")

    # ── Overview attrs (Collection, Color) ───────────────────────────────
    dad = soup.select_one(".description-and-detail")
    if dad:
        for col_div in dad.select(".product-overview-attrs .col-6"):
            h3 = col_div.select_one("h3")
            a  = col_div.select_one("a")
            if h3 and a:
                key = h3.get_text(strip=True)
                val = a.get_text(strip=True)
                if key == "Collection":
                    data["Collection"] = val
                elif key == "Color":
                    data["Color"] = val

    # ── Spec items ───────────────────────────────────────────────────────
    specs_raw: dict[str, str] = {}
    for item in soup.select(".spec-item"):
        name_el = item.select_one(".spec-name")
        val_el  = item.select_one(".spec-value")
        if name_el and val_el:
            k = name_el.get_text(strip=True).rstrip(":")
            v = val_el.get_text(strip=True)
            specs_raw[k] = v

    # Map spec keys → output columns
    SPEC_MAP = {
        "SKU":                        "SKU",
        "UPC":                        "UPC",
        "Shipping Method":            "Shipping Method",
        "Dimmable":                   "Dimming",
        "Lamping Features":           "Lamping",
        "Lamping Included":           "Lamping Included",
        "Lamping Type":               "Lamp Type",
        "Primary Number of Bulbs":    "Bulb Qty",
        "Total Number of Bulbs":      "Socket Qty",
        "Socket":                     "Socket Type",
        "Voltage":                    "Voltage",
        "Wattage Max":                "Wattage",
        "Lead Wire Length":           "Lead Wire Length",
        "Backplate/Canopy Extension": "Canopy Extension",
        "Backplate/Canopy Width":     "Canopy",
        "Dimensions":                 "Dimensions",
        "Width":                      "Width",
        "Height":                     "Height",
        "Depth":                      "Depth",
        "Length":                     "Length",
        "Diameter":                   "Diameter",
        "Extension":                  "Extension",
        "Maximum Adjustable Height":  "Hanging Length",
        "Weight":                     "Weight",
        "Country of Origin":          "Origin",
        "Install Position":           "Mounting",
        "UL Ratings":                 "UL Ratings",
        "Warranty":                   "Warranty",
        "Chain Cord Features":        "Chain Length",
        "Crystal Features":           "Crystal Features",
        "Glass Features":             "Glass Features",
        "Material":                   "Material",
        "Shape":                      "Shape",
        "Features":                   "Specifications",
        "Prop 65":                    "Prop 65",
        "Title 20":                   "Title 20",
        "Brand Product Description":  "Brand Product Description",
        "Brand Category":             "Brand Category",
    }

    for raw_k, raw_v in specs_raw.items():
        out_k = SPEC_MAP.get(raw_k, raw_k)  # unmapped keys go in as-is
        if out_k not in data or not data[out_k]:
            data[out_k] = raw_v

    # ── Numeric cleanup ───────────────────────────────────────────────────
    for dim_field in ("Width", "Height", "Depth", "Diameter", "Length", "Weight",
                      "Wattage", "Voltage"):
        if data.get(dim_field):
            v = safe_float(str(data[dim_field]))
            if v is not None:
                data[dim_field] = v

    # ── Dimensions string ─────────────────────────────────────────────────
    if not data.get("Dimensions"):
        parts = []
        for lbl, k in (("W", "Width"), ("H", "Height"), ("D", "Depth"), ("Dia", "Diameter")):
            if data.get(k):
                parts.append(f'{data[k]}{lbl}')
        if parts:
            data["Dimensions"] = " x ".join(parts)

    # ── Product Family Id ─────────────────────────────────────────────────
    data["Product Family Id"] = extract_family_id(data.get("Product Name", ""))

    # Remove None / empty values
    data = {k: v for k, v in data.items() if v is not None and v != ""}

    return [data]


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    cats = info["categories"]
    if TEST_MODE:
        cats = cats[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
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
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"[{cat['name']}] {len(all_product_urls)} products")

            global_idx = 1
            for url in all_product_urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  ERROR {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
