import asyncio, json, os, sys, re
from pathlib import Path
from urllib.parse import urljoin

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME        = os.environ.get("VENDOR_NAME", "Regina Andrew")
HEADLESS           = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH        = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE          = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.reginaandrew.com"

# Map lowercase spec labels from the page → output field names
SPEC_LABEL_MAP = {
    "height":            "Height",
    "width":             "Width",
    "depth":             "Depth",
    "length":            "Length",
    "diameter":          "Diameter",
    "weight":            "Weight",
    "finish":            "Finish",
    "color":             "Color",
    "material":          "Materials",
    "materials":         "Materials",
    "socket":            "Socket",
    "wattage":           "Wattage",
    "bulb qty":          "Bulb Qty",
    "bulb type":         "Bulb Type",
    "compatible bulb":   "Compatible Bulb",
    "canopy":            "Canopy",
    "shade dims":        "Shade Details",
    "shade details":     "Shade Details",
    "ul rating":         "Rating",
    "wiring type":       "Wiring Type",
    "extension":         "Extension",
    "chain length":      "Chain Length",
    "base":              "Base",
    "seat height":       "Seat Height",
    "seat depth":        "Seat Depth",
    "arm height":        "Arm Height",
    "collection":        "Collection",
    "designer":          "Designer",
    "origin":            "Origin",
    "lead time":         "Lead Time",
    "com":               "COM",
    "col":               "COL",
    "lightsource":       "Lightsource",
    "light source":      "Lightsource",
    "color temperature": "Color Temperature",
    "content":           "Content",
    "cushion":           "Cushion",
    "dimensions":        "Dimensions",
}

# Dimension labels → parse_dimensions
DIM_LABELS = {"height", "width", "depth", "length", "diameter"}


def _parse_jsonld(html: str) -> list[dict]:
    results = []
    for m in re.finditer(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html, re.DOTALL
    ):
        try:
            results.append(json.loads(m.group(1)))
        except Exception:
            pass
    return results


def _clean_img(url: str) -> str:
    """Strip resize query params from image URL."""
    if not url:
        return ""
    if not url.startswith("http"):
        url = urljoin(BASE_URL, url)
    return url.split("?")[0]


def _offers_price(offers) -> float | None:
    if not offers:
        return None
    if isinstance(offers, list):
        offers = offers[0] if offers else {}
    price_val = offers.get("price") or offers.get("priceSpecification", {}).get("price")
    return clean_price(str(price_val)) if price_val is not None else None


async def _parse_specs(page) -> dict:
    """Extract all 'Label: Value' spec divs from the product detail page."""
    specs = {}
    # Grab text from every <div> on the page that matches "word(s): text"
    divs = await page.query_selector_all("div")
    for div in divs:
        try:
            txt = (await div.inner_text()).strip()
        except Exception:
            continue
        # Must be a single short line (spec row, not a big block)
        if "\n" in txt or len(txt) > 200:
            continue
        m = re.match(r'^([A-Za-z][A-Za-z0-9 /]+?)\s*:\s*(.+)$', txt)
        if not m:
            continue
        label = m.group(1).strip().lower()
        value = clean_text(m.group(2).strip())
        if not value or label not in SPEC_LABEL_MAP:
            continue
        field = SPEC_LABEL_MAP[label]
        if field not in specs:
            # For dimension fields store as number only (no inch marks)
            if label in DIM_LABELS:
                num = re.sub(r'["\']|in\b', '', value).strip()
                specs[field] = num
            else:
                specs[field] = value
    return specs


async def scrape_product(page, url: str) -> list[dict]:
    """
    Returns a list of row dicts.
    Usually 1 row per URL, but ProductGroup lighting pages return 1 row per variant.
    """
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    html = await page.content()
    jsonld_blocks = _parse_jsonld(html)

    # Description
    description = ""
    for sel in ("section.item-details-description", ".item-details-description"):
        el = await page.query_selector(sel)
        if el:
            description = clean_text(await el.inner_text())
            break

    # Primary image — prefer high-res data-src from the detail page gallery
    primary_image = ""
    for img_sel in (
        "li.product-details-image-gallery-container img",
        ".product-details-image img",
        ".product-image img",
    ):
        img_el = await page.query_selector(img_sel)
        if img_el:
            src = (
                await img_el.get_attribute("data-src")
                or await img_el.get_attribute("data-zoom-image")
                or await img_el.get_attribute("src")
                or ""
            )
            if src and not src.startswith("data:"):
                primary_image = _clean_img(src)
                break

    # Specs
    specs = await _parse_specs(page)

    # Find Product or ProductGroup in JSON-LD
    product_data = None
    for block in jsonld_blocks:
        if isinstance(block, dict) and block.get("@type") in ("Product", "ProductGroup"):
            product_data = block
            break
        if isinstance(block, list):
            for item in block:
                if isinstance(item, dict) and item.get("@type") in ("Product", "ProductGroup"):
                    product_data = item
                    break

    # ── Fallback: no JSON-LD found ────────────────────────────────────────────
    if not product_data:
        name = ""
        name_el = await page.query_selector("h1.product-details-full-content-header-title")
        if name_el:
            name = sentence_case(await name_el.inner_text())
        sku = ""
        sku_el = await page.query_selector(".product-line-sku-value")
        if sku_el:
            sku = clean_text(await sku_el.inner_text())
        price = None
        price_el = await page.query_selector("span.product-views-price-lead")
        if price_el:
            rate = await price_el.get_attribute("data-rate")
            price = clean_price(rate) if rate else clean_price(await price_el.inner_text())

        row = {
            "Source":            url,
            "Product Name":      name,
            "Product Family Id": extract_family_id(name) if name else "",
            "SKU":               sku,
            "Price":             price,
            "Description":       description,
            "Image URL":         primary_image,
        }
        row.update(specs)
        return [row]

    ptype = product_data.get("@type")

    # ── ProductGroup → one row per variant ───────────────────────────────────
    if ptype == "ProductGroup":
        variants = product_data.get("hasVariant", [])
        if not variants:
            variants = [product_data]

        family_name = sentence_case(product_data.get("name", ""))
        group_img   = _clean_img(
            (product_data.get("image") or [""])[0]
            if isinstance(product_data.get("image"), list)
            else product_data.get("image", primary_image)
        ) or primary_image

        rows = []
        for variant in variants:
            name = sentence_case(variant.get("name", family_name))
            sku  = variant.get("sku", "")
            price = _offers_price(variant.get("offers"))

            img = variant.get("image") or group_img
            if isinstance(img, list):
                img = img[0] if img else group_img
            img = _clean_img(img) or primary_image

            row = {
                "Source":            url,
                "Product Name":      name,
                "Product Family Id": family_name,
                "SKU":               sku,
                "Price":             price,
                "Description":       description,
                "Image URL":         img,
            }
            row.update(specs)
            rows.append(row)
        return rows

    # ── Single Product ────────────────────────────────────────────────────────
    name  = sentence_case(product_data.get("name", ""))
    sku   = product_data.get("sku", "")
    price = _offers_price(product_data.get("offers"))

    img = product_data.get("image") or primary_image
    if isinstance(img, list):
        img = img[0] if img else primary_image
    img = _clean_img(img) or primary_image

    row = {
        "Source":            url,
        "Product Name":      name,
        "Product Family Id": extract_family_id(name) if name else "",
        "SKU":               sku,
        "Price":             price,
        "Description":       description,
        "Image URL":         img,
    }
    row.update(specs)
    return [row]


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from a category listing, handling pagination."""
    links: list[str] = []
    seen:  set[str]  = set()
    page_num = 1

    while True:
        url = listing_url if page_num == 1 else f"{listing_url}?page={page_num}"
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)

        anchors = await page.query_selector_all("a.facets-item-cell-grid-link-image")
        if not anchors:
            break

        found_new = False
        for a in anchors:
            href = await a.get_attribute("href") or ""
            if href:
                full = urljoin(BASE_URL, href)
                if full not in seen:
                    seen.add(full)
                    links.append(full)
                    found_new = True

        # Check for next-page link in <head>
        next_el = await page.query_selector('link[rel="next"]')
        if not next_el or not found_new:
            break

        page_num += 1
        await async_polite_delay()

    return links


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = [c for c in categories if c["links"]][:TEST_MAX_CATEGORIES]
        print(f"[TEST MODE] {len(categories)} categories, max {TEST_MAX_PRODUCTS} products each")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            print(f"\n[{cat['group']}] {cat['name']} — collecting links…")
            seen_urls: set[str] = set()
            urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        urls.append(u)
            if TEST_MODE:
                urls = urls[:TEST_MAX_PRODUCTS]
            print(f"  {len(urls)} products")

            idx = 1
            for url in urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        idx += 1
                    print(f"  [{idx-1}] {url.split('/')[-1]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"\n[Done] Saved -> {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
