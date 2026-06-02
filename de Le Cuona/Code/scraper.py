import asyncio
import json
import os
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    sentence_case,
    clean_price,
    generate_sku,
    extract_family_id,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "de Le Cuona")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get(
    "OUTPUT_PATH",
    str(PROJECT_ROOT / "de Le Cuona" / "Data" / "de Le Cuona.xlsx"),
))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://delecuona.com"


def _parse_dimension_value(text: str, label: str) -> str | None:
    """Extract numeric value for Width/Length/Height from a size string."""
    pattern = rf'{label}\s*[:\s]+([0-9.]+)["“”]'
    m = re.search(pattern, text, re.I)
    if m:
        return m.group(1)
    return None


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a single de Le Cuona product page. Returns one row per product."""
    handle = url.rstrip("/").split("/")[-1]

    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)

    # --- JSON data via .js endpoint ---
    try:
        product_json = await page.evaluate(f"""
            async () => {{
                const r = await fetch('/products/{handle}.js');
                return await r.json();
            }}
        """)
    except Exception:
        product_json = {}

    title = product_json.get("title") or ""

    # Main variant = not "Memo Sample"
    variants = product_json.get("variants") or []
    main_var = next(
        (v for v in variants if "memo sample" not in (v.get("title") or "").lower()),
        variants[0] if variants else {},
    )

    sku = main_var.get("sku") or ""
    price_cents = main_var.get("price") or 0
    price = round(price_cents / 100, 2) if price_cents else None

    # Image URL
    images = product_json.get("images") or []
    var_img = main_var.get("featured_image") or {}
    image_url = ""
    if isinstance(var_img, dict):
        image_url = var_img.get("src") or ""
    if not image_url and images:
        img0 = images[0]
        image_url = img0.get("src") if isinstance(img0, dict) else str(img0)
    if image_url and not image_url.startswith("http"):
        image_url = "https:" + image_url

    # Family / Color from title pattern "Family - Colorway"
    if " - " in title:
        family, color = title.split(" - ", 1)
        family = family.strip()
        color = color.strip()
    else:
        family = title
        color = ""

    # Title case for product name (e.g. "Moya - Rice" not "Moya - rice")
    product_name = title.title()

    row = {
        "Source URL": url,
        "Product Name": product_name,
        "Product Family Id": family,
        "Color": color,
        "SKU": sku,
    }
    if price is not None:
        row["Price"] = price
    if image_url:
        row["Image URL"] = image_url

    # --- HTML scraping for description and extra info ---
    # Description
    desc_el = await page.query_selector(".dlc-description--full")
    if not desc_el:
        desc_el = await page.query_selector(".dlc-description")
    if desc_el:
        raw_desc = clean_text(await desc_el.inner_text())
        raw_desc = re.sub(r"\s*\.\.\.(more|less)\s*", "", raw_desc).strip()
        if raw_desc:
            row["Description"] = raw_desc

    # Size / Use / Availability / Lead Time from dlc-product-extra-info
    extra_el = await page.query_selector("dlc-product-extra-info")
    if extra_el:
        wrappers = await extra_el.query_selector_all(".pdp_extra_info-wrapper")
        for wrapper in wrappers:
            header_el = await wrapper.query_selector("p.h6")
            if not header_el:
                continue
            header = clean_text(await header_el.inner_text()).lower().strip()

            # Collect all <p> text except the h6 heading
            p_texts = []
            for p in await wrapper.query_selector_all("p"):
                cls = (await p.get_attribute("class")) or ""
                if "h6" not in cls:
                    t = clean_text(await p.inner_text())
                    if t:
                        p_texts.append(t)

            combined = " ".join(p_texts)

            if header == "size":
                # "Width: 55\"/ 140cm", "Length: 22\"/ 55cm"
                w = _parse_dimension_value(combined, "Width")
                if w:
                    row["Width"] = w
                l = _parse_dimension_value(combined, "Length")
                if l:
                    row["Length"] = l
                if combined:
                    row["Dimensions"] = combined

            elif header == "use":
                use_text = combined.replace("\xa0", " ").strip()
                if use_text:
                    row["Use"] = use_text

            elif header == "availability":
                # Availability status
                m_avail = re.search(r"\b(Good|Limited|Low Stock|Out of Stock|Discontinued)\b", combined, re.I)
                if m_avail:
                    row["Availability"] = m_avail.group(1)
                # Lead time
                m_lt = re.search(r"Replenishment Lead Time[:\s]+(.+?)\.?\s*$", combined, re.I | re.M)
                if m_lt:
                    row["Lead Time"] = m_lt.group(1).strip()

    return [row]


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from a listing collection, handling pagination."""
    links: list[str] = []
    seen: set[str] = set()
    base_url = listing_url.split("?")[0]
    page_num = 1

    while True:
        url = f"{base_url}?page={page_num}" if page_num > 1 else base_url
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(1000)

        handles: list[str] = await page.evaluate("""
            () => Array.from(document.querySelectorAll('product-card'))
                       .map(c => c.getAttribute('handle'))
                       .filter(Boolean)
        """)

        if not handles:
            break

        new_this_page = 0
        for h in handles:
            if h not in seen:
                seen.add(h)
                links.append(f"{BASE_URL}/products/{h}")
                new_this_page += 1

        print(f"  [Listing] page {page_num}: {len(handles)} cards, {new_this_page} new — total {len(links)}")

        if new_this_page == 0:
            # No new products — reached the end
            break

        page_num += 1

    return links


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    if TEST_MODE:
        print(f"[TEST: max {TEST_MAX_PRODUCTS} products per category]")

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

            # Collect all product URLs across all listing links for this category
            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    base_u = u.split("?")[0]
                    if base_u not in seen_urls:
                        seen_urls.add(base_u)
                        all_product_urls.append(base_u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"[{cat['name']}] {len(all_product_urls)} products to scrape")

            global_idx = 1
            for url in all_product_urls:
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
                    print(f"  ERROR scraping {url}: {e}")

                await async_polite_delay()

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
