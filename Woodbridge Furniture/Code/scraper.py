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
    clean_price,
    generate_sku,
    extract_family_id,
    parse_dimensions,
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Woodbridge Furniture")
HEADLESS = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.woodbridgefurniture.com"


async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Collect all product URLs from a Magento listing page (handles pagination).
    """
    links = []
    seen = set()
    current_url = listing_url

    while True:
        await page.goto(current_url, timeout=60_000, wait_until="networkidle")
        await page.wait_for_timeout(3000)

        # Wait for product grid to render
        try:
            await page.wait_for_selector(".product-item-link", timeout=15_000)
        except Exception:
            pass

        # Collect product links (Woodbridge uses flat slugs, not /products/ paths)
        hrefs = await page.evaluate("""() => {
            const seen = new Set();
            const results = [];
            const selectors = [
                '.product-item-link',
                'a.product-item-photo',
                '.product.name a',
            ];
            for (const sel of selectors) {
                document.querySelectorAll(sel).forEach(el => {
                    // Use el.href for the full absolute URL
                    const href = el.href || el.getAttribute('href');
                    if (href && !seen.has(href)) {
                        seen.add(href);
                        results.push(href);
                    }
                });
            }
            return results;
        }""")

        # Keep product slug URLs; skip category/nav URLs (which contain /products/)
        new_links = [
            h for h in hrefs
            if h not in seen
            and "woodbridgefurniture.com" in h
            and not h.rstrip("/").split("woodbridgefurniture.com", 1)[-1].startswith("/products")
        ]
        for h in new_links:
            seen.add(h)
            links.append(h)

        # Check for next page link
        next_btn = await page.query_selector("a.action.next[rel='next']")
        if not next_btn:
            next_btn = await page.query_selector("li.pages-item-next a")
        if next_btn:
            next_href = await next_btn.get_attribute("href")
            if next_href and next_href.startswith("http") and next_href != current_url:
                current_url = next_href
                await async_polite_delay()
                continue
        break

    return links


async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape a Woodbridge Furniture product detail page.
    Returns list of dicts (one per finish/size variant).
    """
    base = {"Source": url, "Manufacturer": VENDOR_NAME}
    try:
        await page.goto(url, timeout=60_000, wait_until="networkidle")
        await page.wait_for_timeout(3000)

        # Product Name
        for sel in ["h1.page-title .base", "h1.page-title", "h1[itemprop='name']", "h1"]:
            el = await page.query_selector(sel)
            if el:
                base["Product Name"] = clean_text(await el.inner_text())
                break

        # SKU
        for sel in ["[itemprop='sku']", ".product-info-stock-sku .value", ".sku .value"]:
            el = await page.query_selector(sel)
            if el:
                base["SKU"] = clean_text(await el.inner_text())
                break

        # Price
        for sel in ["[data-price-type='finalPrice'] .price",
                    "[itemprop='price']",
                    ".price-box .price",
                    ".product-info-price .price"]:
            el = await page.query_selector(sel)
            if el:
                base["Price"] = clean_price(await el.inner_text())
                break

        # Image URL — use og:image meta (most reliable for Woodbridge)
        og_img = await page.evaluate(
            "document.querySelector('meta[property=\"og:image\"]')?.content"
        )
        if og_img:
            base["Image URL"] = og_img
        else:
            for sel in [".fotorama__img", "img.gallery-placeholder__image",
                        "[data-zoom-image]", ".product-image-photo"]:
                el = await page.query_selector(sel)
                if el:
                    src = await el.get_attribute("data-zoom-image") or await el.get_attribute("src")
                    if src and "placeholder" not in src:
                        base["Image URL"] = src
                        break

        # Description — use og:description meta first
        og_desc = await page.evaluate(
            "document.querySelector('meta[property=\"og:description\"]')?.content"
            " || document.querySelector('meta[name=\"description\"]')?.content"
        )
        if og_desc:
            base["Description"] = clean_text(og_desc)
        else:
            for sel in ["[itemprop='description']", ".product-info-description",
                        ".product.info.detailed"]:
                el = await page.query_selector(sel)
                if el:
                    t = clean_text(await el.inner_text())
                    if t:
                        base["Description"] = t
                        break

        # Extra product info from .product.info.detailed (material, specs)
        detailed = await page.query_selector(".product.info.detailed")
        if detailed:
            detailed_text = clean_text(await detailed.inner_text())
            # Parse material
            m = re.search(r"Material\s*:\s*(.+?)(?:\n|Finish|Spec|$)", detailed_text, re.I)
            if m:
                base["Materials"] = m.group(1).strip()


        # Specifications / attributes (Magento attribute tables)
        spec_data = {}
        spec_rows = await page.query_selector_all(".product-attributes tr, .additional-attributes tr")
        for row in spec_rows:
            th = await row.query_selector("th, td.col.label")
            td = await row.query_selector("td.col.data, td:last-child")
            if th and td:
                key = clean_text(await th.inner_text())
                val = clean_text(await td.inner_text())
                if key and val:
                    spec_data[key] = val

        # Map common Magento attribute names
        attr_map = {
            "Width": ["width", "w"],
            "Height": ["height", "h"],
            "Depth": ["depth", "d"],
            "Diameter": ["diameter", "dia"],
            "Weight": ["weight"],
            "Finish": ["finish", "colour", "color"],
            "Materials": ["material", "materials", "wood species", "wood type"],
            "Collection": ["collection", "series"],
            "Lead Time": ["lead time", "production time"],
            "Origin": ["origin", "country of origin", "made in"],
        }
        for std_key, variants in attr_map.items():
            for spec_key, spec_val in spec_data.items():
                if any(v in spec_key.lower() for v in variants):
                    if std_key in ("Width", "Height", "Depth", "Diameter"):
                        base[std_key] = re.sub(r'["\']', "", spec_val).strip()
                    elif std_key == "Weight":
                        base[std_key] = safe_float(re.sub(r"[^\d.]", "", spec_val))
                    else:
                        base[std_key] = spec_val

        # Store any unmapped specs in Specifications column
        all_specs_text = " | ".join(f"{k}: {v}" for k, v in spec_data.items())
        if all_specs_text:
            base["Specifications"] = all_specs_text

        # If no dimensions from specs, try parsing from description
        if not base.get("Width") and not base.get("Height"):
            dims_text = base.get("Description", "") + " " + base.get("Specifications", "")
            dims = parse_dimensions(dims_text)
            dims.pop("Dimensions", None)
            base.update({k: v for k, v in dims.items() if v})

        # Product Family Id
        if base.get("Product Name"):
            base["Product Family Id"] = extract_family_id(base["Product Name"])

    except Exception as e:
        print(f"    Detail error ({url}): {e}")

    return [base]


async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            cat_url = cat["links"][0]
            writer.add_sheet(cat["name"], cat_url, studio_columns=cat["studio_columns"])

            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"  {cat['name']}: {len(all_product_urls)} products")

            global_idx = 1
            for url in all_product_urls:
                try:
                    variant_rows = await scrape_product(page, url)
                    for row in variant_rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"    ERROR on {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
