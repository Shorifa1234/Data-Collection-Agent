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
    parse_spec_block,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Wesley Hall")
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

BASE_URL = "https://www.wesleyhall.com"

# URL patterns that indicate a textiles/fabric listing page
TEXTILE_URL_PATTERNS = ("/fabrics/", "/fabric/", "/trims", "/leather")


def _is_textile_url(url: str) -> bool:
    return any(p in url for p in TEXTILE_URL_PATTERNS)


async def get_product_links(page, listing_url: str) -> list[tuple[str, str, str, str]]:
    """
    Return list of (product_url, image_url, sku, product_name) from listing page.
    Handles furniture (/styledetail/), search results, and textiles (/fabrics/, /trims) pages.
    Images are taken from listing thumbnails (detail page uses canvas rendering).
    """
    await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(3000)

    if _is_textile_url(listing_url):
        return await _get_textile_links(page, listing_url)
    else:
        return await _get_furniture_links(page, listing_url)


async def _get_furniture_links(page, listing_url: str) -> list[tuple[str, str, str, str]]:
    """Extract product links from furniture / search listing pages."""
    results = []
    seen: set[str] = set()

    # Primary selector: direct styledetail anchors
    # Also covers search result pages which link to the same /styledetail/ URLs
    anchors = await page.query_selector_all("a[href*='/styledetail/']")
    for a in anchors:
        href = await a.get_attribute("href")
        if not href:
            continue
        full_url = BASE_URL + href if href.startswith("/") else href
        if full_url in seen:
            continue
        seen.add(full_url)

        img = await a.query_selector("img")
        img_url = ""
        sku = ""
        name = ""
        if img:
            lazyload = await img.get_attribute("lazyload")
            src = await img.get_attribute("src")
            img_path = lazyload or src or ""
            if img_path:
                img_url = BASE_URL + img_path if img_path.startswith("/") else img_path
                img_url = img_url.split("?")[0]

            alt = await img.get_attribute("alt") or ""
            # alt format: "195-K HYPNOS KING BED 56""
            parts = alt.strip().split(" ", 1)
            if parts:
                sku = parts[0]
                name = parts[1] if len(parts) > 1 else ""

        if not sku:
            m = re.search(r"/id/([^/]+)/", href)
            if m:
                sku = m.group(1)

        results.append((full_url, img_url, sku, name))

    return results


async def _get_textile_links(page, listing_url: str) -> list[tuple[str, str, str, str]]:
    """
    Extract product links from fabric / leather / trim listing pages.
    Wesley Hall textiles use href patterns like /fabric/detail/..., /leather/detail/...,
    /trim/detail/... — fall back to any non-navigation anchor if none are found.
    """
    results = []
    seen: set[str] = set()

    # Try known textile detail URL patterns
    textile_selectors = [
        "a[href*='/fabric/']",
        "a[href*='/leather/']",
        "a[href*='/trim/']",
        "a[href*='/fabrics/detail']",
        "a[href*='/trims/detail']",
    ]
    anchors = []
    for sel in textile_selectors:
        found = await page.query_selector_all(sel)
        if found:
            anchors.extend(found)

    # Generic fallback: all anchors whose href looks like a product detail path
    if not anchors:
        anchors = await page.query_selector_all("a[href]")

    for a in anchors:
        href = await a.get_attribute("href") or ""
        # Skip empty, anchors, external, navigation links
        if not href or href.startswith("#") or href.startswith("javascript"):
            continue
        if href.startswith("http") and BASE_URL not in href:
            continue
        # Skip obvious nav/utility paths
        if any(skip in href for skip in ["/func/", "/search", "/cart", "/account", "/contact"]):
            continue

        full_url = BASE_URL + href if href.startswith("/") else href
        if full_url in seen:
            continue
        seen.add(full_url)

        img = await a.query_selector("img")
        img_url = ""
        sku = ""
        name = ""
        if img:
            for attr in ["lazyload", "data-src", "src"]:
                img_path = await img.get_attribute(attr) or ""
                if img_path and not img_path.startswith("data:"):
                    img_url = BASE_URL + img_path if img_path.startswith("/") else img_path
                    img_url = img_url.split("?")[0]
                    break
            alt = await img.get_attribute("alt") or ""
            parts = alt.strip().split(" ", 1)
            if parts:
                sku = parts[0]
                name = parts[1] if len(parts) > 1 else ""

        # Try to get text name if no image alt
        if not name:
            name = clean_text(await a.inner_text())

        results.append((full_url, img_url, sku, name))

    return results


async def scrape_product(page, url: str, sku: str, listing_image: str, listing_name: str) -> dict:
    """
    Scrape product detail page. Image URL comes from listing page (canvas on detail page).
    For textile pages (/fabrics/, /trims) also captures pattern, repeat, content width, etc.
    """
    data = {
        "Source": url,
        "SKU": sku,
        "Image URL": listing_image,
    }

    try:
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)

        # Try to get a higher-res image from the detail page for textiles
        if _is_textile_url(url) and not listing_image:
            for attr_sel in [
                "img[data-zoom-image]", "img[data-src]",
                ".fabric-image img", ".product-image img", "img.main-image",
            ]:
                img_el = await page.query_selector(attr_sel)
                if img_el:
                    for attr in ["data-zoom-image", "data-src", "src"]:
                        img_path = await img_el.get_attribute(attr) or ""
                        if img_path and not img_path.startswith("data:"):
                            img_url = BASE_URL + img_path if img_path.startswith("/") else img_path
                            data["Image URL"] = img_url.split("?")[0]
                            break
                    if data.get("Image URL"):
                        break

        # Product name — from page heading or breadcrumb
        name = ""
        for sel in ["h1.product-name", "h1.style-name", ".product-title h1",
                    ".style-title", "[class*='product-name']", "h1"]:
            el = await page.query_selector(sel)
            if el:
                name = clean_text(await el.inner_text())
                if name and name.lower() not in ("home", "wesley hall"):
                    break

        data["Product Name"] = name or clean_text(listing_name)

        # Try to find description / spec text
        desc_text = ""
        for sel in [".product-description", ".style-description", ".description",
                    "[class*='description']", ".product-info", ".spec-text",
                    ".product-details p", ".style-info p"]:
            el = await page.query_selector(sel)
            if el:
                t = clean_text(await el.inner_text())
                if t and len(t) > len(desc_text):
                    desc_text = t
        data["Description"] = desc_text

        # Specifications — look for spec sections with measurements
        spec_text = ""
        for sel in [".specifications", ".product-specs", "[class*='spec']",
                    ".dims", ".dimensions", ".product-dimensions"]:
            el = await page.query_selector(sel)
            if el:
                t = clean_text(await el.inner_text())
                if t:
                    spec_text = t
                    break
        if spec_text:
            dims = parse_dimensions(spec_text)
            dims.pop("Dimensions", None)
            data.update(dims)
            data["Specifications"] = spec_text

        # If dimensions not yet found, search targeted areas for dimension patterns
        if not data.get("Width") and not data.get("Height"):
            dim_text = await page.evaluate("""() => {
                const candidates = [
                    ...document.querySelectorAll('table td, dl dt, dl dd, .spec-row'),
                    ...document.querySelectorAll('[class*="dim"], [class*="size"], [class*="measure"]'),
                ];
                return candidates.map(el => el.innerText).join(' | ').slice(0, 2000);
            }""")
            if dim_text:
                dims = parse_dimensions(dim_text)
                data.update({k: v for k, v in dims.items() if v and k != "Dimensions"})

        # COM / fabric info (furniture upholstery)
        for sel in [".com-info", "[class*='com']", ".fabric-info"]:
            el = await page.query_selector(sel)
            if el:
                t = clean_text(await el.inner_text())
                if t:
                    data["COM"] = t[:300]
                    break

        # --- Textile-specific fields (Fabric / Leather / Trim) ---
        if _is_textile_url(url):
            # All detail rows in a structured table or definition list
            detail_rows = await page.evaluate("""() => {
                const rows = {};
                // Try table-based layout
                document.querySelectorAll('tr').forEach(tr => {
                    const cells = tr.querySelectorAll('td, th');
                    if (cells.length >= 2) {
                        const key = cells[0].innerText.trim().replace(/:$/, '');
                        const val = cells[1].innerText.trim();
                        if (key && val) rows[key] = val;
                    }
                });
                // Try dl/dt/dd layout
                const dts = document.querySelectorAll('dl dt');
                dts.forEach(dt => {
                    const dd = dt.nextElementSibling;
                    if (dd && dd.tagName === 'DD') {
                        rows[dt.innerText.trim().replace(/:$/, '')] = dd.innerText.trim();
                    }
                });
                // Try .spec-row or .attribute-row divs
                document.querySelectorAll('[class*="spec-row"], [class*="attribute"]').forEach(el => {
                    const label = el.querySelector('[class*="label"], [class*="key"], strong');
                    const value = el.querySelector('[class*="value"], span:last-child');
                    if (label && value) rows[label.innerText.trim().replace(/:$/, '')] = value.innerText.trim();
                });
                return rows;
            }""")

            TEXTILE_FIELD_MAP = {
                "Content": "Content",
                "Fabric Content": "Content",
                "Content Width": "Width",
                "Usable Width": "Width",
                "Width": "Width",
                "Pattern": "Pattern",
                "Pattern Repeat": "Pattern",
                "Horizontal Repeat": "Horizontal Repeat",
                "H Repeat": "Horizontal Repeat",
                "Vertical Repeat": "Vertical Repeat",
                "V Repeat": "Vertical Repeat",
                "Repeat": "Vertical Repeat",
                "Construction": "Construction",
                "Finish": "Finish",
                "Color": "Color",
                "Colour": "Color",
                "Care Instructions": "Care Instructions",
                "Care": "Care Instructions",
                "Cleaning Code": "Care Instructions",
                "Country of Origin": "Origin",
                "Origin": "Origin",
                "Weight": "Weight",
                "Pile Height": "Pile Height",
                "Thickness": "Thickness",
                "Collection": "Collection",
                "Grade": "Grade",
                "Backing": "Backing",
            }
            for raw_key, raw_val in detail_rows.items():
                mapped = TEXTILE_FIELD_MAP.get(raw_key)
                if mapped:
                    if not data.get(mapped):
                        data[mapped] = clean_text(raw_val)
                else:
                    # Store unmapped fields as-is
                    if not data.get(raw_key):
                        data[raw_key] = clean_text(raw_val)

    except Exception as e:
        print(f"    Detail page error ({url}): {e}")

    # Ensure name is populated
    if not data.get("Product Name"):
        data["Product Name"] = clean_text(listing_name)
    if not data.get("Product Family Id") and data.get("Product Name"):
        data["Product Family Id"] = extract_family_id(data["Product Name"])

    return data


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

            # Collect all product entries from all listing URLs
            seen_urls: set[str] = set()
            all_product_entries: list[tuple] = []
            for listing_url in cat["links"]:
                entries = await get_product_links(page, listing_url)
                for entry in entries:
                    prod_url = entry[0]
                    if prod_url not in seen_urls:
                        seen_urls.add(prod_url)
                        all_product_entries.append(entry)

            if TEST_MODE:
                all_product_entries = all_product_entries[:TEST_MAX_PRODUCTS]
                print(f"  [TEST] {cat['name']}: {len(all_product_entries)} products")
            else:
                print(f"  {cat['name']}: {len(all_product_entries)} products")

            global_idx = 1
            for (prod_url, img_url, sku, listing_name) in all_product_entries:
                try:
                    data = await scrape_product(page, prod_url, sku, img_url, listing_name)
                    if not data.get("SKU"):
                        data["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not data.get("Product Family Id") and data.get("Product Name"):
                        data["Product Family Id"] = extract_family_id(data["Product Name"])
                    writer.write_row(data, category_name=cat["name"])
                    global_idx += 1
                except Exception as e:
                    print(f"    ERROR on {prod_url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
