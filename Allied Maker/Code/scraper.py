"""
scraper.py  —  Allied Maker
-----------------------------
Platform: www.alliedmaker.com (NetSuite SuiteCommerce or custom CMS)

Site structure:
  Listing  : /Products/custitem_type/{Type}  (Chandelier, Pendant, Sconce, etc.)
             All products on one page, no pagination. ~10-25 products per category.
  Product  : /{Product-Slug}   e.g. /Riva-Trimless, /Bola-4-Chandelier

Product page fields:
  Product Name : h1
  SKU          : h4 or heading below name (e.g. "WAC-100")
  Price        : text "$840.00"
  Dimensions   : DIMENSIONS section → "~3.65" L x 5.5" W x 5.5" H"
  Lamping      : LAMPING section → lamp type and base
  Brightness   : BRIGHTNESS section → lumens
  Finishes     : METAL FINISHES, GLASS FINISHES sections
  Materials    : MATERIALS section
  Description  : paragraph text

Variants: Allied Maker products may have multiple finish options (Metal + Glass).
          One row is written per finish combination with the corresponding SKU suffix.
          If no variants are selectable, one row is written for the base product.

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Allied Maker"
    python orchestrator.py "Allied Maker" --test
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

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
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Allied Maker")
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

BASE_URL   = "https://www.alliedmaker.com"
TIMEOUT_MS = 45_000


# ---------------------------------------------------------------------------
# Listing page — all products on one static page
# ---------------------------------------------------------------------------

async def get_product_links(page, listing_url: str, max_products: int = 0) -> list[str]:
    """
    Collect all product URLs from an Allied Maker category listing.
    All items are on a single page with no pagination.
    Product hrefs follow the pattern: /{Product-Slug}
    """
    print(f"  [Listing] {listing_url}")
    try:
        await page.goto(listing_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"  [WARN] {e}")
        return []

    # All product links on Allied Maker listing pages are simple /{Slug} hrefs.
    # Filter out nav/utility links by requiring the path to have exactly one segment.
    hrefs: list[str] = await page.evaluate(
        f"""() => {{
            const base = "{BASE_URL}";
            const seen = new Set();
            const out  = [];
            document.querySelectorAll("a[href]").forEach(a => {{
                const h = a.href;
                try {{
                    const u = new URL(h);
                    // Same host, single path segment (no sub-paths), not a utility page
                    const parts = u.pathname.replace(/^\\//, "").replace(/\\/$/, "").split("/");
                    if (
                        u.hostname === new URL(base).hostname &&
                        parts.length === 1 &&
                        parts[0].length > 2 &&
                        !["Products", "about", "contact", "cart", "account", "search"].includes(parts[0])
                    ) {{
                        const canonical = base + "/" + parts[0];
                        if (!seen.has(canonical)) {{
                            seen.add(canonical);
                            out.push(canonical);
                        }}
                    }}
                }} catch(e) {{}}
            }});
            return out;
        }}"""
    )

    print(f"  [Listing] {len(hrefs)} product URLs found")
    return hrefs[:max_products] if max_products else hrefs


# ---------------------------------------------------------------------------
# Product detail page
# ---------------------------------------------------------------------------

async def _extract_section_text(page, section_heading: str) -> str:
    """
    Find a section by its heading text (e.g. "DIMENSIONS", "LAMPING")
    and return the text of the element(s) that follow it.
    Allied Maker uses h4/h5 headings followed by <p> or plain text.
    """
    result = await page.evaluate(
        f"""() => {{
            const heading = "{section_heading}".toUpperCase();
            const candidates = [...document.querySelectorAll("h3, h4, h5, strong, b")];
            for (const el of candidates) {{
                if (el.innerText.trim().toUpperCase() === heading) {{
                    // Try next sibling elements for value
                    let sib = el.nextElementSibling;
                    const parts = [];
                    while (sib && !["H3","H4","H5","STRONG","B"].includes(sib.tagName)) {{
                        const txt = sib.innerText.trim();
                        if (txt) parts.push(txt);
                        sib = sib.nextElementSibling;
                        if (parts.length >= 3) break;
                    }}
                    if (parts.length) return parts.join(" | ");
                    // Try parent's next sibling
                    const parent = el.parentElement;
                    if (parent && parent.nextElementSibling) {{
                        return parent.nextElementSibling.innerText.trim();
                    }}
                }}
            }}
            return "";
        }}"""
    )
    return clean_text(result or "")


async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape an Allied Maker product detail page.
    Returns one dict per finish variant, or a single dict if no variants.
    """
    base: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [base]

    body_text = clean_text(await page.evaluate("() => document.body.innerText"))

    # ── 1. Product Name ───────────────────────────────────────────────────
    for sel in ["h1.product-name", "h1.item-name", "h1"]:
        el = await page.query_selector(sel)
        if el:
            text = clean_text(await el.inner_text())
            if text and len(text) < 120:
                base["Product Name"] = text
                break

    # ── 2. SKU (model number) ─────────────────────────────────────────────
    # Allied Maker shows the model number as an h4 right below the h1
    for sel in ["h4.model-number", "h4", ".model-number", ".sku", "[class*='model']"]:
        el = await page.query_selector(sel)
        if el:
            text = clean_text(await el.inner_text())
            # Should look like "WAC-100" — letters/numbers/dash, short
            if text and re.match(r"^[A-Z0-9][\w\-]{1,20}$", text, re.IGNORECASE):
                base["SKU"] = text
                break

    # Fallback: grep body for model-number pattern near the top
    if not base.get("SKU"):
        m = re.search(r"\b([A-Z]{2,5}-\d{2,4}[A-Z]?)\b", body_text)
        if m:
            base["SKU"] = m.group(1)

    # ── 3. Price ──────────────────────────────────────────────────────────
    price_m = re.search(r"\$\s*([\d,]+(?:\.\d{2})?)", body_text)
    if price_m:
        base["Price"] = clean_price(price_m.group(1))

    # ── 4. Image URL ──────────────────────────────────────────────────────
    for sel in [
        ".product-image img",
        ".main-image img",
        ".product-gallery img",
        "img.product-img",
        "img[src*='alliedmaker']",
        "img",
    ]:
        img_el = await page.query_selector(sel)
        if img_el:
            src = (
                await img_el.get_attribute("data-zoom-image")
                or await img_el.get_attribute("data-src")
                or await img_el.get_attribute("src")
                or ""
            )
            if src and "placeholder" not in src and not src.endswith(".gif"):
                if src.startswith("//"):
                    src = "https:" + src
                elif not src.startswith("http"):
                    src = urljoin(BASE_URL, src)
                base["Image URL"] = src
                break

    # ── 5. Dimensions ─────────────────────────────────────────────────────
    dim_text = await _extract_section_text(page, "DIMENSIONS")
    if dim_text:
        parsed = parse_dimensions(dim_text)
        base.update({k: v for k, v in parsed.items() if k not in base})

    # ── 6. Lamping ────────────────────────────────────────────────────────
    lamping = await _extract_section_text(page, "LAMPING")
    if lamping:
        base["Lamping"] = lamping

    # ── 7. Brightness ─────────────────────────────────────────────────────
    brightness = await _extract_section_text(page, "BRIGHTNESS")
    if brightness:
        base["Lumens"] = brightness

    # ── 8. Materials ──────────────────────────────────────────────────────
    materials = await _extract_section_text(page, "MATERIALS")
    if materials:
        base["Materials"] = materials

    # ── 9. Finish options ─────────────────────────────────────────────────
    metal_finishes = await _extract_section_text(page, "METAL FINISHES")
    glass_finishes = await _extract_section_text(page, "GLASS FINISHES")
    other_finishes = await _extract_section_text(page, "FINISHES")

    finish_parts = []
    if metal_finishes:
        finish_parts.append(metal_finishes)
    if glass_finishes:
        finish_parts.append(glass_finishes)
    if other_finishes and not finish_parts:
        finish_parts.append(other_finishes)
    if finish_parts:
        base["Finish"] = " | ".join(finish_parts)

    # ── 10. Description ───────────────────────────────────────────────────
    for sel in [".product-description", ".description", ".item-description", "[class*='desc']"]:
        el = await page.query_selector(sel)
        if el:
            text = clean_text(re.sub(r"<[^>]+>", " ", await el.inner_html()))
            if text and len(text) > 15:
                base["Description"] = text
                break

    # Fallback: first substantial paragraph on the page
    if not base.get("Description"):
        paras = await page.query_selector_all("p")
        for p_el in paras:
            text = clean_text(await p_el.inner_text())
            if text and len(text) > 40 and not text.startswith("$"):
                base["Description"] = text
                break

    # ── 11. Tearsheet / spec sheet ────────────────────────────────────────
    for sel in ["a[href*='tearsheet']", "a[href*='spec']", "a[href*='.pdf']"]:
        el = await page.query_selector(sel)
        if el:
            href = await el.get_attribute("href") or ""
            if href:
                base["Tearsheet Link"] = href if href.startswith("http") else urljoin(BASE_URL, href)
                break

    # ── 12. Product Family Id ─────────────────────────────────────────────
    if not base.get("Product Family Id") and base.get("Product Name"):
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    return [base]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main() -> None:
    info       = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer     = ExcelWriter(OUTPUT_PATH, info["vendor_name"])
    categories = info["categories"]

    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor  : {info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls: set[str] = set()
            all_urls:  list[str] = []

            for listing_url in cat["links"]:
                max_p = TEST_MAX_PRODUCTS if TEST_MODE else 0
                for u in await get_product_links(page, listing_url, max_products=max_p):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_urls.append(u)

            if TEST_MODE:
                all_urls = all_urls[:TEST_MAX_PRODUCTS]

            print(f"\n[Category] {cat['name']}: {len(all_urls)} products")

            for idx, url in enumerate(all_urls, 1):
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        row["Manufacturer"] = info["vendor_name"]
                        writer.write_row(row, category_name=cat["name"])
                    print(f"  [{idx}] {url.rstrip('/').split('/')[-1]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay(0.8, 2.0)

            await async_polite_delay(1.0, 2.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
