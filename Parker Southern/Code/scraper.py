"""
scraper.py  —  Parker Southern
--------------------------------
Platform: parkersouthern.com (custom ASP.NET CMS, public, no JS required)

Site structure:
  Listing  : /products/func/cat/{id}    → all products on one page, no pagination
             /fabrics/func/cat/{id}     → all fabrics on one page
             /trims/func/cat/{id}       → all trims on one page
  Product  : /productDetail/func/id/{id}
  Fabric   : /fabricDetail/func/id/{id} (or similar)
  High-res : /downloadit/ps/style/{sku}  (direct download, no auth needed)

No JSON-LD — all data parsed from structured HTML sections:
  Product name  : first <h2> or <h3> in header area
  SKU           : "PRODUCT NUMBER" label text
  Dimensions    : "DIMENSIONS" section → Width / Depth / Height / Seat Height / Seat Depth / Arm Height
  Description   : "DESCRIPTION" section text
  Series        : "SERIES" label
  Finish shown  : "FINISH SHOWN" label
  Fabric shown  : "FABRIC SHOWN" label
  COM Available : derived from whether COM fabric is mentioned
  Lead Time     : "LEAD TIME" label if present

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Parker Southern"
    python orchestrator.py "Parker Southern" --test
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
    generate_sku,
    extract_family_id,
    parse_dimensions,
    clean_price,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Parker Southern")
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

BASE_URL   = "https://www.parkersouthern.com"
TIMEOUT_MS = 45_000


# ---------------------------------------------------------------------------
# Listing page — all products on one page, no pagination
# ---------------------------------------------------------------------------

async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Collect all product URLs from a Parker Southern listing page.
    All items appear on a single page — no pagination needed.
    Works for /products/func/cat/{id}, /fabrics/func/cat/{id}, /trims/func/cat/{id}.
    """
    print(f"  [Listing] {listing_url}")
    try:
        await page.goto(listing_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)
    except Exception as e:
        print(f"  [WARN] Failed to load: {e}")
        return []

    # Collect all anchor hrefs that point to detail pages
    hrefs: list[str] = await page.evaluate(
        """() => {
            const links = [];
            document.querySelectorAll("a[href]").forEach(a => {
                const h = a.getAttribute("href") || "";
                if (
                    h.includes("/productDetail/func/id/") ||
                    h.includes("/fabricDetail/func/id/") ||
                    h.includes("/trimDetail/func/id/")
                ) links.push(a.href);
            });
            return links;
        }"""
    )

    seen: set[str] = set()
    links: list[str] = []
    for h in hrefs:
        if h and h not in seen:
            seen.add(h)
            links.append(h)

    print(f"  [Listing] {len(links)} product URLs found")
    return links


# ---------------------------------------------------------------------------
# Product detail page scraping
# ---------------------------------------------------------------------------

def _sku_from_url(url: str) -> str:
    """Extract the numeric id from /productDetail/func/id/{id}."""
    m = re.search(r"/(?:productDetail|fabricDetail|trimDetail)/func/id/(\d+)", url)
    return m.group(1) if m else ""


async def _get_section_text(page, section_label: str) -> str:
    """
    Find text near a label like 'PRODUCT NUMBER', 'DIMENSIONS', 'DESCRIPTION'.
    Parker Southern uses simple text nodes and <td> or <div> pairs.
    """
    # Try to find a table cell/div that follows a cell/div containing the label
    raw = await page.evaluate(
        f"""() => {{
            const label = "{section_label}".toUpperCase();
            // Check all table cells
            const cells = [...document.querySelectorAll("td, th, div, span, p, li")];
            for (let i = 0; i < cells.length; i++) {{
                const txt = (cells[i].innerText || "").trim().toUpperCase();
                if (txt === label && cells[i + 1]) {{
                    return (cells[i + 1].innerText || "").trim();
                }}
                // Label and value in same cell separated by newline or colon
                if (txt.startsWith(label + "\\n") || txt.startsWith(label + ":")) {{
                    return txt.replace(label, "").replace(":", "").trim();
                }}
            }}
            return "";
        }}"""
    )
    return clean_text(raw or "")


async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape a Parker Southern product detail page.
    Returns a list with a single dict (no variant rows — Parker Southern
    shows one finish/fabric combo per page).
    """
    base: dict = {"Source": url}

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)
    except Exception as e:
        print(f"    [WARN] Could not load: {e}")
        return [base]

    # ── Full page text for regex parsing ──────────────────────────────────
    body_text = clean_text(await page.evaluate("() => document.body.innerText"))

    # ── 1. Product Name ───────────────────────────────────────────────────
    for sel in ["h1", "h2.product-name", "h2", ".product-title", ".item-name"]:
        el = await page.query_selector(sel)
        if el:
            text = clean_text(await el.inner_text())
            if text and len(text) < 120:
                base["Product Name"] = text
                break

    # ── 2. SKU ────────────────────────────────────────────────────────────
    # Pattern: "PRODUCT NUMBER 2012-AL" or "Model: 2012-AL"
    sku_m = re.search(
        r"PRODUCT\s+NUMBER\s+([A-Z0-9\-]+)",
        body_text,
        re.IGNORECASE,
    )
    if sku_m:
        base["SKU"] = sku_m.group(1).strip()
    else:
        # Try CSS selectors
        for sel in [".product-number", ".sku", ".model", "[class*='sku']", "[class*='model']"]:
            el = await page.query_selector(sel)
            if el:
                text = re.sub(r"(?i)(product\s*number|model|sku)\s*[:#]?\s*", "", clean_text(await el.inner_text())).strip()
                if text:
                    base["SKU"] = text
                    break

    # ── 3. Image URL — high-res download ─────────────────────────────────
    # Parker Southern provides a high-res download at /downloadit/ps/style/{sku}
    sku = base.get("SKU", "")
    if sku:
        base["Image URL"] = f"{BASE_URL}/downloadit/ps/style/{sku}"
    else:
        # Fallback: find largest image on page
        for sel in [
            "img[src*='/products/']",
            "img[src*='/images/']",
            ".product-image img",
            "img.main-image",
            "img",
        ]:
            img_el = await page.query_selector(sel)
            if img_el:
                src = await img_el.get_attribute("src") or ""
                if src and not src.endswith(".gif") and "placeholder" not in src:
                    base["Image URL"] = urljoin(BASE_URL, src)
                    break

    # ── 4. Price ──────────────────────────────────────────────────────────
    for sel in [".price", ".list-price", ".product-price", "[class*='price']"]:
        el = await page.query_selector(sel)
        if el:
            price = clean_price(clean_text(await el.inner_text()))
            if price:
                base["Price"] = price
                break

    # ── 5. Description ────────────────────────────────────────────────────
    # Typically labelled "DESCRIPTION" in a section
    desc_m = re.search(
        r"DESCRIPTION\s*\n?\s*(.+?)(?=\n[A-Z ]{3,}\s*\n|\Z)",
        body_text,
        re.IGNORECASE | re.DOTALL,
    )
    if desc_m:
        desc = clean_text(desc_m.group(1))
        if desc and len(desc) < 2000:
            base["Description"] = desc

    # ── 6. Series ─────────────────────────────────────────────────────────
    series_m = re.search(r"SERIES\s*\n?\s*(.+?)(?=\n)", body_text, re.IGNORECASE)
    if series_m:
        base["Collection"] = clean_text(series_m.group(1))

    # ── 7. Finish & Fabric shown ──────────────────────────────────────────
    finish_m = re.search(r"FINISH\s+SHOWN\s*\n?\s*(.+?)(?=\n)", body_text, re.IGNORECASE)
    if finish_m:
        base["Finish"] = clean_text(finish_m.group(1))

    fabric_m = re.search(r"FABRIC\s+SHOWN\s*\n?\s*(.+?)(?=\n)", body_text, re.IGNORECASE)
    if fabric_m:
        base["Fabric"] = clean_text(fabric_m.group(1))

    # ── 8. Dimensions ─────────────────────────────────────────────────────
    # Pattern: "Width 20.5"  or  "Width: 20.5""
    dim_patterns = {
        "Width":      r'Width\s*[:\s]\s*([\d./ ]+)"?',
        "Depth":      r'Depth\s*[:\s]\s*([\d./ ]+)"?',
        "Height":     r'Height\s*[:\s]\s*([\d./ ]+)"?',
        "Seat Height":r'Seat\s+Height\s*[:\s]\s*([\d./ ]+)"?',
        "Seat Depth": r'Seat\s+Depth\s*[:\s]\s*([\d./ ]+)"?',
        "Arm Height": r'Arm\s+Height\s*[:\s]\s*([\d./ ]+)"?',
        "Diameter":   r'Dia(?:meter)?\s*[:\s]\s*([\d./ ]+)"?',
        "Length":     r'Length\s*[:\s]\s*([\d./ ]+)"?',
        "Weight":     r'Weight\s*[:\s]\s*([\d./ ]+)\s*(?:lbs?)?',
    }
    for field, pattern in dim_patterns.items():
        m = re.search(pattern, body_text, re.IGNORECASE)
        if m:
            val = m.group(1).strip().rstrip('"').strip()
            if val:
                base[field] = val

    # Also check textile-specific fields
    content_m = re.search(r"CONTENT\s*\n?\s*(.+?)(?=\n)", body_text, re.IGNORECASE)
    if content_m:
        base["Materials"] = clean_text(content_m.group(1))

    repeat_m = re.search(r"REPEAT\s*\n?\s*(.+?)(?=\n)", body_text, re.IGNORECASE)
    if repeat_m:
        base["Repeat"] = clean_text(repeat_m.group(1))

    lead_m = re.search(r"LEAD\s+TIME\s*\n?\s*(.+?)(?=\n)", body_text, re.IGNORECASE)
    if lead_m:
        base["Lead Time"] = clean_text(lead_m.group(1))

    # ── 9. COM availability ───────────────────────────────────────────────
    if re.search(r"\bCOM\b", body_text, re.IGNORECASE):
        base["COM Available"] = "Yes"

    # ── 10. Tearsheet ─────────────────────────────────────────────────────
    for sel in ["a[href*='tearsheet']", "a[href*='.pdf']", "a[href*='download']"]:
        el = await page.query_selector(sel)
        if el:
            href = await el.get_attribute("href") or ""
            if href:
                base["Tearsheet Link"] = urljoin(BASE_URL, href)
                break

    # ── 11. Product Family Id ─────────────────────────────────────────────
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
            all_urls: list[str] = []

            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
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
                    print(f"  [{idx}] {url.split('/')[-1] or url.split('/')[-2]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay(0.5, 1.5)

            await async_polite_delay(1.0, 2.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
