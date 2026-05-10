"""
scraper.py  —  Chaddock
------------------------
Platform: chaddock.com (Clarity .NET CMS)

Site structure:
  Listing  : /styles?ProdType={types}&PageIndex=N
             Product hrefs: a[href^='/styles/sku/']
             Pagination   : &PageIndex=N (iterate from 1 until empty)
  Product  : /styles/sku/{SKU}

Product page fields (static HTML, no JS needed):
  Product Name    : div#divStyleName  (text, title-case)
  SKU             : URL slug (last path segment)
  Price           : trade-only / login required (blank)
  Image URL       : first img[src*='/vdir/ImageCabinet/Styles/600x600/'] containing the SKU
  Description     : div immediately following the SectionTitle div
  Dimensions      : #trDimensionsOverall td:nth-child(2) — "W D H (inches)" format
  Weight          : #trWeight td:nth-child(2) — "NNN lb"
  Tearsheet Link  : constructed as /style/{SKU}/TP0
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

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
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Chaddock")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL   = "https://chaddock.com"
TIMEOUT_MS = 45_000


def _strip_page_index(url: str) -> str:
    """Remove &PageIndex=N from URL, returning clean base URL."""
    parsed = urlparse(url)
    params = parse_qs(parsed.query, keep_blank_values=True)
    params.pop("PageIndex", None)
    new_query = urlencode({k: v[0] for k, v in params.items()})
    return urlunparse(parsed._replace(query=new_query))


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from a Chaddock category listing."""
    print(f"  [Listing] {listing_url}")
    all_links: list[str] = []
    seen: set[str] = set()

    base_url = _strip_page_index(listing_url)
    page_num = 1

    while True:
        paged_url = base_url if page_num == 1 else f"{base_url}&PageIndex={page_num}"
        try:
            await page.goto(paged_url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
            await page.wait_for_timeout(1500)
        except Exception as e:
            print(f"  [WARN] page {page_num}: {e}")
            break

        links: list[str] = await page.evaluate(
            """() => {
                const seen = new Set();
                const out  = [];
                document.querySelectorAll("a[href^='/styles/sku/']").forEach(a => {
                    const h = a.href || "";
                    if (!seen.has(h)) { seen.add(h); out.push(h); }
                });
                return out;
            }"""
        )

        new_links = [l for l in links if l not in seen]
        for l in new_links:
            seen.add(l)
        all_links.extend(new_links)

        if not new_links:
            break

        page_num += 1
        await async_polite_delay(0.3, 0.8)

    # Deduplicate preserving order
    seen_final: set[str] = set()
    unique: list[str] = []
    for u in all_links:
        if u not in seen_final:
            seen_final.add(u)
            unique.append(u)

    print(f"  [Listing] {len(unique)} products across {page_num - 1} pages")
    return unique


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a Chaddock product detail page."""
    row: dict = {"Source": url}

    # SKU from URL
    sku = url.rstrip("/").split("/")[-1]
    if sku:
        row["SKU"] = sku

    try:
        await page.goto(url, timeout=TIMEOUT_MS, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)
    except Exception as e:
        print(f"    [WARN] {e}")
        return [row]

    # ── 1. Product Name ────────────────────────────────────────────────────────
    name_el = await page.query_selector("#divStyleName")
    if name_el:
        name_text = clean_text(await name_el.inner_text())
        if name_text:
            row["Product Name"] = name_text.title()

    # ── 2. Description ─────────────────────────────────────────────────────────
    desc = await page.evaluate("""
        () => {
            // Description is the div[style*='margin-top:20px'] inside .grid_4,
            // which comes after the SectionTitle div
            const grid = document.querySelector('.grid_4');
            if (!grid) return '';
            const divs = grid.querySelectorAll("div[style*='margin-top']");
            for (const d of divs) {
                const text = d.textContent.trim();
                if (text.length > 30 && !text.includes('FEATURES') && !text.includes('LOGIN')) {
                    return text;
                }
            }
            return '';
        }
    """)
    if desc:
        row["Description"] = clean_text(desc)

    # ── 3. Dimensions ──────────────────────────────────────────────────────────
    dim_el = await page.query_selector(
        "#ctl00_ctl00_ChildBodyContent_PageContent_trDimensionsOverall td:nth-child(2)"
    )
    if dim_el:
        dim_text = clean_text(await dim_el.inner_text())
        # Format: "66.75 W 86.75 D 50 H (inches)" or "30 W 28 D 32 H"
        dim_clean = re.sub(r'\(inches?\)', '', dim_text, flags=re.IGNORECASE).strip()
        parsed = parse_dimensions(dim_clean)
        row.update({k: v for k, v in parsed.items() if k not in row})

    # ── 4. Weight ──────────────────────────────────────────────────────────────
    weight_el = await page.query_selector(
        "#ctl00_ctl00_ChildBodyContent_PageContent_trWeight td:nth-child(2)"
    )
    if weight_el:
        weight_text = clean_text(await weight_el.inner_text())
        m = re.search(r"([\d.]+)", weight_text)
        if m:
            row["Weight"] = m.group(1)

    # ── 5. Image URL ───────────────────────────────────────────────────────────
    imgs = await page.evaluate(
        f"""() => {{
            const out = [];
            document.querySelectorAll("img[src*='/vdir/ImageCabinet/Styles/600x600/']").forEach(img => {{
                const src = img.getAttribute('src') || '';
                if (src) out.push(src);
            }});
            return out;
        }}"""
    )
    if imgs:
        # Prefer the first image that contains the SKU
        sku = row.get("SKU", "")
        preferred = next((s for s in imgs if sku and sku in s), None) or imgs[0]
        # Strip cache-buster query params
        clean_src = preferred.split("?")[0]
        row["Image URL"] = urljoin(BASE_URL, clean_src)

    # ── 6. Tearsheet Link ──────────────────────────────────────────────────────
    if row.get("SKU"):
        row["Tearsheet Link"] = f"{BASE_URL}/style/{row['SKU']}/TP0"

    # ── 7. Additional dimensions / specs from table ────────────────────────────
    extra_dims = await page.evaluate("""
        () => {
            const result = {};
            document.querySelectorAll('table.DetailTable tr').forEach(tr => {
                const tds = tr.querySelectorAll('td');
                if (tds.length >= 2) {
                    const label = tds[0].textContent.trim().toLowerCase().replace(/:$/, '');
                    const value = tds[1].textContent.trim().replace(/\\s+/g, ' ');
                    if (label && value) result[label] = value;
                }
            });
            return result;
        }
    """)
    if extra_dims:
        for k, v in extra_dims.items():
            if k in ("overall", "dimensions") and not row.get("Width"):
                dim_clean = re.sub(r'\(inches?\)', '', v, flags=re.IGNORECASE).strip()
                parsed = parse_dimensions(dim_clean)
                row.update({k2: v2 for k2, v2 in parsed.items() if k2 not in row})
            elif k == "weight" and not row.get("Weight"):
                m = re.search(r"([\d.]+)", v)
                if m:
                    row["Weight"] = m.group(1)
            elif v and k not in ("overall", "dimensions", "weight"):
                row[k.title()] = v

    # ── 8. Product Family Id ───────────────────────────────────────────────────
    if not row.get("Product Family Id") and row.get("Product Name"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    return [row]


async def main() -> None:
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: all {len(categories)} categories, max {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor : {info['vendor_name']}")
    print(f"[Scraper] Mode   : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output : {OUTPUT_PATH}")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(cat["name"], cat["links"][0], studio_columns=cat["studio_columns"])

            seen_urls: set[str] = set()
            all_urls:  list[str] = []

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
                    print(f"  [{idx}] {url.rstrip('/').split('/')[-1]}")
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")
                await async_polite_delay(0.5, 1.5)

            await async_polite_delay(1.0, 2.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
