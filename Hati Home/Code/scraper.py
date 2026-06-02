import asyncio, json, os, sys, re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions,
)

VENDOR_NAME   = os.environ.get("VENDOR_NAME", "Hati Home")
HEADLESS      = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH   = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE     = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CAT  = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PROD = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.hatihome.com"

GET_ACCORDION_TEXT_JS = """
async (heading) => {
    // Hati Home uses Alpine.js: button[aria-controls] with h2 child
    const buttons = document.querySelectorAll('button[aria-controls]');
    for (const btn of buttons) {
        const h2 = btn.querySelector('h2');
        if (!h2) continue;
        if (!h2.textContent.trim().toUpperCase().startsWith(heading.toUpperCase())) continue;
        // Click to expand if not already open
        if (btn.getAttribute('aria-expanded') !== 'true') {
            btn.click();
            await new Promise(r => setTimeout(r, 600));
        }
        const panelId = btn.getAttribute('aria-controls');
        const panel = document.getElementById(panelId);
        return panel ? panel.innerText.trim() : '';
    }
    return '';
}
"""


async def get_accordion_text(page, heading: str) -> str:
    try:
        text = await page.evaluate(GET_ACCORDION_TEXT_JS, heading)
        return clean_text(text or "")
    except Exception:
        return ""


def _parse_dim_line(line: str) -> dict:
    """Parse a single dimension line, stripping seat/arm extras first."""
    line = re.sub(r'\s+\d+(?:\.\d+)?[^\d\sA-Za-z]*\s+SEAT\b.*', '', line, flags=re.IGNORECASE).strip()
    line = re.sub(r'\s+SEAT(?:\s+HEIGHT)?[:\s]+.*', '', line, flags=re.IGNORECASE).strip()
    line = re.sub(r'\s+HEIGHT\s*$', '', line, flags=re.IGNORECASE).strip()
    return parse_dimensions(line) or {}


async def scrape_product(page, url: str) -> tuple:
    base_url = url.split("?")[0]
    await page.goto(base_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)

    handle = base_url.rstrip("/").split("/products/")[-1]

    # Fetch structured data from Shopify .js endpoint
    try:
        pj = await page.evaluate("""
            async (handle) => {
                const res = await fetch('/products/' + handle + '.js');
                return await res.json();
            }
        """, handle)
    except Exception:
        pj = {}

    variants = pj.get("variants", [])
    images = pj.get("images", [])
    options = pj.get("options", [])

    # Is this a real multi-variant product (size/color options, not just "Default Title")?
    is_multi = len(variants) > 1 or (
        len(variants) == 1 and (variants[0].get("title") or "").strip().lower() != "default title"
    )
    option_name = options[0].get("name", "Size") if options else "Size"

    # --- Shared fields (same across all variants) ---
    shared = {"Manufacturer": VENDOR_NAME}

    try:
        h1 = await page.query_selector("h1")
        if h1:
            shared["Product Name"] = clean_text(await h1.inner_text())
    except Exception:
        pass

    # Image — first from images array (shared; variants with their own image handled below)
    if images:
        img = images[0]
        shared["Image URL"] = ("https:" + img) if not img.startswith("http") else img

    # Finish subtitle below h1
    try:
        finish_el = await page.query_selector("p.break-words.font-accent")
        if finish_el:
            ft = clean_text(await finish_el.inner_text())
            if ft:
                shared["Finish"] = ft
    except Exception:
        pass

    # Description
    desc = await get_accordion_text(page, "DETAILS")
    if desc:
        shared["Description"] = desc

    # Materials
    mat_raw = await get_accordion_text(page, "MATERIALS")
    if mat_raw:
        shared["Materials"] = mat_raw

    # Product Family Id
    if shared.get("Product Name"):
        shared["Product Family Id"] = extract_family_id(shared["Product Name"])

    # --- Dimensions ---
    dim_raw = await get_accordion_text(page, "DIMENSIONS")
    dim_lines = [l.strip() for l in re.split(r"[\n•]", dim_raw) if l.strip()] if dim_raw else []
    # For tables with multiple "ACCOMMODATES X DINING CHAIRS:" entries on one line, split by keyword
    if dim_raw and dim_raw.upper().count("ACCOMMODATES") > 1 and len(dim_lines) == 1:
        parts = re.split(r'(?i)(?=ACCOMMODATES\b)', dim_raw)
        dim_lines = [p.strip() for p in parts if p.strip()]

    # Seat Height and Arm Height (shared across all variants)
    if dim_raw:
        for seat_pat in [
            r'SEAT\s+HEIGHT[:\s]+(\d+(?:\.\d+)?)',
            r'(\d+(?:\.\d+)?)[^\d\sA-Za-z]*\s+SEAT\b',
            r'SEAT[:\s]+(\d+(?:\.\d+)?)',
        ]:
            seat_m = re.search(seat_pat, dim_raw, re.IGNORECASE)
            if seat_m:
                shared["Seat Height"] = seat_m.group(1)
                break
        arm_m = re.search(r'ARM(?:\s+HEIGHT)?[:\s]+(\d+(?:\.\d+)?)', dim_raw, re.IGNORECASE)
        if arm_m:
            shared["Arm Height"] = arm_m.group(1)

    # --- Swatch variant URLs ---
    try:
        swatch_hrefs = await page.evaluate("""
            () => {
                const swatchDiv = document.querySelector('div.flex.flex-wrap.gap-2');
                if (!swatchDiv) return [];
                const anchors = swatchDiv.querySelectorAll('a[href*="/products/"]');
                return [...anchors].map(a => a.getAttribute('href'));
            }
        """)
        variant_urls = []
        for h in swatch_hrefs:
            if h and "/products/" in h:
                clean_h = h.split("?")[0]
                full = (BASE_URL + clean_h) if clean_h.startswith("/") else clean_h
                variant_urls.append(full)
    except Exception:
        variant_urls = []

    # --- Build rows ---
    rows = []

    if is_multi:
        # One row per Shopify variant (e.g. two sizes of a dining table)
        for v in variants:
            row = dict(shared)
            row["SKU"] = (v.get("sku") or "").strip()
            row["Price"] = round((v.get("price") or 0) / 100, 2)
            row["Source URL"] = f"{base_url}?variant={v['id']}"
            opt_val = (v.get("option1") or "").strip()
            if opt_val and opt_val.lower() != "default title":
                row[option_name] = opt_val  # "Size" → "82.75\""

            # Per-variant image
            vi = v.get("featured_image")
            if vi:
                src = vi.get("src") if isinstance(vi, dict) else vi
                if src:
                    row["Image URL"] = ("https:" + src) if not src.startswith("http") else src

            # Match dimension line by looking for the size number in the line
            size_num = re.sub(r'[^\d.]', '', opt_val)  # "82.75" from "82.75\""
            matched_line = None
            for line in dim_lines:
                if size_num and size_num in re.sub(r'[^\d.]', ' ', line):
                    matched_line = line
                    break
            if not matched_line and dim_lines:
                matched_line = dim_lines[0]  # fallback to first line

            if matched_line:
                row["Specifications"] = matched_line
                parsed = _parse_dim_line(matched_line)
                row.update({k: pv for k, pv in parsed.items() if pv and not row.get(k)})

            rows.append(row)

    else:
        # Single variant product
        v = variants[0] if variants else {}
        row = dict(shared)
        row["Source URL"] = base_url
        sku = (v.get("sku") or "").strip()
        if sku and sku.lower() != "default title":
            row["SKU"] = sku
        price_cents = v.get("price") or 0
        if price_cents:
            row["Price"] = round(price_cents / 100, 2)

        # Per-variant image override
        vi = v.get("featured_image")
        if vi:
            src = vi.get("src") if isinstance(vi, dict) else vi
            if src:
                row["Image URL"] = ("https:" + src) if not src.startswith("http") else src

        if dim_lines:
            main_dim = dim_lines[0]
            row["Specifications"] = dim_raw
            parsed = _parse_dim_line(main_dim)
            row.update({k: pv for k, pv in parsed.items() if pv and not row.get(k)})

        rows.append(row)

    return rows, variant_urls


def _collect_product_hrefs(hrefs: list, seen: set, links: list):
    for h in hrefs:
        if not h or "/products/" not in h:
            continue
        clean = h.split("?")[0]
        if clean.strip("/") == "products":
            continue
        full = (BASE_URL + clean) if clean.startswith("/") else clean
        if full not in seen:
            seen.add(full)
            links.append(full)


async def get_product_links(page, listing_url: str) -> list[str]:
    links = []
    seen: set = set()
    current_url = listing_url

    while True:
        await page.goto(current_url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(1500)

        hrefs = await page.evaluate("""
            () => {
                const anchors = document.querySelectorAll('a[href*="/products/"]');
                return [...new Set([...anchors].map(a => a.getAttribute('href')))];
            }
        """)
        _collect_product_hrefs(hrefs, seen, links)

        # Shopify pagination: ?page=N
        next_href = await page.evaluate("""
            () => {
                const el = document.querySelector('a[href*="?page="], a.next, a[rel="next"]');
                return el ? el.getAttribute('href') : null;
            }
        """)
        if not next_href:
            break
        current_url = (BASE_URL + next_href) if next_href.startswith("/") else next_href

    return links


async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CAT]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls: set = set()
            queue: list = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        queue.append(u)

            if TEST_MODE:
                queue = queue[:TEST_MAX_PROD]

            print(f"  [{cat['name']}] {len(queue)} products found")

            global_idx = 1
            idx = 0
            while idx < len(queue):
                url = queue[idx]
                idx += 1
                try:
                    rows, variant_urls = await scrape_product(page, url)
                    # Add any new swatch variants discovered on this page
                    if not TEST_MODE:
                        for vu in variant_urls:
                            if vu not in seen_urls:
                                seen_urls.add(vu)
                                queue.append(vu)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [!] Error scraping {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"[Done] Saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
