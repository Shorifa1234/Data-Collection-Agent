import asyncio, json, os, re, sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, clean_price,
    generate_sku, extract_family_id, parse_dimensions,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Dimitriy & Co.")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE          = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_PRODUCTS  = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))

BASE_URL = "https://www.dmitriyco.com"


# ---------------------------------------------------------------------------
# Quote state parsing
# ---------------------------------------------------------------------------

async def get_quote_state(page) -> dict:
    """Parse right-panel quote summary into a flat dict {label: {value, price}}."""
    lines = await page.evaluate(r"""() => {
        const cells = Array.from(document.querySelectorAll('[class*="product-quote-row"]'));
        return cells.map(c => c.innerText.trim()).filter(s => s.length > 0);
    }""")

    result = {}
    skip_labels = {"QUANTITY", "VIEW NET PRICING", "VIEW SLEEPING SIZE AND DIMENSIONS"}

    for line in lines:
        if "\n" not in line:
            continue  # standalone items (price chips, buttons) — skip

        parts = [p.strip() for p in re.split(r"[\n\t]+", line) if p.strip()]
        if not parts:
            continue

        label = parts[0]
        if label in skip_labels or label in result:
            continue

        values = []
        price = None
        for part in parts[1:]:
            if re.match(r"^\+?\$[\d,]+$", part):
                price = part.lstrip("+")
            else:
                values.append(part)

        result[label] = {"value": " ".join(values), "price": price}

    return result


# ---------------------------------------------------------------------------
# Product detail scraping
# ---------------------------------------------------------------------------

async def scrape_product(page, url: str) -> list[dict]:
    """Return a list of dicts — one per size variant (or one if no size selection)."""
    await page.goto(url, timeout=45_000, wait_until="networkidle")
    await page.wait_for_timeout(2000)

    # --- Base data from JSON-LD ---
    base = {"Source URL": url}

    json_ld = await page.evaluate(r"""() => {
        for (const s of document.querySelectorAll('script[type="application/ld+json"]')) {
            try {
                const obj = JSON.parse(s.textContent);
                if (obj['@type'] === 'Product') return obj;
            } catch(e) {}
        }
        return null;
    }""")

    if json_ld:
        base["Product Name"] = clean_text(json_ld.get("name", ""))
        base["Image URL"] = json_ld.get("image", "")
        desc = clean_text(json_ld.get("description", ""))
        if desc:
            base["Description"] = desc

    # Fallback product name from h1
    if not base.get("Product Name"):
        h1 = await page.query_selector("h1")
        if h1:
            base["Product Name"] = clean_text(await h1.inner_text())

    base["Manufacturer"] = VENDOR_NAME

    # Skip staging/placeholder products
    name = base.get("Product Name", "")
    if not name or "[" in name or "PRODUCT TITLE" in name.upper():
        print(f"  [SKIP] Placeholder product: {url}")
        return []

    # Tearsheet PDF link
    tearsheet = await page.evaluate(r"""() => {
        const a = document.querySelector(
            'a[href*="TEARSHEET"], a[href*="tearsheet"], a[href*=".pdf"]'
        );
        return a ? a.href : null;
    }""")
    if tearsheet:
        base["Tearsheet Link"] = tearsheet

    # --- Detect variant dimension options ---
    first_option_title = await page.evaluate(r"""() => {
        const el = document.querySelector(
            '#option-type-0 .lined-title span, #option-type-0 .construction-row-title span'
        );
        return el ? el.innerText.trim().toUpperCase() : null;
    }""")

    # DIMENSION means modular/sectional depth (not a true size variant) — treat as fixed
    has_size_variants = first_option_title and any(
        kw in first_option_title for kw in ("SIZE", "FRAME", "BED")
    ) and "DIMENSION" not in first_option_title

    # --- Collect all finish/wood options as a combined string ---
    all_finish_options = await page.evaluate(r"""() => {
        const results = [];
        for (const container of document.querySelectorAll('.product-option-container')) {
            const titleEl = container.querySelector('.lined-title span, .construction-row-title span');
            if (!titleEl) continue;
            const title = titleEl.innerText.trim().toUpperCase();
            if (title.includes('FINISH') || title.includes('WOOD')) {
                // Only desktop options (avoid mobile duplicates)
                const desktopContainer = container.querySelector('.desktop-only .options-container, .desktop-only');
                const source = desktopContainer || container;
                const opts = source.querySelectorAll('.selection-details-title');
                opts.forEach(o => {
                    const t = o.innerText.trim();
                    if (t) results.push(t);
                });
            }
        }
        return results;
    }""")
    # Deduplicate finish options while preserving order
    all_finish_options = list(dict.fromkeys(all_finish_options))

    variants = []

    if has_size_variants:
        # --- BED FRAME SIZE / SIZE products: one row per option ---
        size_options = await page.query_selector_all(
            "#option-type-0 .selection.text-option"
        )

        for opt in size_options:
            size_name = clean_text(await opt.inner_text())
            await opt.click()
            await page.wait_for_timeout(800)

            quote = await get_quote_state(page)

            row = dict(base)

            # Map option title to column name
            if first_option_title and "BED" in first_option_title:
                row["Bed Frame Size"] = size_name
            else:
                row["Size"] = size_name

            # Dimensions + price from SIZE quote line
            size_q = quote.get("SIZE", {})
            if size_q.get("value"):
                row["Dimensions"] = size_q["value"].replace('"', "").strip()
            if size_q.get("price"):
                row["Price"] = clean_price(size_q["price"])

            _apply_quote_metadata(row, quote, all_finish_options)
            variants.append(row)

    else:
        # --- Fixed-size product: single row ---
        row = dict(base)

        # Dimensions from second .product-detail-customization (right panel)
        dims_els = await page.query_selector_all(".product-detail-customization")
        for el in dims_els:
            text = clean_text(await el.inner_text())
            # The one that matches dimension pattern (W / D / H etc.)
            if re.search(r"[WwDdHhLl]\s*[\d.]+", text):
                row["Dimensions"] = text.replace('"', "").strip()
                break

        # Price from .option-value-price
        price_el = await page.query_selector(".option-value-price")
        if price_el:
            row["Price"] = clean_price(await price_el.inner_text())

        quote = await get_quote_state(page)
        _apply_quote_metadata(row, quote, all_finish_options)
        variants.append(row)

    return variants


def _apply_quote_metadata(row: dict, quote: dict, finish_options: list) -> None:
    """Fill in shared metadata from quote state."""
    # COM Required Yardage
    com_q = quote.get("COM REQUIRED YARDAGE", {})
    if com_q.get("value"):
        row["COM Yardage"] = com_q["value"]

    # COM flag
    fabric_q = quote.get("FABRIC", {})
    if fabric_q.get("value"):
        fab = fabric_q["value"]
        row["Fabric"] = fab
        if "COM" in fab.upper():
            row["COM"] = "Yes"

    # Lead Time
    lt_q = quote.get("LEAD TIME", {})
    if lt_q.get("value"):
        row["Lead Time"] = lt_q["value"]

    # Construction
    const_q = quote.get("CONSTRUCTION", {})
    if const_q.get("value"):
        row["Construction"] = const_q["value"]

    # Layout (sectionals)
    layout_q = quote.get("LAYOUT", {})
    if layout_q.get("value"):
        row["Layout"] = layout_q["value"]

    # Finish (currently selected from quote, all options as semicolon list)
    finish_keys = [k for k in quote if "FINISH" in k or "WOOD" in k or "BRONZE" in k]
    for fk in finish_keys:
        selected = quote[fk].get("value", "")
        col_name = "Finish" if "FINISH" in fk or "BRONZE" in fk else "Wood Finish"
        if selected:
            row[col_name] = selected
        # Also store all available options
        if finish_options:
            row[f"{col_name} Options"] = "; ".join(finish_options)

    # Architectural Bronze Finish (chairs)
    if "ARCHITECTURAL BRONZE FINISH" in quote:
        row.setdefault("Finish", quote["ARCHITECTURAL BRONZE FINISH"].get("value", ""))


# ---------------------------------------------------------------------------
# Listing pages
# ---------------------------------------------------------------------------

async def get_product_links(page, listing_url: str) -> list[str]:
    """Return deduplicated product URLs from a collection page."""
    await page.goto(listing_url, timeout=45_000, wait_until="networkidle")
    await page.wait_for_timeout(1500)

    hrefs = await page.evaluate(r"""() => {
        const anchors = Array.from(document.querySelectorAll('a[href*="/products/"]'));
        return [...new Set(anchors.map(a => a.href.split('?')[0]))];
    }""")

    links = []
    seen = set()
    for h in hrefs:
        if h not in seen and "/products/" in h:
            seen.add(h)
            links.append(h)
    return links


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(cat["name"], cat["links"][0],
                             studio_columns=cat["studio_columns"])

            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            global_idx = 1
            for url in all_product_urls:
                try:
                    variant_rows = await scrape_product(page, url)
                    for variant in variant_rows:
                        if not variant.get("SKU"):
                            variant["SKU"] = generate_sku(
                                info["vendor_name"], cat["name"], global_idx
                            )
                        if not variant.get("Product Family Id") and variant.get("Product Name"):
                            variant["Product Family Id"] = extract_family_id(
                                variant["Product Name"]
                            )
                        writer.write_row(variant, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}")

                await async_polite_delay()

    writer.save()
    print(f"[Done] Saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
