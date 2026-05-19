import asyncio, io, json, os, sys, re
from collections import deque
from pathlib import Path

import requests

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Rowe")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://rowefurniture.com"

# Labels from spec tables → normalised field names
LABEL_MAP = {
    "weight (lb)": "Weight",
    "dimensions (in)": "Dimensions",
    "length (in)": "Length",
    "width (in)": "Width",
    "depth (in)": "Depth",
    "height (in)": "Height",
    "diameter (in)": "Diameter",
    "seat height (in)": "Seat Height",
    "seat depth (in)": "Seat Depth",
    "arm height (in)": "Arm Height",
    "distance between arms (in)": "Distance Between Arms",
    "country of origin": "Origin",
    "collection": "Collection",
    "kd construction": "KD Construction",
    "adjustable floor glides": "Adjustable Floor Glides",
    "number of cushions": "Cushion",
    "number of back pillows": "Back Pillows",
    "standard cushion": "Cushion Fill",
    "allowed patterns": "Allowed Patterns",
    "construction": "Construction",
    "finish": "Finish",
    "material": "Material",
    "fabric / grade": "Grade",
    "grade": "Grade",
    "pattern restriction": "Pattern Restriction",
    "primary color": "Primary Color",
    "secondary color": "Secondary Color",
    "fabric pattern": "Pattern Type",
    "fabric backing": "Fabric Backing",
    "wellness": "Wellness",
    "performance fabric": "Performance",
    "cleaning code": "Cleaning Code",
    "color": "Color",
    "wearability": "Wearability",
    "content": "Content",
    "unbalanced": "Unbalanced",
    "swatch number": "Swatch Number",
    "rub count": "Rub Count",
    "railroaded status": "Railroaded",
    "color code": "Color Code",
    "dye lot": "Dye Lot",
    "sku": "SKU",
    # Leather-specific
    "leather name": "Leather Name",
    "leather type": "Leather Type",
    "kid proof leather": "Kid Proof Leather",
    "leather grade": "Leather Grade",
    "leather color": "Color",
    "tannery": "Tannery",
    "temper": "Temper",
    "thickness": "Thickness",
    "hide size": "Hide Size",
    "average hide size": "Hide Size",
}


def normalise_label(raw: str) -> str:
    return LABEL_MAP.get(raw.lower().strip(), raw.strip())


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product-card URLs from a listing page (handles pagination).

    Preserves the URL hash fragment so JS-rendered category grids load correctly.
    Pagination inserts ?pagenumber=N BEFORE the hash so the JS router stays intact.
    """
    links: list[str] = []
    seen: set[str] = set()
    page_num = 1

    # Separate base path from hash (keep hash for JS-rendered grids)
    if "#" in listing_url:
        base_no_hash = listing_url.split("#")[0].split("?")[0]
        hash_part = "#" + listing_url.split("#", 1)[1]
    else:
        base_no_hash = listing_url.split("?")[0]
        hash_part = ""

    while True:
        if page_num == 1:
            url = base_no_hash + hash_part
        else:
            url = f"{base_no_hash}?pagenumber={page_num}" + hash_part
        try:
            await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        except Exception:
            break
        await page.wait_for_timeout(2500)

        # Product cards always contain rffblob images inside their anchor
        page_links: list[str] = await page.evaluate("""
        () => {
            const found = [];
            for (const a of document.querySelectorAll('a[href]')) {
                const href = a.getAttribute('href');
                if (!href || !href.startsWith('/')) continue;
                if (a.querySelector('img[src*="rffblob"]')) {
                    found.push(href);
                }
            }
            return [...new Set(found)];
        }
        """)

        new_count = 0
        for href in page_links:
            full = BASE_URL + href
            if full not in seen:
                seen.add(full)
                links.append(full)
                new_count += 1

        if new_count == 0:
            break

        # Check for a "Next" pagination link
        has_next: bool = await page.evaluate("""
        () => {
            for (const a of document.querySelectorAll('a')) {
                const t = a.innerText.trim();
                if (t === 'Next' || t === '›' || t === '>') return true;
            }
            return false;
        }
        """)
        if not has_next:
            break
        page_num += 1

    return links


async def _expand_specs(page) -> None:
    """Expand the Specifications accordion if it appears collapsed.

    Rowe furniture pages load with specs already expanded (no "+" indicator).
    Leather/fabric pages show "Specifications +" in collapsed state — detect via
    the "+" substring and click only when collapsed, to avoid toggling expanded specs.
    """
    try:
        needs_click: bool = await page.evaluate("""
        () => {
            // Look for any element whose visible text is "Specifications" followed by a "+"
            for (const el of document.querySelectorAll('a, button, h2, h3, h4, span, div, p')) {
                const txt = (el.innerText || '').trim();
                if (/^Specifications\\s*\\+/.test(txt)) return true;  // collapsed indicator
                if (txt === 'Specifications') {
                    // Check aria-expanded or class-based collapsed state
                    const exp = el.getAttribute('aria-expanded');
                    if (exp === 'false') return true;
                    if ((el.className || '').includes('collapse')) return true;
                }
            }
            return false;
        }
        """)
        if needs_click:
            toggle = await page.query_selector("text=Specifications")
            if toggle:
                await toggle.click()
                await page.wait_for_timeout(1200)
            return

        # Fallback: no collapsed indicator found, but also no tables → try clicking
        table_count: int = await page.evaluate(
            "() => document.querySelectorAll('table').length"
        )
        if table_count == 0:
            toggle = await page.query_selector("text=Specifications")
            if toggle:
                await toggle.click()
                await page.wait_for_timeout(1200)
    except Exception:
        pass


def _fetch_tearsheet_data(pdf_url: str) -> dict:
    """Download a Salsify tearsheet PDF and extract fields as a dict.

    Typical tearsheet fields captured:
      Dimensions(in), Dimensions(cm), Finish, Materials, Features,
      Designer, Collection, Lead Time, and any other label/value pairs.

    Returns an empty dict on any error.
    """
    if not pdf_url:
        return {}
    try:
        import pdfplumber
        resp = requests.get(pdf_url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        data: dict = {}
        with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
                i = 0
                while i < len(lines):
                    line = lines[i]
                    # Common tearsheet label patterns (bold labels on their own line)
                    if line in (
                        "Dimensions(in)", "Dimensions (in)",
                        "Dimensions(cm)", "Dimensions (cm)",
                        "Finish", "Finishes", "Material", "Materials",
                        "Features", "Feature", "Designer", "Collection",
                        "Lead Time", "Origin", "Country of Origin",
                        "Construction", "Frame", "Base", "COM", "COL",
                        "Seat Height", "Seat Depth", "Arm Height",
                        "Description", "Weight", "Size",
                    ):
                        label = line
                        if i + 1 < len(lines):
                            value = lines[i + 1]
                            # Map PDF label -> field name
                            field_map = {
                                "Dimensions(in)": "Dimensions",
                                "Dimensions (in)": "Dimensions",
                                "Dimensions(cm)": "Dimensions_cm",
                                "Dimensions (cm)": "Dimensions_cm",
                                "Finishes": "Finish",
                                "Materials": "Material",
                                "Features": "Features",
                                "Feature": "Features",
                                "Country of Origin": "Origin",
                            }
                            key = field_map.get(label, label)
                            if key and value and key not in data:
                                data[key] = value
                            i += 2
                            continue
                    i += 1
        # Parse Dimensions(in) into Width/Depth/Height if present
        if "Dimensions" in data:
            parsed = parse_dimensions(data["Dimensions"])
            for k in ("Width", "Depth", "Height", "Length", "Diameter"):
                if parsed.get(k) and k not in data:
                    data[k] = parsed[k]
        return data
    except Exception:
        return {}


def _distribute(variant_cols: list[dict], label: str, vals: list[str]) -> None:
    """Write spec values into variant dicts.

    If only 1 value for N variants (shared field like Collection), apply to all.
    Otherwise zip values to columns by position.
    """
    if not label:
        return
    n = len(variant_cols)
    if len(vals) == 1 and n > 1:
        # Shared value — apply to every variant
        for col in variant_cols:
            if not col.get(label):
                col[label] = vals[0]
    else:
        for i, v in enumerate(vals):
            if v and i < n and not variant_cols[i].get(label):
                variant_cols[i][label] = v


async def _parse_spec_tables(page) -> dict:
    """
    Parse ALL spec tables on the page.
    Returns a dict of {normalised_label: value_string}.
    Handles both simple 2-column tables and multi-column SKU-variant tables.
    Multi-column tables produce a nested dict keyed by column index.
    """
    raw = await page.evaluate("""
    () => {
        const tables = [];
        for (const table of document.querySelectorAll('table')) {
            const rows = [];
            for (const tr of table.querySelectorAll('tr')) {
                const cells = Array.from(tr.querySelectorAll('td, th'))
                                   .map(c => c.innerText.trim());
                if (cells.length >= 2 && cells.some(c => c)) {
                    rows.push(cells);
                }
            }
            if (rows.length) tables.push(rows);
        }
        return tables;
    }
    """)

    result: dict = {}
    variant_cols: list[dict] = []  # for multi-SKU products

    for table_rows in raw:
        # Detect multi-column (variant) table: >2 cells in a row
        max_cols = max(len(r) for r in table_rows)

        if max_cols > 2:
            # First row whose first cell is "SKU" defines the variant columns
            sku_row = next(
                (r for r in table_rows if r[0].lower() == "sku"),
                None
            )
            if sku_row:
                n_variants = len(sku_row) - 1
                if not variant_cols:
                    variant_cols = [{} for _ in range(n_variants)]
                for row in table_rows:
                    label = normalise_label(row[0])
                    vals = row[1:n_variants + 1]
                    _distribute(variant_cols, label, vals)
            elif variant_cols:
                # Data table for an already-established variant set (no SKU header here)
                # e.g. "Dimensions & Weights" table on bar stool pages
                n = len(variant_cols)
                for row in table_rows:
                    if len(row) < 2:
                        continue  # section-header row, skip
                    label = normalise_label(row[0])
                    vals = row[1:n + 1]
                    _distribute(variant_cols, label, vals)
            else:
                # Multi-column with no variant context → join values
                for row in table_rows:
                    label = normalise_label(row[0])
                    result[label] = " | ".join(v for v in row[1:] if v)
        else:
            for row in table_rows:
                label = normalise_label(row[0])
                val = row[1] if len(row) > 1 else ""
                if label and val:
                    result[label] = val

    return result, variant_cols


async def scrape_product(page, url: str) -> list[dict]:
    base: dict = {"Source URL": url, "Manufacturer": VENDOR_NAME}

    try:
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    except Exception as e:
        print(f"  [WARN] Failed to load {url}: {e}")
        return [base]
    await page.wait_for_timeout(2000)

    # — Product Name —
    h1 = await page.query_selector("h1")
    if h1:
        base["Product Name"] = clean_text(await h1.inner_text())

    # — SKU (right-side plain text: "SKU: XXXX") —
    sku_text: str = await page.evaluate("""
    () => {
        for (const el of document.querySelectorAll('*')) {
            const t = el.childNodes;
            for (const node of t) {
                if (node.nodeType === 3) {  // TEXT_NODE
                    const txt = node.textContent.trim();
                    if (/^SKU:\\s*\\S+/.test(txt)) return txt;
                }
            }
        }
        // Fallback: any element whose text starts with 'SKU:'
        for (const el of document.querySelectorAll('p, span, div, li')) {
            const txt = el.innerText.trim();
            if (txt.startsWith('SKU:') && txt.length < 60) return txt;
        }
        return '';
    }
    """)
    if sku_text:
        base["SKU"] = re.sub(r"^SKU:\s*", "", sku_text).strip()

    # — Description (paragraph near product info block) —
    desc: str = await page.evaluate("""
    () => {
        // Look for a paragraph that's a sibling of or near the SKU text
        for (const p of document.querySelectorAll('p')) {
            const txt = p.innerText.trim();
            if (txt.length > 30 && !txt.startsWith('SKU') && !/^shop/i.test(txt)) {
                return txt;
            }
        }
        return '';
    }
    """)
    if desc:
        base["Description"] = clean_text(desc)

    # — Tearsheet (Salsify PDF) —
    tearsheet: str = await page.evaluate("""
    () => {
        for (const a of document.querySelectorAll('a[href]')) {
            const href = a.getAttribute('href') || '';
            if (href.includes('salsify') && href.endsWith('.pdf')) return href;
        }
        // Also look for "Catalog Page" link text
        for (const a of document.querySelectorAll('a')) {
            const txt = a.innerText.trim().toLowerCase();
            if ((txt === 'click here' || txt.includes('catalog')) && a.getAttribute('href')?.includes('salsify')) {
                return a.getAttribute('href');
            }
        }
        return '';
    }
    """)
    if tearsheet:
        base["Tearsheet Link"] = tearsheet

    # — Tearsheet PDF data (Finish, Materials, Dimensions, Features, etc.) —
    if tearsheet:
        pdf_data = await asyncio.get_event_loop().run_in_executor(
            None, _fetch_tearsheet_data, tearsheet
        )
        for k, v in pdf_data.items():
            if v and k not in base:
                base[k] = v

    # — Images (rffblob, prefer _1170 resolution) —
    images: list[str] = await page.evaluate("""
    () => {
        const seen = new Set();
        const urls = [];
        for (const img of document.querySelectorAll('img[src*="rffblob"]')) {
            let src = img.getAttribute('src') || '';
            if (!src) continue;
            // Upgrade to highest resolution
            src = src.replace(/_\\d+(\\.jpeg)$/, '_1170$1');
            if (!seen.has(src)) { seen.add(src); urls.push(src); }
        }
        return urls;
    }
    """)
    if images:
        base["Image URL"] = images[0]

    # — Expand specs accordion then parse —
    await _expand_specs(page)
    spec_dict, variant_cols = await _parse_spec_tables(page)


    # If multi-SKU product: build one row per variant
    if variant_cols:
        product_name = base.get("Product Name", "")
        rows: list[dict] = []
        for col in variant_cols:
            row = dict(base)
            row.update(col)
            # Dimensions parsing
            dims_raw = row.pop("Dimensions", "") or col.get("Dimensions", "")
            if dims_raw:
                parsed = parse_dimensions(dims_raw)
                row["Dimensions"] = parsed.get("Dimensions", dims_raw)
                for k in ("Width", "Height", "Depth", "Length", "Diameter"):
                    if parsed.get(k) and not row.get(k):
                        row[k] = parsed[k]
            # Weight — strip " LB" suffix
            w = row.get("Weight", "")
            if w:
                row["Weight"] = re.sub(r"\s*lb.*$", "", str(w), flags=re.IGNORECASE).strip()
            # Product Family Id
            if not row.get("Product Family Id") and product_name:
                row["Product Family Id"] = extract_family_id(product_name)
            rows.append(row)
        return rows

    # Single-SKU product
    row = dict(base)
    # Override SKU if found in spec table (spec table SKU is more reliable for fabrics)
    if "SKU" in spec_dict and not row.get("SKU"):
        row["SKU"] = spec_dict.pop("SKU")
    else:
        spec_dict.pop("SKU", None)

    # Copy all spec fields into row
    for k, v in spec_dict.items():
        if v:
            row[k] = v

    # Dimensions
    dims_raw = row.pop("Dimensions", "")
    if dims_raw:
        parsed = parse_dimensions(dims_raw)
        row["Dimensions"] = parsed.get("Dimensions", dims_raw)
        for k in ("Width", "Height", "Depth", "Length", "Diameter"):
            if parsed.get(k) and not row.get(k):
                row[k] = parsed[k]

    # Weight — strip " LB"
    w = row.get("Weight", "")
    if w:
        row["Weight"] = re.sub(r"\s*lb.*$", "", str(w), flags=re.IGNORECASE).strip()

    # Product Family Id
    if not row.get("Product Family Id") and row.get("Product Name"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    return [row]


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

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )
            print(f"\n[{cat['name']}] Collecting product links …")

            # Clear session/cookies before each category to avoid "recently viewed" contamination
            await page.context.clear_cookies()

            seen_urls: set[str] = set()
            url_queue: deque[str] = deque()
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        url_queue.append(u)

            print(f"  -> {len(url_queue)} candidate URLs")
            global_idx = 1

            while url_queue:
                if TEST_MODE and global_idx > TEST_MAX_PRODUCTS:
                    break

                url = url_queue.popleft()
                try:
                    variant_rows = await scrape_product(page, url)
                except Exception as e:
                    print(f"  [ERR] {url}: {e}")
                    variant_rows = []

                # Detect sub-category pages: no SKU and no dimension/spec data
                is_real_product = any(
                    r.get("SKU") or r.get("Weight") or r.get("Collection")
                    or r.get("Height") or r.get("Dimensions")
                    for r in variant_rows
                )
                if not is_real_product:
                    print(f"  [SUB] {url.split('/')[-1]} -> crawling for products...")
                    for sub_url in await get_product_links(page, url):
                        if sub_url not in seen_urls:
                            seen_urls.add(sub_url)
                            url_queue.appendleft(sub_url)
                    continue

                for row in variant_rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1

                print(f"    OK {global_idx - 1:>3}  {url.split('/')[-1][:55]}"
                      f"  ({len(variant_rows)} row{'s' if len(variant_rows) != 1 else ''})")
                await async_polite_delay()

    writer.save()
    print(f"\nSaved -> {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
