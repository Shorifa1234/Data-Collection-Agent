import asyncio, json, os, re, sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, clean_price,
    generate_sku, extract_family_id, parse_dimensions,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Charles Stewart")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
TEST_MODE   = os.environ.get("TEST_MODE", "false").lower() == "true"
SCRAPE_CATEGORIES = [c.strip() for c in os.environ.get("SCRAPE_CATEGORIES", "").split(",") if c.strip()]
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

BASE_URL = "https://charlesstewartcompany.com"

# Dimension label → field name mapping
DIM_MAP = {
    "overall width":  "Width",
    "overall height": "Height",
    "overall depth":  "Depth",
    "overall length": "Length",
    "diameter":       "Diameter",
    "seat width":     "Seat Width",
    "seat height":    "Seat Height",
    "seat depth":     "Seat Depth",
    "seat length":    "Seat Length",
    "arm height":     "Arm Height",
    "leg height":     "Leg Height",
    "back height":    "Back Height",
}

# COM yardage labels → field names
COM_MAP = {
    "plain fabric":             "COM Plain Fabric (yds)",
    "7\" - 14\" vertical repeat":  "COM 7-14in VR (yds)",
    "15\" - 27\" vr":           "COM 15-27in VR (yds)",
    "15\" - 27\" vertical repeat": "COM 15-27in VR (yds)",
    "over 27\" vr":             "COM Over 27in VR (yds)",
    "over 27\" vertical repeat":  "COM Over 27in VR (yds)",
    "col":                       "COL (sq ft)",
}

# Variant/dimension table column header → field name
TABLE_COL_MAP = {
    "overall w":       "Width",
    "overall width":   "Width",
    "overall d":       "Depth",
    "overall depth":   "Depth",
    "overall h":       "Height",
    "overall height":  "Height",
    "overall l":       "Length",
    "overall length":  "Length",
    "diameter":        "Diameter",
    "seat h":          "Seat Height",
    "seat height":     "Seat Height",
    "seat w":          "Seat Width",
    "seat width":      "Seat Width",
    "seat d":          "Seat Depth",
    "seat depth":      "Seat Depth",
    "arm h":           "Arm Height",
    "arm height":      "Arm Height",
    "com plain ydg.":  "COM Plain Fabric (yds)",
    "com plain yardage": "COM Plain Fabric (yds)",
    "com ydg.":        "COM Plain Fabric (yds)",
    "com yardage":     "COM Plain Fabric (yds)",
}


def _strip_size_suffix(url: str) -> str:
    """Remove WP image size suffix: -410x410 → (original size)."""
    return re.sub(r"-\d+x\d+(?=\.\w{2,4}$)", "", url)


def _parse_variant_cell(cell_text: str) -> tuple[str, str]:
    """
    Split 'King S852-82' → ('King', 'S852-82').
    Returns (size_label, sku).  sku may be '' if not found.
    """
    # SKU pattern: letters+digits with optional hyphens, e.g. S852-82, C820-00, CS51
    m = re.search(r'\b([A-Z][A-Z0-9]{1,5}-\d{2,4}(?:-\d+)?|[A-Z]{2,4}\d{2,4}(?:-\d+)?)\b', cell_text)
    if m:
        sku = m.group(1)
        size = cell_text[:m.start()].strip().rstrip(",").strip()
        return size, sku
    return cell_text.strip(), ""


def _parse_tables_to_variants(raw_tables: list[list[list[str]]], base: dict) -> list[dict]:
    """
    Convert a list of raw HTML tables (list of rows, each row a list of cell texts)
    into variant dicts.  Each non-header row becomes one variant.

    Table structure:
      row 0: [configuration_name, col1_header, col2_header, ...]
      row 1+: [size_SKU_label, val1, val2, ...]
    """
    variants: list[dict] = []

    for table in raw_tables:
        if len(table) < 2:
            continue  # need at least header + 1 data row

        header_row = table[0]
        if not header_row:
            continue

        config_name = clean_text(header_row[0])  # e.g. "Headboard Only" / "Complete Bed"
        col_headers  = [c.strip().lower() for c in header_row[1:]]

        for row in table[1:]:
            if not row or not row[0].strip():
                continue

            size_label, row_sku = _parse_variant_cell(clean_text(row[0]))
            variant = dict(base)  # shallow copy of shared fields

            if row_sku:
                variant["SKU"] = row_sku
            if size_label:
                variant["Size"] = size_label
            if config_name:
                variant["Configuration"] = config_name

            # Build dimension string and individual fields
            dim_parts: list[str] = []
            for col_name, cell_val in zip(col_headers, row[1:]):
                val = _clean_dim_value(cell_val)
                if not val:
                    continue
                field = TABLE_COL_MAP.get(col_name)
                if field:
                    if "COM" in field:
                        variant[field] = _clean_com_value(val)
                    else:
                        variant[field] = val
                        dim_parts.append(f"{field}: {val}")
                else:
                    variant[col_name.title()] = val

            if dim_parts:
                variant["Dimensions"] = ", ".join(dim_parts)

            variants.append(variant)

    return variants


def _clean_dim_value(val: str) -> str:
    """Strip inch marks and trailing units from dimension values: '51"' → '51'."""
    val = val.strip().rstrip('"').strip()
    val = re.sub(r'\s*(in|inch|inches)$', '', val, flags=re.I).strip()
    return val


def _clean_com_value(val: str) -> str:
    """Extract numeric part from COM yardage: '5.5 yds' → '5.5'."""
    m = re.match(r'[\d.]+', val.strip())
    return m.group(0) if m else val.strip()


async def _get_product_image(page) -> str:
    """
    Find the main product image URL.
    Priority:
      1. Non-SVG img inside .wr360-gallery nav (thumbnail swiper)
      2. Any wp-content/uploads img that is not logo / svg / 360thumb / plugin
    Returns full-size URL (size suffix stripped).
    """
    # Try within wr360-gallery nav / thumbnail swiper
    imgs = await page.evaluate("""() => {
        const all = [...document.querySelectorAll('img')];
        return all.map(i => ({ src: i.src, alt: i.alt, cls: i.className }));
    }""")
    skip_patterns = ("svg", "logo", "360thumb", "/plugins/", "Charles-Stewart", "White")
    for img in imgs:
        src = img.get("src", "")
        if (src
                and "wp-content/uploads" in src
                and not any(p in src for p in skip_patterns)
                and not src.startswith("data:")):
            return _strip_size_suffix(src)
    return ""


async def _get_description_and_specs(page) -> dict:
    """
    Extract description, specifications, and downloads from the product page.
    The site uses custom collapse sections that may be hidden.
    We read visible text blocks and PDF links.
    """
    result = {}

    # --- Description: short description block --------------------------------
    # Skip form-like content injected by GravityForms / Popup Maker shortcodes
    _FORM_NOISE = ("Required", "Mailing Address", "ZIP", "Company Name", "Street Address")

    for sel in [
        ".woocommerce-product-details__short-description",
        ".short-description",
        ".product-description",
        ".entry-content .description",
        "div[class*='description']:not([class*='short-description'])",
    ]:
        el = await page.query_selector(sel)
        if el:
            txt = clean_text(await el.inner_text())
            if txt and len(txt) > 20 and not any(noise in txt for noise in _FORM_NOISE):
                result["Description"] = txt
                break

    # --- Collapse / accordion sections: try clicking and reading -------------
    # The site uses custom collapse divs. Find headings that look like
    # accordion triggers and click them to reveal content.
    accordion_sections = {}
    try:
        # Look for any elements whose text is exactly one of these headings
        target_headings = {"Descriptions", "Description", "Specifications", "Specs", "Downloads"}
        triggers = await page.evaluate("""(headings) => {
            const all = [...document.querySelectorAll(
                'p, span, div, h2, h3, h4, a, button, li'
            )];
            return all
                .filter(el => {
                    const t = el.textContent.trim();
                    return headings.includes(t) && el.children.length === 0;
                })
                .map(el => ({
                    text: el.textContent.trim(),
                    tag: el.tagName,
                    cls: el.className,
                    parent_cls: el.parentElement ? el.parentElement.className : '',
                    parent_tag: el.parentElement ? el.parentElement.tagName : '',
                }));
        }""", list(target_headings))

        if triggers:
            for t in triggers:
                # Click the parent element (the trigger wrapper)
                try:
                    parent_sel = f"{t['parent_tag'].lower()}.{t['parent_cls'].split()[0]}" if t['parent_cls'] else t['parent_tag'].lower()
                    matching = await page.query_selector_all(parent_sel)
                    for el in matching:
                        txt = clean_text(await el.inner_text())
                        if txt.strip() in target_headings:
                            await el.click()
                            await page.wait_for_timeout(500)
                            break
                except Exception:
                    pass

            # After clicks, try to read revealed content
            await page.wait_for_timeout(800)

        # Now scrape visible text from the product content area
        # Exclude popups and GravityForms which pollute innerText
        content_text = await page.evaluate("""() => {
            const product_div = document.querySelector('[id^="product-"]');
            if (!product_div) return document.body.innerText;
            // Temporarily hide popups / forms so innerText skips them
            const noise = [...product_div.querySelectorAll('.pum-overlay, .pum-container, .gform_wrapper, [id^="pum-"], [id^="popmake-"]')];
            noise.forEach(el => el.style.display = 'none');
            const text = product_div.innerText;
            noise.forEach(el => el.style.display = '');
            return text;
        }""")

        # Parse sections from the visible text
        lines = [l.strip() for l in content_text.splitlines() if l.strip()]
        current_section = None
        section_content = {}
        section_keywords = {
            "Descriptions": "Description",
            "Description": "Description",
            "Specifications": "Specifications",
            "Specs": "Specifications",
            "Downloads": "Downloads",
        }
        for line in lines:
            if line in section_keywords:
                current_section = section_keywords[line]
                section_content[current_section] = []
            elif current_section:
                # Stop at next heading or unrelated content
                if line in section_keywords or line in ("Print This Page", "View Materials", "Login", "Register"):
                    current_section = None
                elif len(line) > 3:
                    section_content[current_section].append(line)

        for section, parts in section_content.items():
            txt = " ".join(parts).strip()
            if txt and section not in result:
                result[section] = txt

    except Exception:
        pass

    # --- Downloads: collect all PDF links on page ----------------------------
    pdf_links = await page.evaluate("""() => {
        return [...document.querySelectorAll('a')]
            .filter(a => a.href && (a.href.includes('.pdf') || a.href.includes('.PDF')))
            .map(a => ({ text: a.textContent.trim(), href: a.href }));
    }""")
    if pdf_links:
        result["Tearsheet Link"] = pdf_links[0]["href"]
        if len(pdf_links) > 1:
            result["Downloads"] = " | ".join(f"{p['text'] or 'Download'}: {p['href']}" for p in pdf_links)

    return result


async def scrape_product(page, url: str, cat_name: str) -> list[dict]:
    """
    Return a list of dicts — one per variant row.
    Two modes:
      A. Variant-table mode  — page has <table> elements with size rows (e.g. Headboard)
      B. Single-product mode — page uses div.product-dimensions (e.g. Sofa, Table)
    """
    base = {"Source URL": url}

    try:
        await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)

        # ── Shared fields (same for every variant) ────────────────────────
        name_el = await page.query_selector("h1.product_title, h1.entry-title")
        if name_el:
            base["Product Name"] = clean_text(await name_el.inner_text())

        base["Product Family Id"] = extract_family_id(base.get("Product Name", ""))
        base["Manufacturer"] = VENDOR_NAME

        sku_el = await page.query_selector("span.sku")
        if sku_el:
            base["SKU"] = clean_text(await sku_el.inner_text())

        base["Image URL"] = await _get_product_image(page)

        # ── Description, PDF links ────────────────────────────────────────
        extra = await _get_description_and_specs(page)
        base.update({k: v for k, v in extra.items() if v})

        # ── Detect mode: variant table vs single-product ──────────────────
        raw_tables = await page.evaluate("""() => {
            const product_div = document.querySelector('[id^="product-"]') || document.body;
            return [...product_div.querySelectorAll('table')].map(t => {
                return [...t.querySelectorAll('tr')].map(r =>
                    [...r.querySelectorAll('td, th')].map(c => c.innerText.trim())
                );
            });
        }""")

        # Filter to tables that look like dimension/variant tables
        # (first row has column headers like "Overall W", "Overall H", "COM")
        dim_tables = []
        for tbl in raw_tables:
            if len(tbl) < 2:
                continue
            first_row_text = " ".join(tbl[0]).lower()
            if any(kw in first_row_text for kw in ("overall w", "overall h", "overall d", "com")):
                dim_tables.append(tbl)

        if dim_tables:
            # ── MODE A: Variant-table products ────────────────────────────
            variants = _parse_tables_to_variants(dim_tables, base)
            if variants:
                print(f"    [tables] {len(variants)} size variants from {len(dim_tables)} table(s)")
                return variants

        # ── MODE B: Single-product (div.product-dimensions) ───────────────
        dim_data = await page.evaluate("""() => {
            const rows = [...document.querySelectorAll('.product-dimensions .table-row')];
            return rows.map(row => {
                const hdr = row.querySelector('.table-header');
                const val = row.querySelector('.table-value');
                const cell = hdr || val;
                return cell ? cell.textContent.trim() : '';
            }).filter(t => t.includes(':'));
        }""")

        dim_parts = []
        for line in dim_data:
            if ":" not in line:
                continue
            label, _, value = line.partition(":")
            label = label.strip().lower()
            value = _clean_dim_value(value)
            if not value:
                continue
            field = DIM_MAP.get(label)
            if field:
                base[field] = value
            else:
                base[label.title()] = value
            dim_parts.append(f"{label.title()}: {value}")

        if dim_parts:
            base["Dimensions"] = ", ".join(dim_parts)

        # COM yardage lines (single-product mode only)
        com_data = await page.evaluate("""() => {
            const product_div = document.querySelector('[id^="product-"]');
            if (!product_div) return [];
            const noise = [...product_div.querySelectorAll('.pum-overlay,.pum-container,.gform_wrapper,[id^="pum-"],[id^="popmake-"]')];
            noise.forEach(el => el.style.display = 'none');
            const lines = product_div.innerText.split('\\n').map(l => l.trim()).filter(l => l);
            noise.forEach(el => el.style.display = '');
            return lines;
        }""")
        for line in com_data:
            if ":" not in line:
                continue
            label, _, value = line.partition(":")
            field = COM_MAP.get(label.strip().lower())
            if field and value.strip():
                base[field] = _clean_com_value(value.strip())

        # Related Products
        related = await page.evaluate("""() => {
            const links = [...document.querySelectorAll('a')].filter(a =>
                a.textContent.trim().length > 0 &&
                a.closest('.related, .related-products, .upsells, .crosssells')
            );
            return [...new Set(links.map(a => a.textContent.trim()))].slice(0, 5).join(', ');
        }""")
        if related:
            base["Related Products"] = related

    except Exception as e:
        print(f"  [WARN] Failed to scrape {url}: {e}")
        base.setdefault("Product Name", "")

    return [base]


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from a category listing (handles pagination)."""
    links: list[str] = []
    seen: set[str] = set()
    current_url = listing_url
    page_num = 1

    while current_url:
        print(f"    Listing page {page_num}: {current_url}")
        try:
            await page.goto(current_url, timeout=45_000, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
        except Exception as e:
            print(f"    [WARN] Could not load listing page: {e}")
            break

        # Extract product hrefs
        hrefs = await page.evaluate("""() => {
            const sections = [...document.querySelectorAll('section.product, li.product')];
            const out = [];
            for (const sec of sections) {
                const a = sec.querySelector('a[href*="/product/"]');
                if (a) out.push(a.href.split('?')[0]);
            }
            return out;
        }""")

        added = 0
        for href in hrefs:
            if href not in seen:
                seen.add(href)
                links.append(href)
                added += 1
        print(f"      Found {added} new products (total so far: {len(links)})")

        if TEST_MODE and len(links) >= 5:
            break

        # Next page
        next_href = await page.evaluate("""() => {
            const nxt = document.querySelector('nav.woocommerce-pagination a.next, .page-numbers a.next');
            return nxt ? nxt.href : null;
        }""")
        if next_href and next_href != current_url:
            current_url = next_href
            page_num += 1
        else:
            break

    return links


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in info["categories"]:
            if not cat["links"]:
                continue
            if SCRAPE_CATEGORIES and cat["name"] not in SCRAPE_CATEGORIES:
                continue

            print(f"\n{'='*60}")
            print(f"Category: {cat['name']}")

            writer.add_sheet(cat["name"], cat["links"][0], studio_columns=cat["studio_columns"])

            # Collect all product URLs (deduplicated across multiple listing links)
            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            print(f"  Total products to scrape: {len(all_product_urls)}")

            if TEST_MODE:
                all_product_urls = all_product_urls[:5]

            global_idx = 1
            for url in all_product_urls:
                print(f"  [{global_idx}/{len(all_product_urls)}] {url}")
                rows = await scrape_product(page, url, cat["name"])
                for row in rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
                await async_polite_delay()

            if TEST_MODE:
                print(f"  [TEST MODE] 5 products done, next category...")

    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
