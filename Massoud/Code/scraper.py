import asyncio, json, os, re, sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, extract_family_id,
    generate_sku,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Massoud")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

BASE_URL = "https://www.massoudfurniture.com"

# Slugs that are listing/nav pages — not individual products
_LISTING_SLUGS = {
    "sofas", "chairs", "sectionals", "dining", "beds-headboards",
    "benches-ottomans", "fabrics", "leathers", "trims", "throw-pillows",
    "furniture", "textiles", "custom-choices",
    "our-company", "contact", "dealer-portal", "store-locator",
    "lookbook", "search", "cart", "checkout", "account", "my-account",
    # Info/nav pages in the textiles section
    "cushions", "nail-trim", "wood-finishes",
}


# ─────────────────────────────────────────────
# Listing page — collect product URLs
# ─────────────────────────────────────────────

async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs from a paginated listing page."""
    links: list[str] = []
    seen: set[str]   = set()
    current_url      = listing_url
    is_textile       = "/textiles/" in listing_url

    while current_url:
        print(f"    Listing: {current_url}")
        await page.goto(current_url, timeout=60_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)

        hrefs: list[str] = await page.evaluate("""() =>
            Array.from(document.querySelectorAll('a[href]')).map(a => a.href)
        """)

        for h in hrefs:
            clean = h.split("?")[0].split("#")[0].rstrip("/")
            if "massoudfurniture.com" not in clean:
                continue
            path  = clean.replace("https://www.massoudfurniture.com", "")
            parts = [p for p in path.split("/") if p]
            # Product pages have exactly 2 segments: /section/slug
            if len(parts) != 2:
                continue
            section, slug = parts
            # Must be furniture or textiles section
            if section not in ("furniture", "textiles"):
                continue
            # Skip known category/nav slugs (exact match)
            if slug in _LISTING_SLUGS:
                continue
            # Textile listing → only accept textiles URLs; furniture listing → only furniture
            if is_textile and section != "textiles":
                continue
            if not is_textile and section != "furniture":
                continue

            full = f"{BASE_URL}/{section}/{slug}/"
            if full not in seen:
                seen.add(full)
                links.append(full)

        # Pagination
        next_url: str | None = await page.evaluate("""() => {
            const el = document.querySelector(
                'a.next, a[rel="next"], .pagination .next a, .nav-links .next a'
            );
            return el ? el.href : null;
        }""")
        current_url = next_url if next_url and next_url != current_url else None

    return links


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _extract_labeled_dim(text: str, abbr: str) -> str:
    """Extract numeric value for a labeled dimension abbr (W/D/H/L) from a string."""
    m = re.search(rf'\b{abbr}\s*([\d.]+)"?', text, re.IGNORECASE)
    return m.group(1) if m else ""


async def _open_accordions(page) -> None:
    """Force-open all collapsed accordion sections on the page."""
    await page.evaluate("""() => {
        // HTML details/summary
        document.querySelectorAll('details:not([open])').forEach(d => {
            d.setAttribute('open', '');
        });
        // aria-expanded=false buttons
        document.querySelectorAll('[aria-expanded="false"]').forEach(b => {
            try { b.click(); } catch(e) {}
        });
        // Any clickable element that looks like a collapsed accordion header
        document.querySelectorAll('button, [role="button"], h3, h4').forEach(b => {
            const t = b.textContent.trim().toUpperCase();
            if (t === 'DETAILS' || t === 'DIMENSIONS' || t === 'STANDARD FEATURES'
                || t === 'STANDARD FEATURES +' || t === 'ALSO AVAILABLE'
                || t.endsWith('+')) {
                try { b.click(); } catch(e) {}
            }
        });
        // Click any element containing only a '+' sign (accordion toggle)
        document.querySelectorAll('span, div, button').forEach(el => {
            if (el.children.length === 0 && el.textContent.trim() === '+') {
                try { el.parentElement.click(); el.click(); } catch(e) {}
            }
        });
    }""")
    await page.wait_for_timeout(1000)


async def _extract_spec_table(page) -> dict[str, str]:
    """
    Extract key-value spec pairs using multiple DOM strategies.
    Handles single-value siblings AND multi-value dimension rows.
    """
    return await page.evaluate("""() => {
        const result = {};

        const KNOWN = new Set([
            'OVERALL','INSIDE','ARM','SEAT','COM','COL',
            'BACK CUSHION','SEAT CUSHION','THROW PILLOWS',
            'CONTENT','GRADE','REPEATS','SUSTAINABILITY','ORIGIN',
        ]);

        // Helper: only store value if it is longer than what we already have
        function addResult(key, val) {
            val = (val || '').trim();
            if (key && val && val.length > (result[key] || '').length) {
                result[key] = val;
            }
        }

        // Strategy 1: dt/dd definition lists (may only give first dd — let later strategies extend)
        document.querySelectorAll('dt').forEach(dt => {
            const key = dt.textContent.trim().toUpperCase();
            const dd  = dt.nextElementSibling;
            if (dd) addResult(key, dd.textContent);
        });

        // Strategy 2: table rows — join ALL cells after the first (multi-cell dimension rows)
        document.querySelectorAll('tr').forEach(row => {
            const cells = Array.from(row.querySelectorAll('td, th'));
            if (cells.length >= 2) {
                const key = cells[0].textContent.trim().toUpperCase();
                const val = cells.slice(1).map(c => c.textContent.trim()).join(' ').trim();
                addResult(key, val);
            }
        });

        // Strategy 3: leaf element matching known key → collect ALL next siblings' text
        // Handles rows where W, D, H are in consecutive sibling span/div elements
        document.querySelectorAll('*').forEach(el => {
            if (el.children.length > 0) return;  // leaf only
            const key = el.textContent.trim().toUpperCase();
            if (!KNOWN.has(key)) return;

            // Collect all next siblings until we hit another known key or run out
            const parts = [];
            let sib = el.nextElementSibling;
            while (sib) {
                const sibKey = sib.textContent.trim().toUpperCase();
                if (KNOWN.has(sibKey)) break;
                const sibText = sib.textContent.trim();
                if (sibText) parts.push(sibText);
                sib = sib.nextElementSibling;
            }

            // Also try parent element text minus the key itself (captures inline W D H)
            if (el.parentElement) {
                const parentText = el.parentElement.textContent.trim();
                const stripped   = parentText.replace(el.textContent.trim(), '').trim();
                if (stripped && stripped.length > parts.join(' ').length) {
                    addResult(key, stripped);
                    return;
                }
            }

            if (parts.length > 0) addResult(key, parts.join(' '));
        });

        return result;
    }""")


async def _get_image_url(page) -> str:
    """Return the main product hero image URL."""
    url: str = await page.evaluate("""() => {
        const selectors = [
            '.product-gallery__main img',
            '.product-gallery img',
            '.product-image img',
            '.product__image img',
            '.woocommerce-product-gallery__image img',
            'figure.product-image img',
            '.attachment-shop_single',
            '[class*="product-hero"] img',
            '[class*="product-detail"] img',
            'article img',
        ];
        for (const sel of selectors) {
            const img = document.querySelector(sel);
            if (img) {
                const src = img.getAttribute('data-zoom-image') ||
                    img.getAttribute('data-large_image') ||
                    img.getAttribute('data-src') ||
                    img.getAttribute('data-lazy-src') ||
                    img.src || '';
                if (src && !src.startsWith('data:')) return src;
            }
        }
        // Broad: all images, pick largest by src quality heuristic
        for (const img of document.querySelectorAll('main img, #content img, .site-main img, body img')) {
            const src = img.getAttribute('data-src') || img.getAttribute('data-zoom-image') || img.src || '';
            if (src && !src.startsWith('data:') && !src.includes('logo') && !src.includes('icon')
                && !src.includes('placeholder') && src.includes('massoudfurniture')) {
                return src;
            }
        }
        return '';
    }""")
    if not url or url.startswith("data:"):
        return ""
    return url if url.startswith("http") else BASE_URL + url


async def _get_description(page) -> str:
    raw: str = await page.evaluate("""() => {
        const selectors = [
            '.product-description',
            '.product__description',
            '[class*="product-detail"] .description',
            '[class*="description"] p',
            '.entry-summary p',
        ];
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el) {
                const txt = el.textContent.trim();
                if (txt.length > 20) return txt;
            }
        }
        return '';
    }""")
    return clean_text(raw)


async def _get_tearsheet(page) -> str:
    """Return tear sheet URL — must be a direct PDF link, not the generic catalog page."""
    url: str = await page.evaluate("""() => {
        for (const a of document.querySelectorAll('a')) {
            const txt  = a.textContent.trim().toUpperCase();
            const href = (a.href || '').toLowerCase();
            // Only accept a specific tear-sheet PDF, not the generic catalogs/#tear-sheets anchor
            if ((txt.includes('TEAR') || href.includes('tear-sheet'))
                && href.endsWith('.pdf')) {
                return a.href;
            }
        }
        return '';
    }""")
    return url


# ─────────────────────────────────────────────
# Product page scrapers
# ─────────────────────────────────────────────

_INCH = re.compile(r'[″”ʺ"]')  # any inch/quote mark variant

def _strip_inch(s: str) -> str:
    return _INCH.sub("", s).strip()


_DIM_LABEL_MAP = {"W": "Width", "D": "Depth", "H": "Height", "L": "Length"}

def _parse_overall(raw: str) -> dict[str, str]:
    """
    Parse 'W82"D108"H60"' or 'W106" D41" H34"' into Width/Depth/Height.
    Uses finditer so it works even when spans are concatenated without spaces.
    """
    result: dict[str, str] = {}
    text = _strip_inch(raw)
    for m in re.finditer(r'([WDHLwdhl])\s*([\d.]+)', text):
        key = _DIM_LABEL_MAP.get(m.group(1).upper())
        if key and key not in result:
            result[key] = m.group(2)
    return result


def _parse_h_value(raw: str) -> str:
    """Extract the H number from strings like 'H25"' or 'H18"'."""
    text = _strip_inch(raw)
    m = re.search(r'H\s*([\d.]+)', text, re.IGNORECASE)
    return m.group(1) if m else ""



async def _scrape_furniture(page, url: str) -> dict:
    """Parse data from a furniture/seating product page using DOM spec table."""
    data: dict = {}

    specs = await _extract_spec_table(page)

    # OVERALL → Width, Depth, Height, Dimensions
    if "OVERALL" in specs:
        dims = _parse_overall(specs["OVERALL"])
        data.update(dims)
        parts = " ".join(filter(None, [
            f"W{dims.get('Width','')}" if dims.get('Width') else "",
            f"D{dims.get('Depth','')}" if dims.get('Depth') else "",
            f"H{dims.get('Height','')}" if dims.get('Height') else "",
        ]))
        if parts:
            data["Dimensions"] = parts

    # INSIDE → Inside Width, Inside Depth
    if "INSIDE" in specs:
        itext = _strip_inch(specs["INSIDE"])
        mw = re.search(r'W\s*([\d.]+)', itext, re.IGNORECASE)
        md = re.search(r'D\s*([\d.]+)', itext, re.IGNORECASE)
        if mw: data["Inside Width"] = mw.group(1)
        if md: data["Inside Depth"] = md.group(1)

    # ARM → Arm Height
    if "ARM" in specs:
        v = _parse_h_value(specs["ARM"])
        if v: data["Arm Height"] = v

    # SEAT → Seat Height
    if "SEAT" in specs:
        v = _parse_h_value(specs["SEAT"])
        if v: data["Seat Height"] = v

    # COM — numeric value only
    if "COM" in specs:
        m = re.search(r'([\d.]+)', specs["COM"])
        if m: data["COM"] = m.group(1)

    # COL — numeric value only
    if "COL" in specs:
        m = re.search(r'([\d.]+)', specs["COL"])
        if m: data["COL"] = m.group(1)

    # BACK CUSHION
    if "BACK CUSHION" in specs:
        val = clean_text(specs["BACK CUSHION"])
        if val: data["Back"] = val

    # SEAT CUSHION
    if "SEAT CUSHION" in specs:
        val = clean_text(specs["SEAT CUSHION"])
        if val: data["Seat Cushion"] = val

    # THROW PILLOWS
    if "THROW PILLOWS" in specs:
        val = clean_text(specs["THROW PILLOWS"])
        if val: data["Throw Pillows"] = val

    # Tear Sheet (PDF only)
    ts = await _get_tearsheet(page)
    if ts:
        data["Tearsheet Link"] = ts

    return data


async def _scrape_textile(page, page_text: str) -> dict:
    """Parse data from a textile (fabric/leather/trim) product page."""
    data: dict = {}

    # Primary: structured DOM extraction
    specs = await _extract_spec_table(page)

    if "CONTENT" in specs:
        data["Content"] = clean_text(specs["CONTENT"])

    if "GRADE" in specs:
        m = re.search(r'([\d.]+)', specs["GRADE"])
        if m: data["Grade"] = m.group(1)

    if "REPEATS" in specs:
        raw = specs["REPEATS"]
        # "H 0.00" V 0.00""
        m = re.search(r'H\s*([\d.]+)[^\d]*V\s*([\d.]+)', _strip_inch(raw), re.IGNORECASE)
        if m:
            data["Horizontal Repeat"] = m.group(1)
            data["Vertical Repeat"]   = m.group(2)

    if "SUSTAINABILITY" in specs:
        raw_sus = specs["SUSTAINABILITY"].strip()
        # Handle no-space camelCase concat: "Natural MaterialsRecycled Content"
        sus = re.sub(r'([a-z])([A-Z])', r'\1 / \2', raw_sus)
        # Also handle embedded newlines
        sus = re.sub(r'\s*\n\s*', ' / ', sus)
        data["Sustainability"] = clean_text(sus)

    if "ORIGIN" in specs:
        data["Origin"] = clean_text(specs["ORIGIN"])

    # Fallback: regex on page text for any fields still missing
    if "Content" not in data:
        m = re.search(r'\bCONTENT\b[\s\n]+(.+?)(?=\n|GRADE|REPEAT|SUSTAIN|ORIGIN|$)', page_text, re.IGNORECASE)
        if m:
            val = clean_text(m.group(1))
            if val and len(val) < 150: data["Content"] = val

    if "Grade" not in data:
        m = re.search(r'\bGRADE\b[\s\n]+([\d]+)', page_text, re.IGNORECASE)
        if m: data["Grade"] = m.group(1)

    if "Origin" not in data:
        m = re.search(r'\bORIGIN\b[\s\n]+([A-Za-z ,]+?)(?=\n|$)', page_text, re.IGNORECASE)
        if m:
            val = clean_text(m.group(1))
            if val and len(val) < 60: data["Origin"] = val

    return data


# ─────────────────────────────────────────────
# Main product scraper
# ─────────────────────────────────────────────

async def scrape_product(page, url: str, cat_name: str) -> list[dict]:
    """Scrape one product page. Returns list[dict] (single row for Massoud)."""
    base: dict = {"Source URL": url}
    is_textile = "/textiles/" in url

    try:
        await page.goto(url, timeout=60_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)
    except Exception as exc:
        print(f"    [SKIP] Load failed {url}: {exc}")
        return []

    # ── Product Name + SKU ────────────────────────────────────────
    # Massoud puts the model code inside the h1 as a child element.
    # e.g. <h1>Laguna King Bed<span>31KB</span></h1>
    # We want the name text only (parent text nodes) and SKU from the child.
    name_sku: dict = await page.evaluate("""() => {
        const h1 = document.querySelector('h1');
        if (!h1) return {name: '', sku: ''};

        // Collect text-only from direct text nodes (not child elements)
        let nameParts = [];
        let skuEl = null;
        for (const node of h1.childNodes) {
            if (node.nodeType === 3) {  // TEXT_NODE
                const t = node.textContent.trim();
                if (t) nameParts.push(t);
            } else if (node.nodeType === 1) {  // ELEMENT_NODE
                const t = node.textContent.trim();
                // Short alphanumeric = SKU (model code)
                if (/^[A-Z0-9]{1,12}$/i.test(t)) {
                    skuEl = t;
                } else if (t) {
                    nameParts.push(t);
                }
            }
        }
        let name = nameParts.join(' ').trim();
        // Fallback: full h1 text when no child elements
        if (!name) name = h1.textContent.trim();

        // If sku not found in h1, check next sibling elements
        if (!skuEl) {
            let el = h1.nextElementSibling;
            for (let i = 0; i < 4 && el; i++) {
                const t = el.textContent.trim();
                if (/^[A-Z0-9]{1,12}$/i.test(t)) { skuEl = t; break; }
                // Only walk next if el is short/inline — stop at longer elements
                if (t.length > 20) break;
                el = el.nextElementSibling;
            }
        }

        return {name: name, sku: skuEl || ''};
    }""")

    product_name = clean_text(name_sku.get("name", ""))
    if not product_name:
        print(f"    [SKIP] No product name at {url}")
        return []
    base["Product Name"] = product_name

    sku_raw = name_sku.get("sku", "").strip().upper()
    if sku_raw and re.match(r'^[A-Z0-9]{1,12}$', sku_raw):
        base["SKU"] = sku_raw
    else:
        # Fallback: extract trailing code from URL slug
        slug  = url.rstrip("/").rsplit("/", 1)[-1]
        parts = slug.rsplit("-", 1)
        if len(parts) > 1:
            last = parts[-1]
            if re.match(r'^[0-9]+[A-Z]*$|^[A-Z]+[0-9]+[A-Z]*$', last, re.IGNORECASE):
                base["SKU"] = last.upper()

    # ── Image URL ────────────────────────────────────────────────
    img = await _get_image_url(page)
    if img:
        base["Image URL"] = img

    # ── Description ──────────────────────────────────────────────
    desc = await _get_description(page)
    if desc:
        base["Description"] = desc

    # ── Open all accordions ──────────────────────────────────────
    await _open_accordions(page)

    # ── Full page text for regex parsing ─────────────────────────
    page_text: str = await page.evaluate("() => document.body.innerText")

    if is_textile:
        textile_data = await _scrape_textile(page, page_text)
        base.update(textile_data)
    else:
        furniture_data = await _scrape_furniture(page, url)
        base.update(furniture_data)

    return [base]


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────

async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    test_mode    = os.environ.get("TEST_MODE",         "false").lower() == "true"
    test_limit   = int(os.environ.get("TEST_LIMIT",    "5"))
    scrape_cats  = os.environ.get("SCRAPE_CATEGORIES", "")
    allowed_cats = [c.strip() for c in scrape_cats.split(",") if c.strip()] if scrape_cats else []

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in info["categories"]:
            if not cat["links"]:
                continue
            if allowed_cats and cat["name"] not in allowed_cats:
                continue

            print(f"\n[{cat['group']}] {cat['name']}")

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls: set[str]   = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            print(f"  Found {len(all_product_urls)} products")

            if test_mode:
                all_product_urls = all_product_urls[:test_limit]

            global_idx = 1
            for url in all_product_urls:
                print(f"  [{global_idx}] {url}")
                rows = await scrape_product(page, url, cat["name"])
                for row in rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
                await async_polite_delay()

    writer.save()


if __name__ == "__main__":
    asyncio.run(main())
