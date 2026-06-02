import asyncio, json, os, sys, re
from pathlib import Path

# Windows console may default to charmap — force UTF-8 for safe printing
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Forom")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
BASE_URL = "https://www.foromshop.com"


def _map_option(name: str, value: str = "") -> str:
    n = name.lower().strip()
    if n == "size":
        # Map to Dimensions only when value contains a measurement (has digits)
        # e.g. "5'3\" x 7'7\"" or "10.3\"H" → Dimensions
        # but "Small" / "Large" / "Standard" → Size
        return "Dimensions" if value and re.search(r'\d', value) else "Size"
    m = {"color": "Color", "colour": "Color", "finish": "Finish",
         "material": "Material", "style": "Style"}
    return m.get(n, name)


def extract_spec_fields(text: str) -> dict:
    """
    Extract individual spec fields from DETAILS accordion text.
    Multi-value fields (Dimensions, Seat Height, etc.) collect ALL occurrences
    and join with ' | ' so per-section data (e.g. 75" Sofa / 86.75" Sofa) is preserved.
    """
    result = {}

    # Fields where ALL occurrences are joined (multi-section products like sectionals/sofas)
    multi_patterns = [
        (r'(?<!\w)Dimensions?\s*:\s*([^\n]+)', "Dimensions"),
        (r'Seat [Hh]eight\s*:\s*([^\n]+)',     "Seat Height"),
        (r'Seat [Ww]idth\s*:\s*([^\n]+)',      "Seat Width"),
        (r'Seat [Dd]epth\s*:\s*([^\n]+)',      "Seat Depth"),
        (r'Seat [Ll]ength\s*:\s*([^\n]+)',     "Seat Length"),
        (r'Country of Origin\s*:\s*([^\n]+)',  "Origin"),
        (r'(?<!Country of )Origin\s*:\s*([^\n]+)', "Origin"),
    ]
    for pattern, col in multi_patterns:
        matches = [m.group(1).strip().replace('\xa0', ' ')
                   for m in re.finditer(pattern, text, re.IGNORECASE)]
        matches = [v for v in matches if v]
        if matches:
            result[col] = " | ".join(dict.fromkeys(matches))  # deduplicate preserving order

    # Single-value fields (first match wins)
    single_patterns = [
        (r'Lamp Dimensions?\s*:\s*([^\n]+)', "Lamp Dimensions"),
        (r'Materials?\s*:\s*([^\n]+)', "Material"),
        (r'Finish\s*:\s*([^\n]+)', "Finish"),
        (r'Color\s*:\s*([^\n]+)', "Color"),
        (r'Net\s*Weight\s*:\s*([^\n]+)', "Weight"),
        (r'Gross\s*Weight\s*:\s*([^\n]+)', "Gross Weight"),
        (r'Weight\s*:\s*([\d.]+\s*lbs?[^\n]*)', "Weight"),
        (r'Seat [Tt]hickness\s*:\s*([^\n]+)', "Seat Thickness"),
        (r'Arm [Hh]eight\s*:\s*([^\n]+)', "Arm Height"),
        (r'Backrest [Hh]eight\s*:\s*([^\n]+)', "Backrest Height"),
        (r'Backrest [Ww]idth\s*:\s*([^\n]+)', "Backrest Width"),
        (r'Leg [Tt]hickness\s*:\s*([^\n]+)', "Leg Thickness"),
        (r'Back\s*:\s*([^\n]+)', "Back"),
        (r'Backrest\s*:\s*([^\n]+)', "Backrest"),
        (r'Seat\s*:\s*([^\n]+)', "Seat Construction"),
        (r'Frame\s*:\s*([^\n]+)', "Frame"),
        (r'Feet\s*:\s*([^\n]+)', "Base/Foot Type"),
        (r'Structure\s*:\s*([^\n]+)', "Structure"),
        (r'Springing\s*:\s*([^\n]+)', "Springing"),
        (r'Structure Upholstery\s*:\s*([^\n]+)', "Structure Upholstery"),
        (r'Cover\s*:\s*([^\n]+)', "Cover"),
        (r'Upholstery\s*:\s*([^\n]+)', "Upholstery"),
        # Lighting
        (r'Max\s*Watt(?:age)?\s*:\s*([^\n]+)', "Wattage"),
        (r'Watts?\s*:\s*([^\n]+)', "Wattage"),
        (r'Wattage\s*:\s*([^\n]+)', "Wattage"),
        (r'Voltage\s*:\s*([^\n]+)', "Voltage"),
        (r'VOLTAGE\s*([\w\-–/]+)', "Voltage"),
        (r'Fixture\s*:\s*([^\n]+)', "Socket Type"),
        (r'Socket Type\s*:\s*([^\n]+)', "Socket Type"),
        (r'Socket\s*:\s*([^\n]+)', "Socket Type"),
        (r'Recommended Light Source\s*:\s*([^\n]+)', "Recommended Light Source"),
        (r'(?:Number of )?Bulbs?\s*:\s*([^\n]+)', "Bulb Qty"),
        (r'Bulb Type\s*:\s*([^\n]+)', "Bulb Type"),
        (r'Bulb [Ll]ife\s*:\s*([^\n]+)', "Bulb Life"),
        (r'BULB LIFE\s*:\s*([^\n]+)', "Bulb Life"),
        (r'Lightsource\s*:\s*([^\n]+)', "Light Source"),
        (r'Light Source\s*:\s*([^\n]+)', "Light Source"),
        (r'(\d[\d,–\-]+K)\s*color\s*temp', "Color Temperature"),
        (r'[Cc]olor [Tt]emp(?:erature)?\s*:\s*([^\n]+)', "Color Temperature"),
        (r'Dimmabilit(?:y|ies)\s*:\s*([^\n]+)', "Dimming"),
        (r'Dimm(?:able|ing)\s*:\s*([^\n]+)', "Dimming"),
        (r'Lighting Certification\s*:\s*([^\n]+)', "Lighting Certification"),
        (r'Ceiling Cup\s*:\s*([^\n]+)', "Ceiling Cup"),
        (r'Lampshade\s*(?:&|and)\s*Ceiling Cup\s*:\s*([^\n]+)', "Lampshade Details"),
        (r'Cord [Ll]ength\s*:\s*([^\n]+)', "Cable Length"),
        (r'Cable Length\s*:\s*([^\n]+)', "Cable Length"),
        (r'Cable\s*:\s*([\d.,"\s]+(?:in(?:ches?)?|cm|"|\')?[^\n]*)', "Cable Length"),
        (r'Cord [Mm]aterial\s*:\s*([^\n]+)', "Cord Material"),
        (r'Switch Type\s*:\s*([^\n]+)', "Switch Type"),
        (r'Shade [Dd]imension[s]?\s*:\s*([^\n]+)', "Shade Dimension"),
        (r'Shade [Dd]etails?\s*:\s*([^\n]+)', "Shade Details"),
        (r'IP\s*[Rr]ating\s*:\s*([^\n]+)', "IP Rating"),
        (r'\bIP(\d{2})\b', "IP Rating"),
        (r'Lamp [Bb]ase\s*:\s*([^\n]+)', "Lamp Base"),
        (r'Base\s*:\s*([^\n]+)', "Base"),
        (r'Volume\s*:\s*([^\n]+)', "Volume"),
        (r'Note\s*:\s*([^\n]+)', "Note"),
        (r'Environment\s*:\s*([^\n]+)', "Environment"),
        (r'Product Type\s*:\s*([^\n]+)', "Product Type"),
        (r'Designer\s*:\s*([^\n]+)', "Designer"),
        (r'(?:Care|Care Instructions?)\s*:\s*([^\n]+)', "Care Instructions"),
        # Bed/furniture extras
        (r'Leg [Hh]eight\s*:\s*([^\n]+)', "Leg Height"),
        (r'Footboard [Hh]eight\s*:\s*([^\n]+)', "Footboard Height"),
        (r'Headboard [Hh]eight\s*:\s*([^\n]+)', "Headboard Height"),
        (r'Storage [Ss]pace\s*:\s*([^\n]+)', "Storage Space"),
        (r'Cleaning [Cc]ode\s*:\s*([^\n]+)', "Cleaning Code"),
        (r'Lead [Tt]ime\s*:\s*([^\n]+)', "Lead Time"),
        (r'Collection\s*:\s*([^\n]+)', "Collection"),
        (r'Style\s*:\s*([^\n]+)', "Style"),
    ]
    for pattern, col in single_patterns:
        if col in result:
            continue  # already captured by multi_patterns
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip().replace('\xa0', ' ')
            if val:
                result[col] = val

    # Assembly: key:value format (e.g. "Assembly: Required")
    asm = re.search(r'Assembly\s*:\s*([^\n]+)', text, re.IGNORECASE)
    if asm:
        result.setdefault("Assembly Required", asm.group(1).strip())

    # Boolean badge fields — standalone text with no colon
    for badge, col in [
        ("Assembly Required", "Assembly Required"),
        ("Hospitality Approved", "Hospitality Approved"),
        ("Handmade piece", "Production Detail"),
        ("Handmade in Italy", "Production Detail"),
        ("Handmade in Spain", "Production Detail"),
        ("Exclusively for Forom", "Exclusivity"),
    ]:
        if re.search(rf'\b{re.escape(badge)}\b', text, re.IGNORECASE):
            result.setdefault(col, badge)

    # Upholstery Fabric Information block
    fabric_match = re.search(
        r'Upholstery Fabric Information\s*(.*)', text, re.DOTALL | re.IGNORECASE
    )
    if fabric_match:
        fabric_text = fabric_match.group(1).strip()
        if fabric_text and len(fabric_text) > 10:
            result["Fabric Details"] = fabric_text

    return result


async def get_shopify_json(page, product_url: str = "") -> dict:
    """
    Fetch complete Shopify product data.
    Primary: /products/{handle}.json storefront API (returns title, vendor, variants, images).
    Fallback: window globals / script tags.
    """
    # Primary: storefront JSON API
    if product_url:
        handle = product_url.split("/products/")[-1].split("?")[0]
        api_url = f"{BASE_URL}/products/{handle}.json"
        try:
            data = await page.evaluate("""async (url) => {
                try {
                    const r = await fetch(url, {credentials: 'include'});
                    if (!r.ok) return null;
                    return await r.json();
                } catch(e) { return null; }
            }""", api_url)
            if data and isinstance(data, dict) and "product" in data:
                return data["product"]
        except Exception:
            pass

    # Fallback: window globals
    for expr in ["window.productJSON", "window.__product__",
                 "window.theme && window.theme.product"]:
        try:
            r = await page.evaluate(
                f"(function(){{try{{return {expr};}}catch(e){{return null;}}}})()"
            )
            if r and isinstance(r, dict) and "variants" in r:
                return r
        except Exception:
            pass

    return {}


async def _dismiss_cookie_banner(page):
    """Click Accept on cookie consent banner if present."""
    try:
        for sel in [
            "button:has-text(\"ACCEPT\")",
            "button:has-text(\"Accept All\")",
            "button:has-text(\"Accept\")",
            "#onetrust-accept-btn-handler",
            "button[class*=\"accept\"]",
        ]:
            try:
                btn = page.locator(sel).first
                if await btn.is_visible(timeout=1500):
                    await btn.click()
                    await page.wait_for_timeout(600)
                    return
            except Exception:
                pass
    except Exception:
        pass


def _parse_body_text(body: str) -> dict:
    """
    Parse visible page text into field dict.
    Strips accordion icons (▼▶►◄) so headings are clean before regex matching.
    """
    # Remove accordion arrow symbols from every line so patterns match cleanly
    cleaned_lines = [re.sub(r'[▶▼►◄▸▾]+', '', ln).rstrip() for ln in body.split('\n')]
    body = '\n'.join(cleaned_lines)

    d = {}

    # --- Inline center labels: DESIGNER / DIMENSIONS / MATERIALS / PRODUCTION DETAIL ---
    # Label is on its own line; capture multi-line values until next ALL-CAPS label or blank separator
    for label, col in [
        ("DESIGNER", "Designer"),
        ("DIMENSIONS", "Dimensions"),
        ("MATERIALS", "Material"),
        ("PRODUCTION DETAIL", "Production Detail"),
    ]:
        m = re.search(
            rf'^{label}\s*$\s*((?:[^\n]+\n?)+?)(?=\n{{2,}}|^[A-Z]{{3,}}|^DETAILS|^DESIGNER|^CARE|\Z)',
            body, re.MULTILINE
        )
        if m:
            val = m.group(1).strip()
            if val and len(val) < 600:
                d[col] = val

    # --- DETAILS accordion section ---
    det = re.search(
        r'^DETAILS\s*$\s*(.*?)(?=^DESIGNERS|^CARE|^SHIPPING|^FINANCING|\Z)',
        body, re.DOTALL | re.MULTILINE | re.IGNORECASE
    )
    if det:
        content = det.group(1).strip()
        if content:
            d["__details__"] = content

    # --- DESIGNERS' NOTES ---
    notes = re.search(
        r"^DESIGNERS[^\n]*$\s*(.*?)(?=^CARE|^SHIPPING|^FINANCING|\Z)",
        body, re.DOTALL | re.MULTILINE | re.IGNORECASE
    )
    if notes:
        text = notes.group(1).strip()
        if text and len(text) > 10:
            d["Description"] = text

    # --- CARE ---
    care = re.search(
        r'^CARE\s*$\s*(.*?)(?=^SHIPPING|^FINANCING|\Z)',
        body, re.DOTALL | re.MULTILINE | re.IGNORECASE
    )
    if care:
        text = care.group(1).strip()
        if text and len(text) > 5:
            d["Care Instructions"] = text

    return d


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape one Forom product page. Returns list of dicts (one per variant)."""
    base_url = url.split("?")[0]
    await page.goto(base_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    # Dismiss cookie banner (appears on first visit)
    await _dismiss_cookie_banner(page)

    # Shopify product JSON via storefront API — reliable source for title/vendor/images/variants
    shopify = await get_shopify_json(page, base_url)

    # Expand collapsed accordions using offsetHeight check (safe — won't close already-open ones)
    # Targets: DETAILS, DESIGNERS' NOTES, CARE
    await page.evaluate("""() => {
        const targets = ['DETAILS', 'DESIGNERS', 'CARE'];
        document.querySelectorAll('button.collapsible-trigger, button[aria-expanded], .collapsible__button').forEach(btn => {
            const txt = (btn.textContent || '').trim().toUpperCase();
            if (!targets.some(t => txt.startsWith(t))) return;
            // Check if content sibling is currently hidden (height 0 or display none)
            const parent = btn.closest('.collapsible, [class*="collapsible-wrap"], [class*="product-"]') || btn.parentElement;
            const content = parent && parent.querySelector('[class*="collapsible-content"], [class*="collapsible-body"]');
            const isHidden = content
                ? (content.offsetHeight === 0 || window.getComputedStyle(content).display === 'none')
                : btn.getAttribute('aria-expanded') === 'false';
            if (isHidden) btn.click();
        });
    }""")
    await page.wait_for_timeout(1000)

    # Click Read more if DESIGNERS' NOTES is truncated
    try:
        read_more = page.locator("a:has-text('Read More'), button:has-text('Read More'), a:has-text('Read more')").first
        if await read_more.is_visible(timeout=1000):
            await read_more.click()
            await page.wait_for_timeout(600)
    except Exception:
        pass

    body_text = await page.inner_text("body")
    parsed = _parse_body_text(body_text)

    # Tearsheet / Spec Sheet links (must use JS to get href)
    js_data = await page.evaluate("""
        () => {
            const d = {};
            const links = Array.from(document.querySelectorAll('a'));

            // Tearsheet / Download a quote link
            const quoteEl = links.find(
                a => (a.innerText || '').trim().toLowerCase().includes('download a quote') ||
                     (a.href && /\\.pdf/i.test(a.href) && a.href.includes('cdn'))
            );
            if (quoteEl) d.tearsheet_url = quoteEl.href;

            // Spec Sheet link (separate PDF)
            const specEl = links.find(
                a => (a.innerText || '').trim().toLowerCase().includes('spec sheet') ||
                     (a.textContent || '').trim().toLowerCase() === 'spec sheet'
            );
            if (specEl && specEl.href !== d.tearsheet_url) d.spec_sheet_url = specEl.href;

            // All other CDN file links (3D files)
            const excluded = new Set([d.tearsheet_url, d.spec_sheet_url].filter(Boolean));
            d.cdn_links = links
                .filter(a => a.href && a.href.includes('cdn/shop/files') && !excluded.has(a.href))
                .map(a => a.href);
            return d;
        }
    """)

    base = {"Source URL": base_url, "Manufacturer": VENDOR_NAME}

    # Product name — from API (reliable)
    title = shopify.get("title", "")
    if not title:
        el = await page.query_selector(".product-single__title, h1")
        if el:
            title = clean_text(await el.inner_text())
    if title:
        base["Product Name"] = clean_text(title)
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    # Brand/vendor — store when it differs from Forom (e.g. OWL, FAINA)
    vendor = shopify.get("vendor", "")
    if vendor and vendor.lower() not in ("forom", ""):
        base["Brand"] = vendor

    # Fields from body text parsing
    for col in ("Designer", "Dimensions", "Material", "Production Detail",
                "Description", "Care Instructions"):
        if parsed.get(col):
            base[col] = clean_text(parsed[col])

    # DETAILS accordion content — extract individual spec fields
    details_text = parsed.get("__details__", "")
    if details_text:
        spec_fields = extract_spec_fields(details_text)
        for col, val in spec_fields.items():
            if col == "Weight":
                numeric = re.sub(r'[^\d.]', '', val)
                base.setdefault("Weight", numeric if numeric else val)
            elif col == "Dimensions":
                base.setdefault("Dimensions", val)
            else:
                base.setdefault(col, val)

        # Build Specifications string from all key:value lines
        spec_lines = [ln.strip() for ln in details_text.split("\n")
                      if ln.strip() and ":" in ln]
        if spec_lines:
            base.setdefault("Specifications", " | ".join(spec_lines))
        elif details_text.strip():
            base.setdefault("Specifications", details_text.strip())

    # Tearsheet / quote link
    if js_data.get("tearsheet_url"):
        base["Tearsheet Link"] = js_data["tearsheet_url"]

    # Spec Sheet link (separate from tearsheet)
    if js_data.get("spec_sheet_url"):
        base["Spec Sheet"] = js_data["spec_sheet_url"]

    # CDN file links — only non-image files (3D models, CAD, ZIP, etc.)
    image_exts = re.compile(r'\.(png|jpg|jpeg|webp|gif|svg|avif)(\?|$)', re.IGNORECASE)
    cdn = [h for h in js_data.get("cdn_links", []) if not image_exts.search(h)]
    if cdn:
        base.setdefault("3D Model", "; ".join(cdn))

    # --- Build variant rows from Shopify JSON ---
    variants  = shopify.get("variants", [])
    options   = shopify.get("options", [])

    option_names = []
    for o in options:
        option_names.append(o.get("name", "") if isinstance(o, dict) else str(o))

    # Build variant-id → image-src lookup (strip query params for clean URL)
    all_images = shopify.get("images", [])
    images_by_variant: dict[int, str] = {}
    for img in all_images:
        src = img.get("src", "").split("?")[0]
        for vid in (img.get("variant_ids") or []):
            images_by_variant[vid] = src

    # Default image = first product image
    default_image = all_images[0].get("src", "").split("?")[0] if all_images else ""

    if not variants:
        if default_image:
            base["Image URL"] = default_image
        return [base]

    rows = []
    for v in variants:
        row = dict(base)

        row["SKU"] = v.get("sku") or ""

        # Forom's storefront API returns prices in cents (e.g. "74800.00" = $748.00)
        price_raw = v.get("price", 0) or 0
        try:
            row["Price"] = round(float(price_raw) / 100, 2) if price_raw else None
        except (ValueError, TypeError):
            row["Price"] = clean_price(str(price_raw))

        # Compare-at / original price (before sale)
        compare_raw = v.get("compare_at_price") or 0
        try:
            if compare_raw:
                row["Compare Price"] = round(float(compare_raw) / 100, 2)
        except (ValueError, TypeError):
            pass

        # Variant-specific image
        v_img = images_by_variant.get(v.get("id", 0), "")
        row["Image URL"] = v_img if v_img else default_image

        # Variant URL
        row["Source URL"] = f"{base_url}?variant={v['id']}"

        # Map option values to standard column names (skip "Default Title" placeholder)
        for i, opt_val in enumerate([v.get("option1"), v.get("option2"), v.get("option3")]):
            if opt_val and opt_val != "Default Title" and i < len(option_names):
                col = _map_option(option_names[i], opt_val)
                row.setdefault(col, opt_val)

        rows.append(row)

    return rows if rows else [base]


async def _links_via_api(page, handle: str, product_types: list[str] | None = None) -> list[str]:
    """Use Shopify collection products.json API via fetch().
    If product_types given, only include products whose product_type matches one of them.
    """
    links = []
    pt_lower = [t.lower() for t in product_types] if product_types else []
    page_num = 1
    while True:
        api_url = f"{BASE_URL}/collections/{handle}/products.json?limit=250&page={page_num}"
        try:
            data = await page.evaluate(f"""
                async () => {{
                    try {{
                        const r = await fetch('{api_url}', {{credentials: 'include'}});
                        if (!r.ok) return null;
                        return await r.json();
                    }} catch(e) {{ return null; }}
                }}
            """)
        except Exception as e:
            print(f"  API fetch error: {e}")
            break
        if not data:
            break
        products = data.get("products", [])
        if not products:
            break
        for p in products:
            if pt_lower:
                pt = p.get("product_type", "").lower()
                if not any(pt == f.lower() or f.lower() in pt or pt in f.lower() for f in pt_lower):
                    continue
            links.append(f"{BASE_URL}/products/{p['handle']}")
        print(f"  API page {page_num}: {len(products)} products (total {len(links)})")
        if len(products) < 250:
            break
        page_num += 1
    return links


async def _links_via_html(page, listing_url: str) -> list[str]:
    """HTML scraping for collection pages — handles Boost Commerce dynamic rendering."""
    links = []
    seen  = set()
    page_num = 1

    # For paginated plain URLs (no existing ?) use ?page=N; otherwise &page=N
    sep = "&" if "?" in listing_url else "?"

    while True:
        paginated = f"{listing_url}{sep}page={page_num}"
        await page.goto(paginated, timeout=60_000, wait_until="domcontentloaded")
        # Wait for Boost Commerce / JS to finish rendering product cards
        try:
            await page.wait_for_selector('a[href*="/products/"]', timeout=8000)
        except Exception:
            pass
        await page.wait_for_timeout(2000)
        await _dismiss_cookie_banner(page)

        hrefs = await page.evaluate("""
            () => [...new Set(
                Array.from(document.querySelectorAll('a[href*="/products/"]'))
                    .map(a => a.href.split('?')[0])
                    .filter(h => {
                        try {
                            const p = new URL(h).pathname;
                            return p.startsWith('/products/') && p.length > '/products/'.length;
                        } catch(e) { return false; }
                    })
            )]
        """)

        new_count = 0
        for h in hrefs:
            if h not in seen:
                seen.add(h)
                links.append(h)
                new_count += 1

        print(f"  HTML page {page_num}: {new_count} new links (total {len(links)})")
        if new_count == 0:
            break
        page_num += 1
        await async_polite_delay()

    return links


def _parse_boost_types(url: str) -> list[str]:
    """Extract pf_pt_product_type values from a Boost Commerce filter URL."""
    from urllib.parse import urlparse, parse_qs, unquote_plus
    qs = parse_qs(urlparse(url).query)
    return [unquote_plus(t) for t in qs.get("pf_pt_product_type", [])]


async def get_product_links(page, listing_url: str) -> list[str]:
    """Route listing URL to the best scraping strategy:
    - Boost filter URL (?pf_pt_product_type=...) → API on base handle with type filter
    - Plain collection URL → API; fall back to HTML if API returns < 5 results
    """
    if "pf_pt_product_type" in listing_url:
        # Outdoor and other Boost-filtered collections
        handle = listing_url.split("?")[0].rstrip("/").split("/collections/")[-1]
        product_types = _parse_boost_types(listing_url)
        links = await _links_via_api(page, handle, product_types)
        if not links:
            # Fallback: HTML scraping on the filtered URL
            links = await _links_via_html(page, listing_url)
        return links
    elif "?" not in listing_url:
        handle = listing_url.rstrip("/").split("/collections/")[-1]
        links = await _links_via_api(page, handle)
        # If API returned suspiciously few results, cross-check with HTML
        if len(links) < 5:
            html_links = await _links_via_html(page, listing_url)
            existing = set(links)
            for h in html_links:
                if h not in existing:
                    links.append(h)
                    existing.add(h)
        return links
    else:
        return await _links_via_html(page, listing_url)


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    filter_cats = [c.strip() for c in os.environ.get("SCRAPE_CATEGORIES", "").split(",") if c.strip()]
    # TEST_MODE / TEST_MAX_PRODUCTS (set by orchestrator --test flag) or TEST_LIMIT (manual)
    test_limit: int | None = None
    if os.environ.get("TEST_MODE", "").lower() == "true":
        test_limit = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))
        max_cats   = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
    else:
        tl = os.environ.get("TEST_LIMIT", "")
        test_limit = int(tl) if tl.isdigit() else None
        max_cats   = None

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        # Accept cookies once at session start so it doesn't block any page
        await page.goto(BASE_URL, timeout=30_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
        await _dismiss_cookie_banner(page)

        cats_done = 0
        for cat in info["categories"]:
            if not cat["links"]:
                continue
            if filter_cats and cat["name"] not in filter_cats:
                continue
            if max_cats and cats_done >= max_cats:
                break
            cats_done += 1

            writer.add_sheet(cat["name"], cat["links"][0], studio_columns=cat["studio_columns"])
            print(f"\n=== [{cat['group']}] {cat['name']} ===")

            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            # Apply per-category product limit when TEST_LIMIT is set
            if test_limit:
                all_product_urls = all_product_urls[:test_limit]

            print(f"  Total unique products: {len(all_product_urls)}")

            global_idx = 1
            for url in all_product_urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        row["Manufacturer"] = VENDOR_NAME
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                    slug = url.split("/products/")[-1]
                    print(f"  [{global_idx - len(rows)}] {slug} -> {len(rows)} row(s)")
                except Exception as e:
                    print(f"  ERROR {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"\nDone. Output: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
