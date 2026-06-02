import asyncio, json, os, sys, re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Blueprint Lighting")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
                        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "4"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://blueprintlighting.com"

# Dimension fields — strip to numeric only
DIM_FIELDS = {"Height", "Width", "Depth", "Diameter", "Length",
               "Shade Height", "Backplate Diameter"}

# Full-label spec → column name
SPEC_MAP = {
    "height":               "Height",
    "width":                "Width",
    "depth":                "Depth",
    "diameter":             "Diameter",
    "length":               "Length",
    "overall height":       "Height",
    "shade height":         "Shade Height",
    "backplate diameter":   "Backplate Diameter",
    "canopy diameter":      "Canopy",
    "canopy":               "Canopy",
    # Electrical
    "socket":               "Socket Type",
    "socket type":          "Socket Type",
    "bulb":                 "Bulb Type",
    "bulb type":            "Bulb Type",
    "bulb requirement":     "Bulb Type",
    "wattage":              "Wattage",
    "voltage":              "Voltage",
    "certification":        "Certification",
    "certifications":       "Certification",
    # Physical
    "material":             "Material",
    "materials":            "Materials",
    "finish":               "Finish",
    "finish options":       "Finish",
    # Production
    "lead time":            "Lead Time",
    "production":           "Production",
    "origin":               "Origin",
    "designer":             "Designer",
    "collection":           "Collection",
    "location rating":      "Location Rating",
    "damp location rated":  "Location Rating",
}

# Abbreviated single-letter dimension labels used by some products
# D is ambiguous: Diameter when no Width present; Depth when Width also exists
ABBREV_DIM = {
    "h":   "Height",
    "w":   "Width",
    "l":   "Length",
    "dia": "Diameter",
}


def _make_https(url: str) -> str:
    if not url:
        return url
    url = url.strip()
    if url.startswith("//"):
        return "https:" + url
    if not url.startswith("http"):
        return "https://" + url
    return url


def _upgrade_image(src: str) -> str:
    """Swap any _NxN Shopify size suffix for full resolution."""
    return re.sub(r"_\d+x(\d+)(\.\w+)(\?|$)", r"\2\3", src)


def _num_from_val(val: str) -> str | None:
    """Extract leading number from strings like '16.52 in | 42 cm'. Returns None if no number."""
    m = re.match(r"([0-9]+(?:\.[0-9]+)?)", val.strip())
    return m.group(1) if m else None


async def _fetch_product_api(page, handle: str) -> dict:
    """Fetch /products/<handle>.json?currency=USD — returns the 'product' sub-dict."""
    try:
        resp = await page.evaluate(
            f"async()=>{{const r=await fetch('/products/{handle}.json?currency=USD');"
            f"return await r.json();}}"
        )
        return resp.get("product", {}) if isinstance(resp, dict) else {}
    except Exception as e:
        print(f"  [API] /products/{handle}.json failed: {e}")
        return {}


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return product page URLs via Shopify collection products.json API."""
    slug = listing_url.rstrip("/").split("/collections/")[-1].split("/")[0]
    api_url = f"{BASE_URL}/collections/{slug}/products.json?limit=250"

    links: list[str] = []
    try:
        resp = await page.evaluate(
            f"async()=>{{const r=await fetch('{api_url}');"
            f"return await r.json();}}"
        )
        for p in resp.get("products", []):
            handle = p.get("handle", "")
            if handle:
                links.append(f"{BASE_URL}/collections/{slug}/products/{handle}")
    except Exception as e:
        print(f"  [listing API fallback] {e}")
        await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
        hrefs = await page.evaluate(
            "()=>Array.from(document.querySelectorAll('a[href*=\"/products/\"]'))"
            ".map(a=>a.href)"
        )
        seen: set[str] = set()
        for h in hrefs:
            base_h = h.split("?")[0]
            if base_h not in seen and "/products/" in base_h:
                seen.add(base_h)
                links.append(base_h)

    return links


async def scrape_product(page, url: str) -> list[dict]:
    """Return a list of row dicts — one per variant."""
    clean_url = url.split("?")[0]
    handle    = clean_url.rstrip("/").split("/products/")[-1]

    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2500)

    base: dict = {"Source URL": url, "Manufacturer": VENDOR_NAME}

    # ── 1. Shopify product API — title, body_html, images, variants, options ──
    product = await _fetch_product_api(page, handle)

    variants_data: list[dict] = [{}]
    option_names:  list[str]  = []

    if product:
        base["Product Name"] = clean_text(product.get("title", ""))

        raw_desc = product.get("body_html", "")
        base["Description"] = clean_text(re.sub(r"<[^>]+>", " ", raw_desc))

        # Highest-resolution image from images list
        images = product.get("images", [])
        if images:
            src = images[0].get("src", "") if isinstance(images[0], dict) else str(images[0])
            base["Image URL"] = _make_https(_upgrade_image(src))

        variants_data = product.get("variants", [{}]) or [{}]

        # Option names: [{name, position, values}, ...]
        for opt in product.get("options", []):
            name = opt.get("name", "") if isinstance(opt, dict) else str(opt)
            option_names.append(name)
    else:
        # DOM fallback for product name
        try:
            h1 = await page.query_selector("h1")
            if h1:
                base["Product Name"] = clean_text(await h1.inner_text())
        except Exception:
            pass
        # DOM fallback for image
        try:
            img_el = await page.query_selector(
                ".product__media img, .product-single__media img, "
                ".product__photo img, [class*='product'] img"
            )
            if img_el:
                src = (await img_el.get_attribute("data-src")
                       or await img_el.get_attribute("src") or "")
                base["Image URL"] = _make_https(_upgrade_image(src))
        except Exception:
            pass

    if base.get("Product Name"):
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    # ── 2. Scrape spec section from rendered page ────────────────────────────
    # Open all accordions so metafield content becomes visible
    try:
        await page.evaluate(r"""
            () => {
                // Click any button/summary containing spec-related keywords
                const els = document.querySelectorAll(
                    'button, summary, [role="button"], [aria-expanded]'
                );
                for (const el of els) {
                    const txt = el.textContent || '';
                    if (/spec|dimension|detail|description|features?/i.test(txt)) {
                        try { el.click(); } catch(e) {}
                    }
                }
                // Also open <details> elements
                document.querySelectorAll('details').forEach(d => d.open = true);
            }
        """)
        await page.wait_for_timeout(800)
    except Exception:
        pass

    # Gather text from all metafield elements + product info sections
    spec_text = ""
    try:
        spec_text = await page.evaluate(r"""
            () => {
                const parts = [];
                // All metafield multi-line text fields (specs, dims, electrical)
                document.querySelectorAll('.metafield-multi_line_text_field, .metafield-rich_text_field')
                    .forEach(el => parts.push(el.innerText));
                // Product description / info wrappers
                const sels = [
                    '.product__description', '.product-single__description',
                    '.rte', '.product__info-wrapper', '.product__meta',
                    '.product-meta', 'main'
                ];
                const seen = new Set();
                for (const s of sels) {
                    for (const el of document.querySelectorAll(s)) {
                        if (!seen.has(el)) { seen.add(el); parts.push(el.innerText); }
                    }
                }
                return parts.join('\n');
            }
        """)
    except Exception:
        try:
            spec_text = await page.inner_text("body")
        except Exception:
            pass

    # --- parse key: value or key\tvalue lines ---
    spec_kv: dict[str, str] = {}
    for line in spec_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        # Try colon separator first
        if ":" in line:
            key, _, val = line.partition(":")
            key = key.strip().lower()
            val = val.strip()
            if key and val and len(key) < 60 and len(val) < 300:
                spec_kv[key] = val
        # Try tab separator
        elif "\t" in line:
            parts = line.split("\t", 1)
            if len(parts) == 2:
                key = parts[0].strip().lower()
                val = parts[1].strip()
                if key and val and len(key) < 60 and len(val) < 300:
                    spec_kv[key] = val

    # Apply SPEC_MAP (full labels)
    for src_key, dst_col in SPEC_MAP.items():
        if src_key in spec_kv and not base.get(dst_col):
            raw = spec_kv[src_key]
            if dst_col in DIM_FIELDS:
                num = _num_from_val(raw)
                if num:
                    base[dst_col] = num
            else:
                base[dst_col] = raw

    # --- parse abbreviated labels: H / W / D / L / Dia ---
    abbrev_found: dict[str, str] = {}
    for line in spec_text.split("\n"):
        line = line.strip()
        m = re.match(
            r"^(H|W|L|Dia|D)\s*[:\-]\s*([0-9]+(?:\.[0-9]+)?)\s*in",
            line, re.IGNORECASE
        )
        if m:
            lbl = m.group(1).lower()
            val = m.group(2)
            abbrev_found[lbl] = val

    for lbl, val in abbrev_found.items():
        if lbl in ABBREV_DIM:
            col = ABBREV_DIM[lbl]
            if not base.get(col):
                base[col] = val
        elif lbl == "d":
            # D = Diameter if no Width; else Depth
            col = "Depth" if ("w" in abbrev_found or base.get("Width")) else "Diameter"
            if not base.get(col):
                base[col] = val

    # --- regex fallback: full-label "Label: N in" ---
    # [0-9]+(?:\.[0-9]+)? requires at least one digit — avoids matching lone "."
    for col, pattern in [
        # (?<!Backplate ) prevents "Backplate Diameter" from setting Diameter
        ("Diameter", r"(?<!Backplate )(?<!\w)Diameter\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?)\s*in"),
        ("Height",   r"\bHeight\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?)\s*in"),
        ("Width",    r"\bWidth\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?)\s*in"),
        ("Depth",    r"\bDepth\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?)\s*in"),
        ("Length",   r"\bLength\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?)\s*in"),
    ]:
        if not base.get(col):
            m2 = re.search(pattern, spec_text, re.IGNORECASE)
            if m2:
                base[col] = m2.group(1)

    # --- regex fallback: Blueprint inch-letter format  24" H x 6" W x 24" D ---
    # Also handles: 61.5"h x 11" diameter
    # Restrict to the DIMENSIONS accordion text to avoid false matches in description
    dim_section_m = re.search(
        r'DIMENSIONS?\s*\n(.*?)(?:\n[A-Z]{4,}|\Z)', spec_text, re.DOTALL | re.IGNORECASE
    )
    dim_section = dim_section_m.group(1) if dim_section_m else spec_text

    inch_dims: dict[str, str] = {}
    for m3 in re.finditer(
        r'([0-9]+(?:\.[0-9]+)?)\s*[""]\s*([HhWwDdLl])\b', dim_section
    ):
        val, lbl = m3.group(1), m3.group(2).upper()
        inch_dims.setdefault(lbl, val)  # first occurrence wins

    lbl_to_col = {"H": "Height", "W": "Width", "L": "Length"}
    for lbl, val in inch_dims.items():
        if lbl in lbl_to_col:
            if not base.get(lbl_to_col[lbl]):
                base[lbl_to_col[lbl]] = val
        elif lbl == "D":
            col = "Depth" if "W" in inch_dims else "Diameter"
            if not base.get(col):
                base[col] = val

    # Sentence-style diameter: "11\" diameter at the base"
    # Only for circular fixtures (no Width+Depth already set)
    if not base.get("Diameter") and not (base.get("Width") and base.get("Depth")):
        m4 = re.search(
            r'([0-9]+(?:\.[0-9]+)?)\s*[""]\s*diameter', dim_section, re.IGNORECASE
        )
        if m4:
            base["Diameter"] = m4.group(1)

    # --- damp location ---
    if not base.get("Location Rating") and re.search(
        r"damp\s+location", spec_text, re.IGNORECASE
    ):
        base["Location Rating"] = "Damp Location Rated"

    # --- UL certification fallback ---
    if not base.get("Certification") and re.search(
        r"UL\s*(?:&|and)\s*cUL|UL\s+(?:listed|certified)",
        spec_text, re.IGNORECASE
    ):
        base["Certification"] = "UL & cUL Certified"

    # --- Lead time fallback ---
    if not base.get("Lead Time"):
        m3 = re.search(r"(\d+[-–]\d+\s*weeks?)", spec_text, re.IGNORECASE)
        if m3:
            base["Lead Time"] = m3.group(1)

    # --- socket / bulb from running text ---
    if not base.get("Socket Type"):
        m4 = re.search(r"(E\d+)\s+(?:socket|bulb)", spec_text, re.IGNORECASE)
        if m4:
            base["Socket Type"] = m4.group(1)

    # --- bulb qty ---
    if not base.get("Bulb Qty"):
        m5 = re.search(r"(?:Uses?\s+)?(\w+)\s+\(\d+\)\s+(E\d+)\s+(?:bulb|socket)",
                       spec_text, re.IGNORECASE)
        if m5:
            base["Bulb Qty"] = m5.group(1)

    # ── 3. Tearsheet link ────────────────────────────────────────────────────
    try:
        tearsheets = await page.evaluate(r"""
            () => Array.from(document.querySelectorAll('a[href]'))
                 .filter(a => /tearsheet|tear.?sheet/i.test(a.href + a.textContent))
                 .map(a => a.href)
        """)
        if tearsheets:
            base["Tearsheet Link"] = _make_https(tearsheets[0])
    except Exception:
        pass

    # ── 4. Build one row per variant ─────────────────────────────────────────
    result_rows: list[dict] = []
    for variant in variants_data:
        row = dict(base)

        row["SKU"] = variant.get("sku") or variant.get("SKU") or ""

        # Price from API (USD string like "2450.00" via ?currency=USD)
        price_str = variant.get("price", "")
        if price_str:
            row["Price"] = clean_price(str(price_str))

        # Weight: Shopify stores in grams — convert to lbs
        weight_g = variant.get("grams") or variant.get("weight")
        if weight_g:
            row["Weight"] = round(safe_float(str(weight_g)) / 453.592, 2)

        # Map option slots → column names
        for i, slot in enumerate(["option1", "option2", "option3"], start=1):
            val = variant.get(slot, "")
            if not val or val.lower() in ("default title", "none", ""):
                continue
            col_name = (option_names[i - 1]
                        if i - 1 < len(option_names) else f"Option {i}")
            col_name = col_name.replace("Color / Finish", "Finish").replace("Color/Finish", "Finish")
            row[col_name] = val

        # Variant-specific image
        feat_img = variant.get("featured_image")
        if feat_img and isinstance(feat_img, dict) and feat_img.get("src"):
            row["Image URL"] = _make_https(_upgrade_image(feat_img["src"]))

        # Variant-specific URL
        if len(variants_data) > 1 and variant.get("id"):
            row["Source URL"] = f"{clean_url}?variant={variant['id']}"

        result_rows.append(row)

    return result_rows if result_rows else [base]


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        # Navigate once to blueprintlighting.com so fetch() calls are same-origin
        await page.goto(BASE_URL, timeout=45_000, wait_until="domcontentloaded")

        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    base_u = u.split("?")[0]
                    if base_u not in seen_urls:
                        seen_urls.add(base_u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"  [{cat['name']}] {len(all_product_urls)} products")

            global_idx = 1
            for purl in all_product_urls:
                try:
                    rows = await scrape_product(page, purl)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(
                                info["vendor_name"], cat["name"], global_idx
                            )
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(
                                row["Product Name"]
                            )
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [Error] {purl}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
