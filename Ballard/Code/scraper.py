import asyncio, json, os, re, sys
from pathlib import Path
from playwright.async_api import async_playwright

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME      = os.environ.get("VENDOR_NAME", "Ballard")
HEADLESS         = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH      = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE        = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CAT     = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PROD    = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL  = "https://www.ballarddesigns.com"
UA        = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
             "AppleWebKit/537.36 (KHTML, like Gecko) "
             "Chrome/124.0.0.0 Safari/537.36")

# Category name corrections per CLAUDE.md standards
CAT_NAME_FIX = {
    "Flush Mount":  "Flush Mounted",
    "Table Lamp":   "Table Lamps",
    "Floor Lamp":   "Floor Lamps",
}


# ── helpers ──────────────────────────────────────────────────────────────────

def build_image_url(sku: str, finish_code: str) -> str:
    """High-res product image URL for a given SKU + finish code."""
    src = f"ballarddesigns/{sku}_{finish_code.upper()}" if finish_code else f"ballarddesigns/{sku}_main"
    return (
        f"https://akamai-scene7.ballarddesigns.com/is/image/ballarddesigns/"
        f"T_Template?$UPDP_hero$&$src={src}&defaultImage={sku}_main"
    )


def _convert_fractions(s: str) -> str:
    """Convert mixed fractions like '28 1/2' to decimals like '28.5'."""
    def _repl(m):
        whole = float(m.group(1)) if m.group(1) else 0
        return str(round(whole + int(m.group(2)) / int(m.group(3)), 4)).rstrip("0").rstrip(".")
    s = re.sub(r'(\d+)\s+(\d+)/(\d+)', _repl, s)
    s = re.sub(r'\b(\d+)/(\d+)\b', lambda m: str(round(int(m.group(1)) / int(m.group(2)), 4)).rstrip("0").rstrip("."), s)
    return s


def _parse_hwdv(s: str) -> tuple:
    """Extract (height, width, depth) strings from a dimension value like '20"H X 79.5"W X 21.5"D'."""
    h = re.search(r'([\d.]+)\s*"?\s*H\b', s, re.I)
    w = re.search(r'([\d.]+)\s*"?\s*W\b', s, re.I)
    d = re.search(r'([\d.]+)\s*"?\s*D\b', s, re.I)
    return (h.group(1) if h else None, w.group(1) if w else None, d.group(1) if d else None)


def parse_spec_block(text: str) -> dict:
    """
    Parse the text of div.prodSpecContent into structured fields.

    Captures Overall dims + sub-dimension lines (Seat/Arms/Footrest for seating;
    Backplate/Shade/Canopy/Chain/Hanging Length for lighting) plus Construction,
    Lighting section (Wattage/Bulb Type/Wiring Type), Country of Origin, Additional Info.

    Returns _size_dims (internal key) when size-prefixed Overall lines are present
    (King/Queen/Twin/Full beds) so scrape_product can assign the right dims per variant.
    """
    data = {}

    # Weight — (56 lbs) inside any Overall line
    weight_m = re.search(r'Overall[^:\n]*:.*?\((\d+(?:\.\d+)?)\s*lbs?\)', text, re.I)
    if weight_m:
        data["Weight"] = weight_m.group(1)

    # Size-prefixed Overall dims (King/Queen/Twin/Full beds)
    # e.g. "King Overall: 57 1/2"H X 81 1/2"W X 85 1/4"D"
    SIZE_PREFIXES = r'(?:King|Queen|Twin|Full|Cal\.?\s*King)'
    size_dims: dict[str, str] = {}
    for m in re.finditer(rf'({SIZE_PREFIXES})\s+Overall[^:\n]*:\s*([^\n]+)', text, re.I):
        sz = m.group(1).strip().title().replace("Cal.", "Cal").replace("Cal ", "Cal. ")
        dims_str = re.sub(r'\s*\([^)]+\)', '', m.group(2)).strip()
        dims_str = _convert_fractions(dims_str)
        if dims_str:
            size_dims[sz] = dims_str
    if size_dims:
        data["_size_dims"] = size_dims

    # Plain Overall dimensions — first "Overall:" line NOT preceded by a size word
    plain_overall = None
    for m in re.finditer(r'^(.*?)(Overall[^:\n]*:\s*)([^\n]+)', text, re.I | re.MULTILINE):
        prefix = m.group(1).strip()
        if not re.search(r'(?:King|Queen|Twin|Full|Cal)', prefix, re.I):
            plain_overall = m
            break
    if plain_overall:
        dims_str = re.sub(r'\s*\([^)]+\)', '', plain_overall.group(3)).strip()
        dims_str = re.sub(r'(\d[\d.]*"H)[^X]*&[^X]*\d[\d.]*"H', r'\1', dims_str)
        dims_str = _convert_fractions(dims_str)
        if dims_str:
            data["Dimensions"] = dims_str
    elif size_dims:
        # Use first size as default Dimensions when no plain Overall exists
        data["Dimensions"] = next(iter(size_dims.values()))

    # Sub-dimension lines inside the Dimensions section
    dim_section_m = re.search(
        r'Dimensions:\s*\n(.*?)(?=Construction:|Country of Origin:|Lighting:|$)',
        text, re.I | re.DOTALL
    )
    if dim_section_m:
        for line in dim_section_m.group(1).split("\n"):
            line = line.strip()
            if not line:
                continue

            # Lines without ":" — handle standalone thickness (rugs: "Approx. 1/4" Thick")
            if ":" not in line:
                m_thick = re.match(r'Approx\.\s*([\d][^"\n]*)"?\s*Thick\b', line, re.I)
                if m_thick:
                    val = _convert_fractions(m_thick.group(1).strip().rstrip('"'))
                    data.setdefault("Pile Height", val + '"')
                continue

            label, _, value = line.partition(":")
            label_key = label.strip().lower()
            value_raw = value.strip()
            value = _convert_fractions(value_raw)
            h, w, d = _parse_hwdv(value)

            # Seating sub-dims
            if label_key in ("seat", "seat dimensions"):
                if h: data["Seat Height"] = h
                if w: data["Seat Width"]  = w
                if d: data["Seat Depth"]  = d
            elif label_key in ("arm", "arms", "arm height"):
                if h: data["Arm Height"] = h
            elif "footrest" in label_key:
                if h: data["Footrest"] = h + '"H'

            # Lighting sub-dims
            elif label_key == "backplate":
                data["Canopy"] = value_raw
            elif label_key in ("ceiling canopy", "canopy"):
                data["Canopy"] = value_raw
            elif label_key == "shade":
                data["Shade Details"] = value_raw
            elif label_key == "chain":
                data["Chain Length"] = value_raw
            elif "hanging length" in label_key:
                # "Min/Max Hanging Length: 24 3/8" & 96 3/8""
                parts = re.split(r'\s*&\s*', value_raw)
                data["Min Drop"]      = _convert_fractions(parts[0].strip())
                data["Hanging Length"] = _convert_fractions(parts[-1].strip())

    # Construction → Material
    mat_m = re.search(r'Construction:\s*\n?(.*?)(?:\n[A-Z][^\n:]+:|\Z)', text, re.I | re.DOTALL)
    if mat_m:
        mat = clean_text(mat_m.group(1))
        if mat:
            data["Material"] = mat

    # Lighting section → Wattage, Bulb Type, Wiring Type
    lighting_m = re.search(r'Lighting:\s*\n?([^\n]+(?:\n(?![A-Z][^\n:]+:)[^\n]+)*)', text, re.I)
    if lighting_m:
        lt = lighting_m.group(1).strip()
        watt_m = re.search(r'(\d+(?:\.\d+)?)\s*W\b', lt, re.I)
        if watt_m:
            data["Wattage"] = watt_m.group(1) + "W"
        btype_m = re.search(r'type\s+([A-Z]\d*)\b', lt, re.I)
        if btype_m:
            data["Bulb Type"] = "Type " + btype_m.group(1).upper()
        if re.search(r'\bhardwire\b', lt, re.I):
            data["Wiring Type"] = "Hardwire"
        elif re.search(r'\bplug.?in\b', lt, re.I):
            data["Wiring Type"] = "Plug-in"

    # Country of Origin
    orig_m = re.search(r'Country of Origin:\s*\n?([^\n]+)', text, re.I)
    if orig_m:
        orig = clean_text(orig_m.group(1))
        if orig:
            data["Origin"] = orig

    # Additional Info — multi-line, route to Care Instructions (rugs) or Assembly Required
    info_m = re.search(r'Additional Info:\s*\n?(.*?)(?=\n[A-Z][^\n]*:|\Z)', text, re.I | re.DOTALL)
    if info_m:
        info = clean_text(info_m.group(1))
        if info:
            if re.search(r'\b(vacuum|sweep|rinse|clean|wash|care|dry)\b', info, re.I):
                data["Care Instructions"] = info
            else:
                data["Assembly Required"] = info
            # Also set Assembly Required flag when the text explicitly says so
            if re.search(r'\bassembly required\b', info, re.I):
                data["Assembly Required"] = "Assembly required"

    return data


async def get_json_ld_products(page) -> list[dict]:
    """Return all JSON-LD Product objects on the current page."""
    scripts = await page.eval_on_selector_all(
        'script[type="application/ld+json"]',
        'els => els.map(e => e.textContent)',
    )
    products = []
    for s in scripts:
        try:
            obj = json.loads(s)
            if obj.get("@type") == "Product":
                products.append(obj)
        except Exception:
            pass
    return products


async def _reset_page(page) -> None:
    """Cancel any pending navigation and navigate to about:blank to reset page state."""
    try:
        await page.evaluate("() => window.stop()")
    except Exception:
        pass
    try:
        await page.goto("about:blank", timeout=8_000, wait_until="domcontentloaded")
        await page.wait_for_timeout(500)
    except Exception:
        pass


# ── listing ──────────────────────────────────────────────────────────────────

async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Return deduplicated canonical product URLs from all pages of a listing.
    Each JSON-LD Product entry on the listing page is a colour/finish variant;
    we deduplicate by the canonical URL (same base ID, different finishes share it).
    """
    seen: set[str] = set()
    urls: list[str] = []

    await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
    try:
        await page.wait_for_selector("article.c-card", timeout=10_000)
    except Exception:
        pass  # some pages render slowly; proceed anyway
    await page.wait_for_timeout(1_000)

    while True:
        for obj in await get_json_ld_products(page):
            canonical = obj.get("url", "").split("?")[0]
            if canonical and canonical not in seen:
                seen.add(canonical)
                urls.append(canonical)

        # Pagination: "Page X of Y"
        pag_el = await page.query_selector(".c-pagination__page-count")
        pag_text = (await pag_el.inner_text()).strip() if pag_el else ""
        m = re.match(r"Page\s+(\d+)\s+of\s+(\d+)", pag_text, re.I)
        if not m or int(m.group(1)) >= int(m.group(2)):
            break

        next_btn = await page.query_selector(".c-pagination__button__next:not([disabled])")
        if not next_btn:
            break
        await next_btn.click()
        try:
            await page.wait_for_selector("article.c-card", timeout=8_000)
        except Exception:
            pass
        await page.wait_for_timeout(1_000)

    return urls


# ── product detail ────────────────────────────────────────────────────────────

async def scrape_product(page, url: str) -> list[dict]:
    """
    Return one dict per finish/size variant found on the product detail page.

    Variant strategy:
      1. Read all swatch buttons (div.c-universal-options__option-swatch-container)
         → finish code from data-cs-override-id, finish name from span.visually-hidden
         → image constructed from SKU + finish code (no clicking needed)
      2. If no swatches → single-row fallback using the main JSON-LD image.

    Fields shared across variants: Product Name, Product Family Id, SKU, Price,
        Description, Specifications, Material, Origin, Assembly Required,
        Dimensions, Height, Width, Depth, Weight.
    Fields that change per variant: Finish, Image URL.
    """
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2_500)

    base: dict = {"Source URL": url}

    # ── JSON-LD ──
    ld_obj = None
    for obj in await get_json_ld_products(page):
        ld_obj = obj
        break

    if ld_obj:
        base["Product Name"] = clean_text(ld_obj.get("name", ""))
        base["SKU"]          = ld_obj.get("sku") or ld_obj.get("mfPartNumber", "")
        base["Product Family Id"] = extract_family_id(base["Product Name"])

        imgs = ld_obj.get("image", [])
        base["_main_image"] = imgs[0] if isinstance(imgs, list) and imgs else (imgs or "")

        offers = ld_obj.get("offers", {})
        price_raw = offers.get("lowPrice") or offers.get("price", "")
        if price_raw:
            base["Price"] = safe_float(str(price_raw))

    # ── Description (first substantial paragraph) ──
    for p_el in await page.query_selector_all("p"):
        txt = clean_text(await p_el.inner_text())
        if len(txt) > 80 and not txt.lower().startswith("sign up"):
            # Strip trailing "features:" label if present
            txt = re.sub(r'\s+[\w\s]+features?\s*:?\s*$', '', txt, flags=re.I).strip()
            base["Description"] = txt
            break

    # ── Feature bullets ──
    # The feature list is an unclassed <ul> with 3+ items; skip nav/footer lists
    SKIP_CLS = {"navigation", "footer", "search", "carousel", "pip", "suggestion", "second-level"}
    for ul_el in await page.query_selector_all("ul"):
        ul_cls = (await ul_el.get_attribute("class") or "").lower()
        if any(s in ul_cls for s in SKIP_CLS):
            continue
        items = await ul_el.query_selector_all("li")
        texts = [clean_text(await li.inner_text()) for li in items]
        texts = [t for t in texts if len(t) > 4]
        if len(texts) >= 3:
            spec_str = "; ".join(texts)
            base["Specifications"] = spec_str
            # Extract lighting fields from bullets when not in the spec block
            # e.g. "150W max type A bulb; 3-way switch; 12' black plug-in cord"
            if not base.get("Wattage"):
                wm = re.search(r'(\d+(?:\.\d+)?)\s*W\b', spec_str, re.I)
                if wm:
                    base["_bullet_wattage"] = wm.group(1) + "W"
            if not base.get("Bulb Type"):
                bm = re.search(r'type\s+([A-Z]\d*)\b', spec_str, re.I)
                if bm:
                    base["_bullet_bulbtype"] = "Type " + bm.group(1).upper()
            if not base.get("Wiring Type"):
                if re.search(r'\bhardwire\b', spec_str, re.I):
                    base["_bullet_wiring"] = "Hardwire"
                elif re.search(r'\bplug.?in\b', spec_str, re.I):
                    base["_bullet_wiring"] = "Plug-in"
            break

    # ── Spec block (Dimensions tab content) ──
    # Click the DIMENSIONS tab first — some products (floor lamps, pendants) lazy-load it
    try:
        all_tabs = await page.query_selector_all("button, [role='tab'], li.c-tabs__item a")
        for tab_el in all_tabs:
            tab_txt = (await tab_el.inner_text()).strip().upper()
            if "DIMENSION" in tab_txt:
                await tab_el.click()
                await page.wait_for_timeout(800)
                break
    except Exception:
        pass

    size_dims: dict[str, str] = {}
    # Use JS innerText so hidden elements (display:none) are also read
    spec_text: str = await page.evaluate("""() => {
        const el = document.querySelector('div.prodSpecContent');
        return el ? el.innerText : '';
    }""") or ""
    if not spec_text.strip():
        # Fallback: scan all visible tab panels for 'Overall'
        spec_text = await page.evaluate("""() => {
            const candidates = document.querySelectorAll(
                '[role="tabpanel"], .c-tabs__panel, .tab-content, .prodSpec, .product-spec'
            );
            for (const el of candidates) {
                const t = el.innerText || '';
                if (t.includes('Overall') || t.includes('Dimensions')) return t;
            }
            return '';
        }""") or ""
    if spec_text.strip():
        spec_data = parse_spec_block(spec_text)
        size_dims = spec_data.pop("_size_dims", {})
        base.update(spec_data)

    # Apply bullet-derived lighting fields only if spec block didn't provide them
    if not base.get("Wattage")   and base.get("_bullet_wattage"):  base["Wattage"]     = base["_bullet_wattage"]
    if not base.get("Bulb Type") and base.get("_bullet_bulbtype"): base["Bulb Type"]   = base["_bullet_bulbtype"]
    if not base.get("Wiring Type") and base.get("_bullet_wiring"): base["Wiring Type"] = base["_bullet_wiring"]
    for k in ("_bullet_wattage", "_bullet_bulbtype", "_bullet_wiring"):
        base.pop(k, None)

    # ── Swatches → one variant row per finish ──
    swatch_containers = await page.query_selector_all(
        "div.c-universal-options__option-swatch-container"
    )

    variants: list[dict] = []
    sku = base.get("SKU", "")

    for container in swatch_containers:
        btn = await container.query_selector("button")
        if not btn:
            continue

        # Finish code: data-cs-override-id="pdp_finish_dpw" → "DPW"
        override_id = await btn.get_attribute("data-cs-override-id") or ""
        m = re.search(r"pdp_\w+_(\w+)$", override_id)
        finish_code = m.group(1).upper() if m else ""

        # Finish name: span.visually-hidden → "Deep Walnut swatch 1 of 5"
        # Use textContent (not innerText) to bypass CSS text-transform: uppercase
        vh = await btn.query_selector("span.visually-hidden")
        finish_name = ""
        if vh:
            vh_text = (await vh.evaluate("el => el.textContent") or "").strip()
            m2 = re.match(r"^(.+?)\s+swatch\b", vh_text, re.I)
            if m2:
                finish_name = m2.group(1).strip().title()

        if not finish_code and not finish_name:
            continue

        row = dict(base)
        if finish_name:
            row["Finish"] = finish_name
        row["Image URL"] = build_image_url(sku, finish_code) if finish_code else base.get("_main_image", "")
        # For size-variant products (beds): match variant name to size-specific dims
        if size_dims and finish_name:
            for sz_key, sz_dims in size_dims.items():
                if sz_key.lower() in finish_name.lower():
                    row["Dimensions"] = sz_dims
                    break
        variants.append(row)

    # ── Fallback: single row ──
    if not variants:
        row = dict(base)
        row["Image URL"] = base.get("_main_image", "")
        variants = [row]

    # Strip internal key
    for v in variants:
        v.pop("_main_image", None)

    return variants


# ── main ─────────────────────────────────────────────────────────────────────

async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    try:
        async with async_playwright() as pw:
            browser = await pw.chromium.launch(
                headless=HEADLESS,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )
            ctx = await browser.new_context(
                user_agent=UA,
                viewport={"width": 1280, "height": 900},
                locale="en-US",
            )
            await ctx.add_init_script(
                "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
            )
            page = await ctx.new_page()

            # Akamai bypass: establish a clean session via the homepage first
            print("[Ballard] Homepage warmup...")
            await page.goto(BASE_URL + "/", timeout=45_000, wait_until="domcontentloaded")
            await page.wait_for_timeout(2_000)

            cats_done = 0
            for cat in info["categories"]:
                if not cat["links"]:
                    continue
                if TEST_MODE and cats_done >= TEST_MAX_CAT:
                    break

                cat_name = CAT_NAME_FIX.get(cat["name"], cat["name"])
                writer.add_sheet(cat_name, cat["links"][0], studio_columns=cat["studio_columns"])

                seen_urls: set[str] = set()
                all_product_urls: list[str] = []
                for listing_url in cat["links"]:
                    try:
                        for u in await get_product_links(page, listing_url):
                            if u not in seen_urls:
                                seen_urls.add(u)
                                all_product_urls.append(u)
                    except Exception as exc:
                        print(f"  LISTING ERROR [{cat_name}] {listing_url}: {exc}")
                        await _reset_page(page)

                if TEST_MODE:
                    all_product_urls = all_product_urls[:TEST_MAX_PROD]

                print(f"[{cat_name}] {len(all_product_urls)} products to scrape")
                cats_done += 1

                global_idx = 1
                for url in all_product_urls:
                    try:
                        variant_rows = await scrape_product(page, url)
                        for variant in variant_rows:
                            if not variant.get("SKU"):
                                variant["SKU"] = generate_sku(VENDOR_NAME, cat_name, global_idx)
                            if not variant.get("Product Family Id") and variant.get("Product Name"):
                                variant["Product Family Id"] = extract_family_id(variant["Product Name"])
                            writer.write_row(variant, category_name=cat_name)
                            global_idx += 1
                    except Exception as exc:
                        print(f"  ERROR {url}: {exc}")
                        await _reset_page(page)
                    await async_polite_delay()

    finally:
        writer.save()
        print(f"[Ballard] Saved -> {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
