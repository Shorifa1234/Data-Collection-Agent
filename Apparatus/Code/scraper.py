import asyncio, json, os, sys, re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Apparatus")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

BASE_URL = "https://apparatusstudio.com"


def _to_abs(url: str) -> str:
    if not url:
        return ""
    if url.startswith("//"):
        return "https:" + url
    if url.startswith("/"):
        return BASE_URL + url
    return url


def _first_price(text: str, usd_only: bool = False) -> float | None:
    """Extract the first clean numeric price from text (prefers USD, falls back to GBP)."""
    # USD first
    usd = re.search(r'\$\s*([\d,]+(?:\.\d+)?)', text)
    if usd:
        return clean_price(usd.group(1))
    if usd_only:
        return None
    # GBP/EUR fallback
    other = re.search(r'[£€]\s*([\d,]+(?:\.\d+)?)', text)
    if other:
        return clean_price(other.group(1))
    # Bare number after "Price from :"
    bare = re.search(r'PRICE\s+FROM\s*:?\s*([\d,]+(?:\.\d+)?)', text, re.IGNORECASE)
    if bare:
        return clean_price(bare.group(1))
    return None


async def _section_raw(page, heading: str) -> str:
    """
    Click the h2 matching `heading` and return the raw multi-line text
    of its following siblings (up to the next h2). Newlines preserved.
    """
    h2s = await page.query_selector_all("h2")
    target = None
    for h2 in h2s:
        t = clean_text(await h2.inner_text())
        if heading.lower() in t.lower():
            target = h2
            break
    if not target:
        return ""
    try:
        await target.click(timeout=3000)
        await page.wait_for_timeout(700)
    except Exception:
        pass

    raw = await page.evaluate(
        """(el) => {
            let lines = [];
            let node = el.nextElementSibling;
            let steps = 0;
            while (node && steps < 40) {
                if (node.tagName && node.tagName.toLowerCase() === 'h2') break;
                // Collect all text nodes, preserving line breaks via innerText
                const t = (node.innerText || node.textContent || '').trim();
                if (t) lines.push(t);
                node = node.nextElementSibling;
                steps++;
            }
            return lines.join('\\n');
        }""",
        target
    )
    return raw.strip()


async def _product_images(page) -> list[str]:
    """CDN product images, logo/brand images excluded."""
    imgs = await page.query_selector_all(
        "img[src*='cdn/shop/files'], img[srcset*='cdn/shop/files']"
    )
    urls, seen = [], set()
    for img in imgs:
        src = (await img.get_attribute("src") or "").split("?")[0]
        if not src:
            continue
        lc = src.lower()
        if any(x in lc for x in ["logo", "branding", "icon", "badge", "web_logo"]):
            continue
        full = _to_abs(src)
        if full not in seen:
            seen.add(full)
            urls.append(full)
    return urls


def _parse_about(raw: str) -> tuple[str, str, str]:
    """
    Split 'About the Series' raw text into (description, collection, material).
    The section contains optional sub-headers REFERENCE and COMPOSITION.
    """
    description = collection = material = ""

    ref_idx  = re.search(r'\bREFERENCE\s*:', raw, re.IGNORECASE)
    comp_idx = re.search(r'\bCOMPOSITION\s*:', raw, re.IGNORECASE)

    desc_end = min(
        ref_idx.start()  if ref_idx  else len(raw),
        comp_idx.start() if comp_idx else len(raw),
    )
    description = clean_text(raw[:desc_end])

    if ref_idx:
        ref_start = ref_idx.end()
        ref_end = comp_idx.start() if (comp_idx and comp_idx.start() > ref_idx.start()) else len(raw)
        ref_text = raw[ref_start:ref_end]
        refs = [l.strip() for l in ref_text.splitlines() if l.strip()]
        collection = " / ".join(refs)

    if comp_idx:
        comp_text = raw[comp_idx.end():]
        # Remove lines that look like variant labels "WORD : WORD"
        lines = [l.strip() for l in comp_text.splitlines() if l.strip()]
        mat_lines = [l for l in lines if not re.match(r'^[A-Z\s]+\s*:\s*[A-Z\s]+$', l)]
        material = " / ".join(mat_lines)

    return description, collection, material


def _parse_finishes(raw: str) -> list[tuple[str, float | None]]:
    """
    Parse METAL FINISH section into list of (finish_name, price_adder_pct_or_None).
    Each non-blank line that is not a price line is a finish.
    """
    results = []
    for line in raw.splitlines():
        line = line.strip()
        if not line or re.match(r'^[£$€\d]', line):
            continue
        adder_m = re.search(r'ADD\s+(\d+)%', line, re.IGNORECASE)
        finish_name = re.sub(r'\s*ADD\s+\d+%.*', '', line, flags=re.IGNORECASE).strip()
        if not finish_name:
            continue
        adder = int(adder_m.group(1)) if adder_m else None
        results.append((sentence_case(finish_name), adder))
    return results


def _clean_tech_field(raw: str, pattern: str) -> str:
    """Extract value of a labeled field, first line only."""
    m = re.search(pattern, raw, re.IGNORECASE)
    if not m:
        return ""
    val = m.group(1).strip()
    return clean_text(val.splitlines()[0]) if val else ""


def _parse_dims_from_text(raw: str) -> dict:
    """Parse dimension and weight fields from a raw text block."""
    data: dict = {}
    # Diameter: "Ø 18 IN"
    dia_m = re.search(r'[Øø]\s*([\d.]+)', raw)
    if dia_m:
        data["Diameter"] = dia_m.group(1)
    # HEIGHT TO ORDER
    if re.search(r'HEIGHT\s+TO\s+ORDER', raw, re.IGNORECASE):
        data["Height"] = "To Order"
    # number-first "39 W" / label-first "W 39"
    # Exclusions: W/ (leather price text), LBS (weight abbreviation)
    _LABEL_EXCLUSION = {"W": r"(?![/\d])", "H": r"(?![/\d])", "D": r"(?![/\d])", "L": r"(?![/\dBS])"}
    for label, field in [("H", "Height"), ("W", "Width"), ("D", "Depth"), ("L", "Length")]:
        if field in data:
            continue
        excl = _LABEL_EXCLUSION[label]
        m = re.search(rf'([\d.]+)\s+{label}{excl}', raw, re.IGNORECASE)
        if not m:
            m = re.search(rf'(?<![/\d]){label}\s+([\d.]+)', raw, re.IGNORECASE)
        if m:
            data[field] = m.group(1)
    # Weight
    wt_m = re.search(r'(?:APPROX\.?\s*)?([\d.]+)\s*LBS?', raw, re.IGNORECASE)
    if wt_m:
        data["Weight"] = safe_float(wt_m.group(1))
    # Build Dimensions string
    parts = []
    for field in ["Height", "Width", "Depth", "Length", "Diameter"]:
        v = data.get(field)
        if v and v != "To Order":
            parts.append(f"{field[0]} {v}")
    if parts:
        data["Dimensions"] = " x ".join(parts)
    return data


async def _get_size_variants(page, product_name: str) -> list[dict]:
    """
    Detect and expand sub-variant h2 sections like "ARROW : LARGE" / "ARROW : SMALL".
    Returns list of dicts with per-size fields (Size, Dimensions, Width/Length, Weight, Price).
    Returns [] if no size sub-sections exist.
    """
    SKIP_KEYWORDS = {
        "about", "composition", "metal finish", "leather", "suede",
        "technical", "lead time", "reference", "installation",
        "dimensions", "view tear", "inquire",
    }
    # Base name = first segment before " : " in product name (e.g. "ARROW" from "ARROW")
    base_name = product_name.split(" : ")[0].strip().upper()

    h2s = await page.query_selector_all("h2")
    size_h2s = []
    for h2 in h2s:
        label = clean_text(await h2.inner_text()).strip()
        upper = label.upper()
        if not upper.startswith(base_name):
            continue
        if " : " not in upper:
            continue
        if any(kw in upper.lower() for kw in SKIP_KEYWORDS):
            continue
        # Skip if the h2 label is the product name itself (not a sub-variant)
        if upper == product_name.upper():
            continue
        size_h2s.append((label, h2))

    if not size_h2s:
        return []

    results = []
    for size_label, h2_el in size_h2s:
        try:
            await h2_el.click(timeout=3000)
            await page.wait_for_timeout(800)
        except Exception:
            pass

        raw = await page.evaluate(
            """(el) => {
                let lines = [];
                let node = el.nextElementSibling;
                let steps = 0;
                while (node && steps < 30) {
                    if (node.tagName && node.tagName.toLowerCase() === 'h2') break;
                    const t = (node.innerText || node.textContent || '').trim();
                    if (t) lines.push(t);
                    node = node.nextElementSibling;
                    steps++;
                }
                return lines.join('\\n');
            }""",
            h2_el
        )

        size_data = _parse_dims_from_text(raw)
        size_data["Size"] = sentence_case(size_label)

        # Price specific to this size — prefer USD; GBP fallback accepted
        p = _first_price(raw)
        if p:
            size_data["_size_price"] = p  # store separately, applied later

        results.append(size_data)

    return results


async def scrape_product(page, url: str, vendor_name: str) -> list[dict]:
    await page.goto(url, timeout=60_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2500)

    base: dict = {"Source URL": url, "Manufacturer": vendor_name}

    # ── Product Name ──────────────────────────────────────────────
    h1 = await page.query_selector("h1")
    if h1:
        base["Product Name"] = clean_text(await h1.inner_text())
    if base.get("Product Name"):
        base["Product Family Id"] = extract_family_id(base["Product Name"])

    # ── Price (try CSS selector first, then body text) ────────────
    price_text = ""
    for sel in ["[class*='price']", ".price", "[data-price]", ".product__price"]:
        el = await page.query_selector(sel)
        if el:
            price_text = await el.inner_text()
            if price_text.strip():
                break
    if not price_text:
        body_snip = (await page.inner_text("body"))[:3000]
        price_text = body_snip

    p = _first_price(price_text)
    if p:
        base["Price"] = p

    # ── Full body text (for dimensions, weight) ───────────────────
    body_text = await page.inner_text("body")

    # ── Images ───────────────────────────────────────────────────
    imgs = await _product_images(page)
    if imgs:
        base["Image URL"] = imgs[0]

    # ── About the Series → Description, Collection, Material ─────
    about_raw = await _section_raw(page, "about the series")
    if about_raw:
        desc, coll, mat = _parse_about(about_raw)
        if desc:
            base["Description"] = desc
        if coll:
            base["Collection"] = coll
        if mat:
            base["Material"] = mat

    # ── Dimensions + Weight ──────────────────────────────────────
    dim_raw = await _section_raw(page, "dimensions")
    if not dim_raw:
        m = re.search(r'([\d.]+\s*[HWDhwd]\s*[\d.]+)', body_text)
        if m:
            dim_raw = m.group(0)
    if dim_raw:
        base.update(_parse_dims_from_text(dim_raw))
    # Weight from body if not already in dim section
    if not base.get("Weight"):
        wt_m = re.search(r'(?:APPROX\.?\s*)?([\d.]+)\s*LBS?', body_text, re.IGNORECASE)
        if wt_m:
            base["Weight"] = safe_float(wt_m.group(1))

    # ── Technical Specs ──────────────────────────────────────────
    tech_raw = await _section_raw(page, "technical spec")
    if tech_raw:
        base["Specifications"] = clean_text(tech_raw)

        v = _clean_tech_field(tech_raw, r'VOLTAGE\s*:?\s*([^\n]+)')
        if v:
            base["Voltage"] = v

        # Lamping: first LED/wattage line
        lamp_m = re.search(r'((?:INCLUDES?\s+)?[\d]+\s*[Xx×]\s*[\d]+W[^\n]*)', tech_raw)
        if not lamp_m:
            lamp_m = re.search(r'(LAMPING\s*:\s*[^\n]+)', tech_raw, re.IGNORECASE)
        if lamp_m:
            base["Lamping"] = clean_text(lamp_m.group(1))

        lumen_m = re.search(r'([\d,]+)\s*LUMENS?', tech_raw, re.IGNORECASE)
        if lumen_m:
            base["Lumens"] = lumen_m.group(1).replace(",", "")

        # Handles: "2700K", "2000-3000K WARM DIM", "2000 - 2800K"
        cct_m = re.search(r'([\d]+\s*[-–]\s*[\d]+K|[\d]+K)', tech_raw)
        if cct_m:
            base["Color Temperature"] = cct_m.group(1).strip()

        life_m = re.search(r'([\d,]+)\s*HOURS?', tech_raw, re.IGNORECASE)
        if life_m:
            base["Driver Life"] = life_m.group(1).replace(",", "") + " hours"

    # ── Lead Time ────────────────────────────────────────────────
    lead_raw = await _section_raw(page, "lead time")
    if lead_raw and "VIEW TEAR" not in lead_raw.upper():
        base["Lead Time"] = clean_text(lead_raw)

    # ── Tear Sheet ───────────────────────────────────────────────
    tear_el = await page.query_selector("a[href*='.pdf']")
    if tear_el:
        href = await tear_el.get_attribute("href") or ""
        if href:
            base["Tearsheet Link"] = _to_abs(href)

    # ── Finish — all options joined, one row per product ─────────
    finish_raw = await _section_raw(page, "metal finish")
    finishes = _parse_finishes(finish_raw) if finish_raw else []
    if finishes:
        base["Finish"] = ", ".join(name for name, _ in finishes)

    # ── Size sub-variants (e.g. ARROW : LARGE / ARROW : SMALL) ──
    # One row per size sub-section; no finish multiplication
    size_variants = await _get_size_variants(page, base.get("Product Name", ""))

    if size_variants:
        rows = []
        for sv in size_variants:
            row = dict(base)
            size_price = sv.pop("_size_price", None)
            row.update(sv)
            if size_price:
                row["Price"] = size_price
            rows.append(row)
        return rows

    return [base]


async def get_product_links(page, listing_url: str) -> list[str]:
    await page.goto(listing_url, timeout=60_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(3000)

    seen: set[str] = set()
    links: list[str] = []
    no_new = 0

    while no_new < 3:
        prev_len = len(links)
        for a in await page.query_selector_all("a[href*='/products/']"):
            href = await a.get_attribute("href") or ""
            if "/products/" in href and "/collections/" not in href and "/pages/" not in href:
                full = _to_abs(href.split("?")[0])
                if full not in seen:
                    seen.add(full)
                    links.append(full)

        no_new = 0 if len(links) > prev_len else no_new + 1

        load_more = await page.query_selector(
            "a.js-load-more, button.load-more, [data-load-more], "
            "a[class*='load-more'], button[class*='load-more']"
        )
        if load_more:
            try:
                await load_more.click(timeout=3000)
                await page.wait_for_timeout(3000)
                no_new = 0
                continue
            except Exception:
                pass

        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await page.wait_for_timeout(2500)

    return links


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    test_mode  = os.environ.get("TEST_MODE", "false").lower() == "true"
    test_limit = int(os.environ.get("TEST_LIMIT", "5"))

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in info["categories"]:
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
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if test_mode:
                all_product_urls = all_product_urls[:test_limit]

            global_idx = 1
            for url in all_product_urls:
                try:
                    rows = await scrape_product(page, url, info["vendor_name"])
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  [ERROR] {url}: {e}", file=sys.stderr)
                await async_polite_delay()

    writer.save()
    print(f"[Done] Saved -> {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
