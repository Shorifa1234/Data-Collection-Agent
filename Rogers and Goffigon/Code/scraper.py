import asyncio, json, os, re, sys
import requests
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text, clean_price, generate_sku, extract_family_id,
    sentence_case, async_polite_delay,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Rogers and Goffigon")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.rogersandgoffigon.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


# ── HTML / JSON helpers ──────────────────────────────────────────────────────

def extract_product_json(html: str) -> dict | None:
    """
    Extract the main product JSON object from the Next.js RSC payload.
    In the HTML the JSON is JS-string-encoded: keys/values use \\" instead of ".
    We find the unique \"product\":{\"brand_name\": marker then count raw braces
    (which are not escaped) to locate the object boundaries.
    """
    marker = '\\"product\\":{\\"brand_name\\":'
    idx = html.find(marker)
    if idx == -1:
        return None

    open_brace = idx + html[idx:].index('{')
    depth = 0
    i = open_brace
    bs = chr(92)  # backslash
    while i < len(html):
        c = html[i]
        if c == bs:
            i += 2
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                raw = html[open_brace:i + 1]
                # Unescape JS-string encoding: \\" -> "  and  \\\\ -> \\
                raw = raw.replace('\\"', '"').replace('\\\\', '\\')
                try:
                    return json.loads(raw)
                except json.JSONDecodeError:
                    return None
        i += 1
    return None


def get_product_links(html: str, collection: str) -> list[str]:
    """Return unique product href paths for the given collection slug."""
    pattern = rf'href="(/collection/{re.escape(collection)}/[^"]+)"'
    links = re.findall(pattern, html)
    return list(dict.fromkeys(links))  # preserve order, deduplicate


# ── Field parsers ─────────────────────────────────────────────────────────────

def _parse_repeat(repeat_str: str | None) -> tuple[str | None, str | None]:
    """Split 'H-5/8" V-5/8"' into (horizontal, vertical) without inch marks."""
    if not repeat_str:
        return None, None
    s = repeat_str.replace('"', '').strip()
    m = re.match(r'H-(\S+)\s+V-(\S+)', s, re.IGNORECASE)
    if m:
        return m.group(1), m.group(2)
    return None, None


def _clean_width(val: str | None) -> str | None:
    """'125\" APRX' → '125'"""
    if not val:
        return None
    cleaned = val.replace('"', '')
    cleaned = re.sub(r'\s*(APRX|APPROXIMATE)\s*', '', cleaned, flags=re.IGNORECASE).strip()
    return cleaned if cleaned else None


def _to_title(val: str | None) -> str | None:
    if not val:
        return None
    return val.strip().title() if val.strip() else None


# ── Row builder ───────────────────────────────────────────────────────────────

def build_row(prod: dict) -> dict:
    item_num    = prod.get("item_number") or ""
    style_name  = _to_title(prod.get("style_name") or prod.get("pdf_style_name") or "")
    color_name  = _to_title(prod.get("color_name") or "")
    collection  = prod.get("collection") or ""
    image_url   = prod.get("image_url") or ""

    row: dict = {
        "Manufacturer":       VENDOR_NAME,
        "Source":             f"{BASE_URL}/collection/{collection}/{item_num}",
        "Image URL":          image_url,
        "Product Name":       f"{style_name} {color_name}".strip() if style_name else item_num,
        "Product Family Id":  style_name or extract_family_id(item_num),
        "SKU":                item_num,
        "Color":              color_name,
    }

    # Sub-brand / mill vendor
    brand = prod.get("brand_name") or prod.get("vendor_name")
    if brand:
        row["Brand"] = _to_title(brand)

    # Price
    price = prod.get("wholesale_price")
    if price:
        row["Price"] = clean_price(str(price))

    # Contents / material composition
    content = prod.get("style_content")
    if content:
        row["Contents"] = clean_text(content)

    # Repeat — full string + parsed H/V
    repeat = prod.get("style_repeat")
    if repeat:
        row["Repeat"] = repeat.strip()
        h_rep, v_rep = _parse_repeat(repeat)
        if h_rep:
            row["Horizontal Repeat"] = h_rep
        if v_rep:
            row["Vertical Repeat"] = v_rep

    # Width
    width = _clean_width(prod.get("style_width"))
    if width:
        row["Width"] = width

    # Weight
    wt = prod.get("weight")
    if wt is not None:
        row["Weight"] = str(wt)

    # Lead time (in days)
    lt = prod.get("lead_time")
    if lt is not None:
        row["Lead Time"] = f"{lt} days"

    # Design attributes (WIDE WIDTH, SOLID, TEXTURE, SHEER etc.)
    design = prod.get("product_design_name")
    if design:
        row["Description"] = clean_text(design)

    # Finish
    finish = prod.get("style_finish")
    if finish:
        row["Finish"] = clean_text(finish)

    # Style comment (extra notes, e.g. hide size for leather)
    comment = prod.get("style_comment")
    if comment:
        row["Notes"] = clean_text(comment)

    # Testing certifications
    tests = prod.get("style_tests")
    if tests:
        row["Testing"] = clean_text(tests)

    # Care / cleaning note
    cleaning = prod.get("cleaning_note")
    if cleaning:
        row["Care Instructions"] = clean_text(cleaning)

    # CFA availability
    cfa = prod.get("cfa_offered")
    if cfa is not None:
        row["CFA Available"] = "Yes" if cfa else "No"

    # Selling unit
    unit = prod.get("selling_unit_name")
    if unit and unit.lower() not in ("unknown",):
        row["Selling Unit"] = unit

    # Color codes
    primary = prod.get("primary_color_code")
    if primary:
        row["Primary Color"] = _to_title(primary)
    secondary = prod.get("secondary_color_code")
    if secondary:
        row["Secondary Color"] = _to_title(secondary)

    # Misc / usage
    misc = prod.get("misc_name")
    if misc:
        row["Misc"] = misc
    usage = prod.get("usage_name")
    if usage:
        row["Usage"] = _to_title(usage)

    # Collection name (named collection, if any)
    coll_name = prod.get("collection_name")
    if coll_name:
        row["Collection"] = _to_title(coll_name)

    # Website description
    web_desc = prod.get("style_website_description") or prod.get("item_website_description")
    if web_desc:
        row["Web Description"] = clean_text(web_desc)

    # Remove None / empty
    return {k: v for k, v in row.items() if v is not None and v != ""}


# ── HTTP helpers (run in thread pool) ────────────────────────────────────────

def _fetch(session: requests.Session, url: str) -> str:
    r = session.get(url, headers=HEADERS, timeout=25)
    r.raise_for_status()
    return r.text


def _collect_listing_links(session: requests.Session, base_url: str,
                            collection: str) -> list[str]:
    """Paginate through listing and return all unique product hrefs for collection."""
    seen: set[str] = set()
    result: list[str] = []
    page = 1
    while True:
        html = _fetch(session, f"{base_url}?page={page}")
        links = get_product_links(html, collection)
        new = [l for l in links if l not in seen]
        if not new:
            break
        seen.update(new)
        result.extend(new)
        page += 1
    return result


# ── Category scraper ──────────────────────────────────────────────────────────

async def scrape_category(session: requests.Session, cat: dict,
                          writer: ExcelWriter) -> None:
    if TEST_MODE:
        print(f"  [TEST: max {TEST_MAX_PRODUCTS} products per category]")

    # Collect product URLs from ALL listing links (deduplicated)
    seen_urls: set[str] = set()
    all_product_paths: list[str] = []

    for listing_url in cat["links"]:
        # Derive the collection slug from the URL (last path segment)
        collection_slug = listing_url.rstrip("/").split("/")[-1]
        paths = await asyncio.to_thread(
            _collect_listing_links, session, listing_url, collection_slug
        )
        for p in paths:
            if p not in seen_urls:
                seen_urls.add(p)
                all_product_paths.append(p)
            if TEST_MODE and len(all_product_paths) >= TEST_MAX_PRODUCTS:
                break
        if TEST_MODE and len(all_product_paths) >= TEST_MAX_PRODUCTS:
            break

    label = " [TEST]" if TEST_MODE else ""
    print(f"  {cat['name']}: {len(all_product_paths)} products to scrape{label}")

    global_idx = 1
    for path in all_product_paths:
        url = BASE_URL + path
        try:
            html = await asyncio.to_thread(_fetch, session, url)
            prod = extract_product_json(html)
            if not prod:
                print(f"  [WARN] No product JSON at {path}")
                continue

            row = build_row(prod)
            if not row.get("SKU"):
                row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
            writer.write_row(row, category_name=cat["name"])
            global_idx += 1
        except Exception as e:
            print(f"  [ERROR] {path}: {e}")

        await asyncio.sleep(0.3)

    print(f"  {cat['name']}: wrote {global_idx - 1} rows")


# ── Main ──────────────────────────────────────────────────────────────────────

async def main() -> None:
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])
    session = requests.Session()

    categories = info["categories"]
    if TEST_MODE and TEST_MAX_CATEGORIES < len(categories):
        categories = categories[:TEST_MAX_CATEGORIES]

    for cat in categories:
        if not cat["links"]:
            continue
        writer.add_sheet(cat["name"], cat["links"][0],
                         studio_columns=cat["studio_columns"])
        await scrape_category(session, cat, writer)

    session.close()
    writer.save()
    print(f"\nSaved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
