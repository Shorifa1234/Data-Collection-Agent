import asyncio, json, os, re, sys
from pathlib import Path
from urllib.parse import urlparse

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from bs4 import BeautifulSoup

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    clean_price, generate_sku, extract_family_id,
    parse_dimensions, safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Formations")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.formationsusa.com"
EMAIL    = "procurement@studiodesigner.com"
PASSWORD = "20Catalog26"

# Unicode fraction to decimal
_FRAC_MAP = [
    ("½", ".5"), ("¼", ".25"), ("¾", ".75"),
    ("⅓", ".333"), ("⅔", ".667"),
    ("⅛", ".125"), ("⅜", ".375"), ("⅝", ".625"), ("⅞", ".875"),
]

def _convert_fractions(s):
    for char, suffix in _FRAC_MAP:
        s = re.sub("(\\d)" + re.escape(char), r"\g<1>" + suffix, s)
        s = s.replace(char, "0" + suffix)
    return s


# Formations pages use U+201D (right double quotation mark) as inch mark.
_SMART_QUOTES = (chr(0x201C), chr(0x201D), chr(0x2018), chr(0x2019))

def _norm_quotes(s):
    for ch in _SMART_QUOTES:
        s = s.replace(ch, '"')
    return s


_FRAC_CLASS = r"[\d\xbd\xbc\xbe⅓⅔⅛⅜⅝⅞]"

_IS_DIM_LINE = re.compile(
    _FRAC_CLASS + r'+["]?\s*(?:dia|[dwhl])\b',
    re.IGNORECASE,
)
_DIM_TOKEN = re.compile(r'([\d.]+)["]?\s*(dia|[dwhl])\b', re.IGNORECASE)
_SEAT_H_RE = re.compile(r'seat\s+height\s*:?\s*(' + _FRAC_CLASS + r'+)["]?', re.IGNORECASE)
_ARM_H_RE  = re.compile(r'arm\s+height\s*:?\s*(' + _FRAC_CLASS + r'+)["]?',  re.IGNORECASE)
_WEIGHT_RE = re.compile(r'^([\d.]+)\s*lbs?\.?\s*$', re.IGNORECASE)
_PRICE_RE  = re.compile(r'^\$[\d,]+', re.IGNORECASE)

_LABEL_MAP = {"d": "Depth", "w": "Width", "h": "Height", "l": "Length", "dia": "Diameter"}


def _parse_dim_line(raw):
    s = _norm_quotes(_convert_fractions(raw))
    result = {}
    for m in _DIM_TOKEN.finditer(s):
        val = safe_float(m.group(1))
        col = _LABEL_MAP.get(m.group(2).lower())
        if col and val is not None:
            result[col] = val
    return result


def _parse_single_dim(raw):
    s = _norm_quotes(_convert_fractions(raw.strip()))
    return safe_float(s)


async def _login(page):
    await page.goto(f"{BASE_URL}/login", timeout=45_000, wait_until="domcontentloaded")
    await page.fill("input[name=email]",    EMAIL)
    await page.fill("input[name=password]", PASSWORD)
    await page.click("input[type=submit]")
    await page.wait_for_timeout(3000)


def _extract_links_from_soup(soup):
    seen, links = set(), []
    for a in soup.select(".product .image a, .product h4 a"):
        href = a.get("href", "")
        if href and "/catalog/" in href and href not in seen:
            seen.add(href)
            links.append(href)
    return links


async def get_product_links(page, listing_url):
    await page.goto(listing_url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)
    html = await page.content()
    soup = BeautifulSoup(html, "html.parser")

    if soup.select_one("#product") and not soup.select_one("#products"):
        return [listing_url]

    links = _extract_links_from_soup(soup)

    cat_id = None
    for script in soup.find_all("script"):
        text = script.string or ""
        m = re.search(r"category_id['\"]?\s*[,:]\s*['\"]?(\d+)", text)
        if m:
            cat_id = m.group(1)
            break

    if not cat_id:
        print(f"    [warn] no category_id for {listing_url}")
        return links

    cat_link = urlparse(listing_url).path.lstrip("/")
    current_page = 0

    while True:
        try:
            raw = await page.evaluate(f"""
                async () => {{
                    const fd = new FormData();
                    fd.append('category_id',  '{cat_id}');
                    fd.append('current_page', '{current_page}');
                    fd.append('category_link', '{cat_link}');
                    const r = await fetch('/ajax/get_category_products_for_pagination/', {{
                        method: 'POST', body: fd
                    }});
                    return await r.text();
                }}
            """)
            data = json.loads(raw)
        except Exception as e:
            print(f"    AJAX error page {current_page}: {e}")
            break

        html_chunk = data.get("html") or ""
        if not html_chunk or html_chunk.strip() in ("", "null"):
            break

        new_links = _extract_links_from_soup(BeautifulSoup(html_chunk, "html.parser"))
        if not new_links:
            break
        for lnk in new_links:
            if lnk not in set(links):
                links.append(lnk)
        current_page += 1
        await asyncio.sleep(0.6)

    return list(dict.fromkeys(links))


async def scrape_product(page, url):
    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(1500)
    html = await page.content()
    soup = BeautifulSoup(html, "html.parser")

    prod = soup.select_one("#product")
    if not prod:
        print(f"    [warn] no #product on {url}")
        return []

    data = {"Source": url, "Manufacturer": VENDOR_NAME}

    sku_el = prod.select_one(".sku, span.sku")
    data["SKU"] = clean_text(sku_el.get_text()) if sku_el else ""

    h1 = prod.select_one("h1")
    raw_name = clean_text(h1.get_text()) if h1 else ""
    data["Product Name"] = raw_name.title() if raw_name else ""

    zoom_img = soup.select_one("img.zoom, #product img")
    if zoom_img:
        src = zoom_img.get("src", "")
        src = src.replace("/medium_", "/")
        data["Image URL"] = src

    tearsheet = prod.select_one("a.tearsheet[href]")
    if tearsheet:
        data["Tearsheet Link"] = tearsheet.get("href", "")

    desc_lines = []
    dims_found = False

    for p_tag in prod.select(".product_info p"):
        raw_html = str(p_tag)
        chunk = re.sub(r"<br\s*/?>", "\n", raw_html, flags=re.IGNORECASE)
        lines = [clean_text(BeautifulSoup(l, "html.parser").get_text()) for l in chunk.split("\n")]
        lines = [l for l in lines if l]

        for line in lines:
            line_n  = _norm_quotes(line)
            line_cf = _convert_fractions(line_n)

            if _IS_DIM_LINE.search(line_n) and not dims_found:
                dim_dict = _parse_dim_line(line_n)
                if dim_dict:
                    data.update(dim_dict)
                    parts = []
                    for lbl, key in (("W", "Width"), ("D", "Depth"), ("H", "Height"),
                                     ("Dia", "Diameter"), ("L", "Length")):
                        if key in dim_dict:
                            parts.append(f"{dim_dict[key]}{lbl}")
                    if parts:
                        data["Dimensions"] = " x ".join(parts)
                    dims_found = True
                    continue

            m = _SEAT_H_RE.search(line_n)
            if m:
                v = _parse_single_dim(m.group(1))
                if v is not None:
                    data["Seat Height"] = v
                continue

            m = _ARM_H_RE.search(line_n)
            if m:
                v = _parse_single_dim(m.group(1))
                if v is not None:
                    data["Arm Height"] = v
                continue

            m = _WEIGHT_RE.search(line_cf)
            if m:
                v = safe_float(m.group(1))
                if v is not None:
                    data["Weight"] = v
                continue

            if _PRICE_RE.search(line_n):
                continue

            if line_n:
                desc_lines.append(line_n)

    if desc_lines:
        data["Description"] = clean_text(" ".join(desc_lines))

    data["Product Family Id"] = extract_family_id(data.get("Product Name", ""))
    return [data]


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    cats = info["categories"]
    if TEST_MODE:
        cats = cats[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        print("[Formations] Logging in ...")
        await _login(page)
        print("[Formations] Logged in OK")

        for cat in cats:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls = set()
            all_product_urls = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"  [{cat['name']}] {len(all_product_urls)} products")

            global_idx = 1
            for url in all_product_urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"    ERROR {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
