import asyncio, json, os, sys, re, time
import requests
from pathlib import Path
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text,
    generate_sku, extract_family_id,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Loloi")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.loloirugs.com"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
})


def _parse_rug_size(size_str: str) -> dict:
    """
    Parse rug size string into Width, Length (decimal feet), and optional Shape.

    Handles:
      "2'-3\" x 3'-9\""    → Width=2.25,  Length=3.75
      "6'-0\" x 6'-0\" Round" → Width=6.0, Length=6.0, Shape=Round
      "7' x 9' Oval"        → Width=7.0,  Length=9.0,  Shape=Oval
      "18\" x 18\" Sample"  → Width=1.5,  Length=1.5
      "10'"                 → (just feet)
    Returns a dict with str values (numeric strings, no units).
    """
    result: dict[str, str] = {}

    # Detect shape suffix
    shape_match = re.search(r'\b(Round|Oval|Square)\b', size_str, re.I)
    if shape_match:
        result["Shape"] = shape_match.group(1).title()

    # Strip non-dimensional words (Sample, Round, Oval, Square)
    clean = re.sub(r'\b(Sample|Round|Oval|Square)\b', '', size_str, flags=re.I).strip()

    # Split on × or x
    parts = re.split(r'\s*[xX×]\s*', clean)

    def to_inches(token: str) -> str | None:
        """Convert any size token to inches (integer or .5 precision)."""
        token = token.strip()
        if not token:
            return None
        # Feet-inches: 2'-3" or 2'3" or 2'-0"
        m = re.match(r"(\d+)['′][-\s]?(\d+)?[\"″]?$", token)
        if m:
            feet   = int(m.group(1))
            inches = int(m.group(2)) if m.group(2) else 0
            total  = feet * 12 + inches
            return str(total)
        # Just feet: 10'
        m = re.match(r"(\d+)['′]$", token)
        if m:
            return str(int(m.group(1)) * 12)
        # Just inches: 18"
        m = re.match(r"(\d+(?:\.\d+)?)[\"″]$", token)
        if m:
            return str(round(float(m.group(1)), 2)).rstrip("0").rstrip(".")
        # Plain number fallback: treat as inches
        m = re.match(r"(\d+(?:\.\d+)?)$", token)
        if m:
            return str(round(float(m.group(1)), 2)).rstrip("0").rstrip(".")
        return None

    dims = [to_inches(p) for p in parts if p.strip()]
    dims = [d for d in dims if d]

    if len(dims) >= 2:
        result["Width"]  = dims[0]
        result["Length"] = dims[1]
    elif len(dims) == 1:
        # Single dimension = diameter (round rug listed with one value)
        result["Diameter"] = dims[0]

    return result


def _parse_tags(tags_raw) -> dict[str, list[str]]:
    if isinstance(tags_raw, list):
        tags_str = ", ".join(tags_raw)
    else:
        tags_str = str(tags_raw)
    result: dict[str, list[str]] = {}
    for tag in tags_str.split(", "):
        if ":" in tag:
            k, v = tag.split(":", 1)
            result.setdefault(k.strip(), []).append(v.strip())
    return result


def _fetch_product(handle: str) -> list[dict]:
    """Fetch JSON + HTML for one product, return one dict per size variant."""
    product_url = f"{BASE_URL}/products/{handle}"

    # --- JSON ---
    try:
        r = SESSION.get(f"{product_url}.json", timeout=25)
        if r.status_code != 200:
            return []
        p = r.json().get("product", {})
    except Exception as e:
        print(f"  JSON error {handle}: {e}")
        return []

    # --- HTML ---
    try:
        r_html = SESSION.get(product_url, timeout=25)
        soup = BeautifulSoup(r_html.text, "html.parser")
    except Exception:
        soup = BeautifulSoup("", "html.parser")

    # ── Tags ──
    tag_dict = _parse_tags(p.get("tags", ""))
    collection    = ", ".join(tag_dict.get("Collection", []))
    colors        = " / ".join(tag_dict.get("Color", []))
    style         = ", ".join(tag_dict.get("Style", []))
    collaboration = ", ".join(tag_dict.get("Collaboration", []))

    # Richer collaboration text from HTML
    collab_el = soup.find(class_="collaboration-border")
    if collab_el:
        raw = collab_el.get_text(strip=True)
        # Normalize garbled × characters
        raw = re.sub(r"[נ××]", "×", raw)
        if raw:
            collaboration = raw

    # Designer = part before "× Loloi"
    designer = ""
    if "×" in collaboration:
        parts = re.split(r"\s*×\s*", collaboration)
        non_loloi = [pt.strip() for pt in parts if "loloi" not in pt.lower()]
        if non_loloi:
            designer = non_loloi[0]

    # ── DL specs ──
    spec_data:  dict[str, str] = {}
    tech_specs: dict[str, str] = {}
    for dl in soup.find_all("dl"):
        if "spacer-2--pl" not in (dl.get("class") or []):
            continue
        for div in dl.find_all("div"):
            dt = div.find("dt")
            dd = div.find("dd")
            if not dt or not dd:
                continue
            key  = dt.get_text(strip=True)
            vals = [li.get_text(strip=True) for li in dd.find_all("li")]
            if not vals:
                vals = [dd.get_text(strip=True)]

            if key == "Technical Specs":
                for item in vals:
                    if ":" in item:
                        tk, tv = item.split(":", 1)
                        # Strip inch/foot marks from numeric values
                        tv_clean = tv.strip().rstrip('"').rstrip("'").rstrip('"').strip()
                        tech_specs[tk.strip()] = tv_clean
            else:
                spec_data[key] = ", ".join(v for v in vals if v)

    # ── MAP Policy ──
    map_policy = ""
    map_dl = soup.find("dl", class_=re.compile(r"spacer-4--mb"))
    if map_dl:
        dd_el = map_dl.find("dd")
        if dd_el:
            strings = [t.strip() for t in dd_el.strings if t.strip() and len(t.strip()) > 3]
            map_policy = strings[0] if strings else ""

    # ── Care Instructions ──
    care_instructions = ""
    for acc in soup.find_all("div", class_="accordion__item"):
        txt = acc.get_text(strip=True)
        if txt.lower().startswith("care instructions"):
            # Strip the heading
            body = txt[len("Care Instructions"):].strip()
            care_instructions = clean_text(body)
            break

    # ── Product Name / Family ID ──
    product_name = p.get("title", "")
    m = re.match(r"^([A-Z]+-\d+[A-Z]*)", product_name.upper())
    family_id = m.group(1) if m else extract_family_id(product_name)

    # ── Description ──
    body_html = p.get("body_html", "")
    description = clean_text(BeautifulSoup(body_html, "html.parser").get_text()) if body_html else ""

    # ── Image URL ──
    images    = p.get("images", [])
    image_url = images[0]["src"] if images else ""

    # ── Tearsheet ──
    tearsheet = f"{BASE_URL}/products/{handle}?view=spec-sheet"

    # ── Base shared dict ──
    base: dict = {
        "Manufacturer":      VENDOR_NAME,
        "Source URL":        product_url,
        "Image URL":         image_url,
        "Product Name":      product_name,
        "Product Family Id": family_id,
        "Description":       description,
        "Collection":        collection,
        "Color":             colors,
        "Style":             style,
        "Construction":      spec_data.get("Construction", ""),
        "Material":          spec_data.get("Material", ""),
        "Pile Height":       tech_specs.get("Pile Height", ""),
        "Backing":           tech_specs.get("Backing", ""),
        "Country of Origin": spec_data.get("Country of Origin", ""),
        "MAP Policy":        map_policy,
        "Tearsheet Link":    tearsheet,
    }
    if designer:
        base["Designer"] = designer
    if collaboration:
        base["Collaboration"] = collaboration
    if care_instructions:
        base["Care Instructions"] = care_instructions

    # Drop empty strings
    base = {k: v for k, v in base.items() if v not in (None, "")}

    # ── One row per size variant ──
    rows = []
    for variant in p.get("variants", []):
        row  = dict(base)
        row["SKU"]  = variant.get("sku", "")
        size = variant.get("title", "")
        if size and size.lower() not in ("default title", ""):
            row["Size"]       = size
            row["Dimensions"] = size
            # Parse Width / Length / Shape (no Height/Depth for rugs)
            dim_fields = _parse_rug_size(size)
            for df_key, df_val in dim_fields.items():
                if df_val:
                    row[df_key] = df_val
        grams = variant.get("grams", 0)
        if grams:
            row["Weight"] = round(grams / 453.592, 2)
        # Per-variant image (if assigned)
        fi = variant.get("featured_image")
        if fi and fi.get("src"):
            row["Image URL"] = fi["src"]
        rows.append(row)

    return rows if rows else [dict(base)]


async def get_product_links(page, listing_url: str) -> list[str]:
    """Return all product URLs by paginating the Shopify JSON API."""
    collection_base = listing_url.rstrip("/")
    urls: list[str] = []
    pg = 1
    while True:
        api_url = f"{collection_base}/products.json?limit=250&page={pg}"
        try:
            r = await asyncio.to_thread(SESSION.get, api_url, timeout=20)
            products = r.json().get("products", [])
        except Exception as e:
            print(f"  API error page {pg}: {e}")
            break
        if not products:
            break
        for prod in products:
            urls.append(f"{BASE_URL}/products/{prod['handle']}")
        pg += 1
        await asyncio.sleep(0.2)
    return urls


async def scrape_product(page, url: str) -> list[dict]:
    """Fetch and parse one product; return list of variant dicts."""
    handle = url.split("/products/")[-1].split("?")[0]
    rows = await asyncio.to_thread(_fetch_product, handle)
    return rows


async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    if TEST_MODE:
        print(f"[TEST: max {TEST_MAX_PRODUCTS} products per category]")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        cats = info["categories"]
        if TEST_MODE:
            cats = cats[:TEST_MAX_CATEGORIES]

        for cat in cats:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen_urls:        set[str]  = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"[{cat['name']}] {len(all_product_urls)} products to scrape")

            global_idx = 1
            for url in all_product_urls:
                try:
                    variant_rows = await scrape_product(page, url)
                    for variant in variant_rows:
                        if not variant.get("SKU"):
                            variant["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                        if not variant.get("Product Family Id") and variant.get("Product Name"):
                            variant["Product Family Id"] = extract_family_id(variant["Product Name"])
                        writer.write_row(variant, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  ERROR {url}: {e}")
                await asyncio.sleep(0.3)

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
