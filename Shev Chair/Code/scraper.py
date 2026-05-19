import asyncio, json, os, re, sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, clean_price,
    generate_sku, extract_family_id, parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Shev Chair")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://shevchair.com"


def _parse_summary(html_text: str) -> dict:
    """Parse all product fields from the page HTML."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_text, "html.parser")
    data: dict = {}

    summary = soup.find(class_="summary")
    if not summary:
        return data

    # Name
    h1 = summary.find("h1")
    if h1:
        data["Product Name"] = clean_text(h1.get_text())

    # SKU / Model
    sku_el = summary.find("span", class_="sku")
    if sku_el:
        data["SKU"] = clean_text(sku_el.get_text())

    # Price (WooCommerce — contract items usually $0)
    price_el = summary.find(class_="woocommerce-Price-amount")
    if price_el:
        p = clean_price(price_el.get_text(strip=True))
        if p and p > 0:
            data["Price"] = p

    # Spec Sheet / Flyer links
    for a in summary.find_all("a", href=True):
        href = a["href"]
        label = a.get_text(strip=True).lower()
        if ".pdf" in href.lower():
            if "spec" in href.lower() or "spec" in label:
                data["Spec Sheet"] = href
            elif "flyer" in href.lower() or "flyer" in label:
                data["Flyer"] = href
            elif "Spec Sheet" not in data:
                data["Spec Sheet"] = href

    # WooCommerce product attributes table — check summary AND tab-additional_information
    # (Fabric products: Supplier, Grade, Finish, Color)
    def _find_attr_table(root):
        return root.find(
            "table",
            class_=lambda c: c and any(
                kw in " ".join(c)
                for kw in ("shop_attributes", "woocommerce-product-attributes")
            )
        )

    attr_table = _find_attr_table(summary)
    if not attr_table:
        # Also try the additional information tab panel
        tab_panel = soup.find(id="tab-additional_information")
        if tab_panel:
            attr_table = _find_attr_table(tab_panel)

    if attr_table:
        for row in attr_table.find_all("tr"):
            th = row.find("th")
            td = row.find("td")
            if th and td:
                key = clean_text(th.get_text())
                val = clean_text(td.get_text())
                if key and val and len(key) < 60:
                    data[key] = val
    else:
        # Fallback only for fabric-type pages (no Standard Features bullets)
        has_features = bool(summary.find(lambda t: t.name in ("h2","h3","h4")
                                         and "standard features" in t.get_text().lower()))
        if not has_features:
            _parse_fabric_attrs_from_text(summary.get_text(separator="|", strip=True), data)

    # Description: long paragraphs before Standard Features section
    desc_parts = []
    for el in summary.descendants:
        tag = getattr(el, "name", None)
        if tag in ("h2", "h3", "h4"):
            txt = clean_text(el.get_text())
            if any(kw in txt.lower() for kw in ("standard features", "options", "dimensions")):
                break
            continue
        if tag == "p":
            txt = clean_text(el.get_text())
            if not txt or "Model #" in txt or "Download" in txt:
                continue
            if len(txt) > 40:
                desc_parts.append(txt)
    if desc_parts:
        data["Description"] = " ".join(desc_parts[:5])

    # Standard Features and Options — bullet points
    features, options = [], []
    current_section = None
    for el in summary.find_all(["h2", "h3", "h4", "li"]):
        tag = el.name
        txt = clean_text(el.get_text())
        if not txt:
            continue
        if tag in ("h2", "h3", "h4"):
            tl = txt.lower()
            if "standard features" in tl:
                current_section = "features"
            elif "options" in tl:
                current_section = "options"
            elif any(kw in tl for kw in ("dimensions", "categories", "you may")):
                current_section = None
            continue
        if tag == "li":
            if current_section == "features":
                features.append(txt)
                _extract_feature_fields(txt, data)
            elif current_section == "options":
                options.append(txt)

    if features:
        data["Standard Features"] = "; ".join(features)
    if options:
        data["Options"] = "; ".join(options)

    # Dimensions — parsed from plain text (the "table" is CSS-only, no <table> element).
    # Summary text contains: "Overall 19.75 23.5 37.5 Seat 18.25 18.5 19"
    # after the "Dimensions Width Depth Height" header.
    summary_text = summary.get_text(separator=" ", strip=True)
    _parse_dimensions_from_text(summary_text, data, product_name=data.get("Product Name", ""))

    return data


def _first_dim_val(raw: str) -> str:
    """From '60 or 72' (with optional inch marks) return '60'."""
    v = raw.replace('″', '').replace('”', '').replace('"', '').strip()
    m = re.match(r"[\d.]+", v)
    return m.group() if m else v


def _parse_dimensions_from_text(text: str, data: dict, product_name: str = ""):
    """Extract dimensions from the plain text of the summary.

    Handles two layouts:
      - Chair/stool: "Overall W D H" / "Seat W D H"
      - Table: labelled rows "Round W D H", "Rectangle W D H"
        → picks the row whose label matches the product name.

    Strips U+2033 double-prime inch marks before running regexes so that
    values like '60″ or 72″' are matched correctly.
    'W or W2' patterns → first value taken via _first_dim_val().
    """
    # Strip inch marks so they don't break digit regexes
    text_c = text.replace('″', '').replace('”', '').replace('"', '')

    m_section = re.search(
        r"Dimensions\s+Width\s+Depth\s+Height(.+?)(?:Options|Categor|$)",
        text_c, re.IGNORECASE | re.DOTALL,
    )
    if not m_section:
        return
    dim_text = m_section.group(1).strip()

    # Chair/stool: "Overall W D H"
    m_overall = re.search(
        r"Overall\s+([\d.]+(?:\s+or\s+[\d.]+)?)\s+([\d.]+(?:\s+or\s+[\d.]+)?)\s+([\d.]+)",
        dim_text, re.IGNORECASE,
    )

    # Table: labelled rows e.g. "Round 60 or 72  60 or 72  29"
    # Anchor: start-of-string or preceded by whitespace (handles single-space separators)
    row_pat = re.compile(
        r"(?:(?<=\s)|^)([A-Za-z][A-Za-z /\-]*?)\s+"
        r"([\d.]+(?:\s+or\s+[\d.]+)?)\s+"
        r"([\d.]+(?:\s+or\s+[\d.]+)?)\s+"
        r"([\d.]+)",
        re.IGNORECASE,
    )
    named_rows: dict[str, tuple] = {}
    for m in row_pat.finditer(dim_text):
        label = m.group(1).strip().lower()
        if label in ("overall", "seat", "width", "depth", "height"):
            continue
        named_rows[label] = (m.group(2), m.group(3), m.group(4))

    if m_overall:
        w, d, h = m_overall.group(1), m_overall.group(2), m_overall.group(3)
    elif named_rows:
        pname = product_name.lower()
        selected_label, selected_vals = None, None
        for label, vals in named_rows.items():
            # 4-char stem covers "round"→"round" and "rectangle"→"rectangular"
            if label[:4] in pname:
                selected_label, selected_vals = label, vals
                break
        if selected_vals is None:
            selected_label, selected_vals = next(iter(named_rows.items()))
        # Diameter for round products
        if selected_label and "round" in selected_label:
            data.setdefault("Diameter", _first_dim_val(selected_vals[0]))
        w, d, h = selected_vals
    else:
        return

    data.setdefault("Width",  _first_dim_val(w))
    data.setdefault("Depth",  _first_dim_val(d))
    data.setdefault("Height", _first_dim_val(h))

    # Seat row: "Seat W D H"
    m_seat = re.search(
        r"Seat\s+([\d.]+(?:\s+or\s+[\d.]+)?)\s+([\d.]+(?:\s+or\s+[\d.]+)?)\s+([\d.]+)",
        dim_text, re.IGNORECASE,
    )
    if m_seat:
        data.setdefault("Seat Width",  _first_dim_val(m_seat.group(1)))
        data.setdefault("Seat Depth",  _first_dim_val(m_seat.group(2)))
        data.setdefault("Seat Height", _first_dim_val(m_seat.group(3)))

    # Build Dimensions string
    w_val = data.get("Width", "")
    d_val = data.get("Depth", "")
    h_val = data.get("Height", "")
    dim_parts = [f"W {w_val}" if w_val else "", f"D {d_val}" if d_val else "", f"H {h_val}" if h_val else ""]
    dim_str = " x ".join(p for p in dim_parts if p)
    if dim_str:
        data.setdefault("Dimensions", dim_str)


def _parse_fabric_description_tab(html_text: str, data: dict):
    """Parse detailed fabric specs from #tab-description panel.

    HTML structure: each <p> has a <strong> section name followed by
    <br/>-separated "Key: Value" lines.

    Captures all fields: Content, Backing, Fabric Weight, Fabric Width,
    Roll Size, Repeat H/V, Directional, Railroaded, Abrasion, Flammability,
    Recommended Cleaning, and every other key:value line found.
    """
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_text, "html.parser")
    panel = soup.find(id="tab-description")
    if not panel:
        return

    for para in panel.find_all("p"):
        # Each <p> may start with a <strong> section label — ignore it as a key
        strong = para.find("strong")
        section_label = clean_text(strong.get_text()) if strong else ""

        # Get the text nodes separated by <br> tags
        # Replace <br> with newline then split
        for br in para.find_all("br"):
            br.replace_with("\n")
        raw_lines = [ln.strip() for ln in para.get_text().splitlines()]

        for line in raw_lines:
            line = clean_text(line)
            if not line or line == section_label:
                continue  # skip the section heading itself

            if ":" not in line:
                # Whole-line value (e.g., Recommended Cleaning text)
                if section_label and len(line) > 5:
                    data.setdefault(section_label, line)
                continue

            # "Key: Value" pair
            key_raw, _, val_raw = line.partition(":")
            key = clean_text(key_raw)
            val = clean_text(val_raw)
            if not key or not val or len(key) > 80:
                continue

            # Rename to avoid collisions with chair/table fields
            if key == "Width":
                key = "Fabric Width"
            elif key == "Weight":
                key = "Fabric Weight"
            elif key == "Finish" and data.get("Finish"):
                key = "Finish Description"

            data.setdefault(key, val)

    # Repeat H/V — special format "Repeat H – 13.88 " V – 13.0 ""
    full_text = panel.get_text(separator=" ", strip=True)
    m_rep = re.search(
        r"Repeat\s+H\s*[–-]\s*([\d.]+).*?V\s*[–-]\s*([\d.]+)",
        full_text, re.IGNORECASE
    )
    if m_rep:
        data.setdefault("Horizontal Repeat", m_rep.group(1))
        data.setdefault("Vertical Repeat", m_rep.group(2))


def _parse_fabric_attrs_from_text(text: str, data: dict):
    """Extract Supplier/Grade/Finish/Color from summary plain text when no <table> found."""
    known_attrs = ["Supplier", "Grade", "Finish", "Color", "Pattern", "Content", "Width"]
    tokens = [t.strip() for t in text.split("|") if t.strip()]
    for i, tok in enumerate(tokens):
        if tok in known_attrs and i + 1 < len(tokens):
            val = tokens[i + 1]
            # Skip navigation noise
            if len(val) < 80 and val not in known_attrs and "Category" not in val:
                data.setdefault(tok, val)


def _extract_feature_fields(txt: str, data: dict):
    """Extract structured fields from a feature bullet text."""
    tl = txt.lower()
    if tl.startswith("weight:"):
        m = re.search(r"[\d.]+", txt)
        if m:
            data.setdefault("Weight", m.group())
    elif re.match(r"c\.o\.m\.\s*:", txt, re.IGNORECASE):
        # "C.O.M.: 1 Yard" — extract just the value
        val = re.sub(r"(?i)c\.o\.m\.\s*:\s*", "", txt).strip()
        data.setdefault("COM", val)
    elif "seat style:" in tl:
        data.setdefault("Seat Style", txt.split(":", 1)[1].strip())
    elif "weight capacity" in tl:
        data.setdefault("Weight Capacity", txt)
    elif "fire rating" in tl and "meets" in tl:
        data.setdefault("Fire Rating", txt)
    elif "stacks" in tl and "high" in tl:
        data.setdefault("Stack Height", txt)
    elif ("16 gauge" in tl or "aluminum" in tl) and "frame" in tl:
        data.setdefault("Frame", txt)
    elif "available finishes" in tl:
        data.setdefault("Finish", txt.split(":", 1)[1].strip() if ":" in txt else txt)
    elif "warranty" in tl and re.search(r"\d+.?year", tl):
        data.setdefault("Warranty", txt)


def _clean_dim_val(raw: str) -> str:
    """Strip inch marks and normalise a dimension value."""
    v = raw.replace('"', '').replace('"', '').replace('"', '').replace("'", '').strip()
    return v if v else ""


async def _safe_goto(page, url: str, retries: int = 3):
    """Navigate to a URL with retry on rate-limit / timeout."""
    for attempt in range(retries):
        try:
            resp = await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
            if resp and resp.status == 429:
                wait_s = 15 * (attempt + 1)
                print(f"  [429 rate-limit] waiting {wait_s}s before retry...")
                await page.wait_for_timeout(wait_s * 1000)
                continue
            return resp
        except Exception as e:
            if attempt < retries - 1:
                await page.wait_for_timeout(8000)
            else:
                raise e
    return None


async def scrape_product(page, url: str) -> list[dict]:
    """Scrape a single Shev Chair product page."""
    await _safe_goto(page, url)
    await page.wait_for_timeout(2500)

    html = await page.content()
    data = _parse_summary(html)
    data["Source URL"] = url
    data["Manufacturer"] = VENDOR_NAME

    if data.get("Product Name"):
        data["Product Family Id"] = extract_family_id(data["Product Name"])

    # Fabric description tab — detailed specs (Content, Backing, Weight, Width, Repeat, etc.)
    _parse_fabric_description_tab(html, data)

    # Image URL — prefer wp-post-image or woocommerce_single, avoid SVG/emoji
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    for selector in [
        "img.wp-post-image",
        "img.attachment-woocommerce_single",
        "div.woocommerce-product-gallery__image img",
        ".woocommerce-product-gallery img",
    ]:
        for img in soup.select(selector):
            src = (
                img.get("data-large_image")
                or img.get("data-src")
                or img.get("src", "")
            )
            if src and ".svg" not in src.lower() and "placeholder" not in src and "emoji" not in src:
                data["Image URL"] = src
                break
        if data.get("Image URL"):
            break

    return [data]


async def get_product_links(page, listing_url: str) -> list[str]:
    """Collect all product URLs from a listing page with pagination."""
    links: list[str] = []
    url = listing_url

    while url:
        await _safe_goto(page, url)
        await page.wait_for_timeout(2000)

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(await page.content(), "html.parser")

        # Collect unique product hrefs
        for a in soup.select(".woocommerce ul.products li a[href]"):
            href = a["href"]
            if "/products/" in href and href not in links:
                links.append(href)

        # Next page
        next_a = soup.select_one("a.next.page-numbers, .woocommerce-pagination a.next")
        url = next_a["href"] if next_a and next_a.get("href") else None

    return links


async def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    if TEST_MODE:
        print(f"[TEST: max {TEST_MAX_PRODUCTS} products per category]")

    cats = info["categories"]
    if TEST_MODE:
        cats = cats[:TEST_MAX_CATEGORIES]

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in cats:
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

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"[{cat['name']}] {len(all_product_urls)} products to scrape")

            global_idx = 1
            for url in all_product_urls:
                try:
                    rows = await scrape_product(page, url)
                    for row in rows:
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(VENDOR_NAME, cat["name"], global_idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"  ERROR {url}: {e}")
                # Polite delay — 4-6 s to respect rate limits
                await page.wait_for_timeout(4000)
                await async_polite_delay()

            # Auto-save after every category — protects against mid-run crashes
            writer.save()
            print(f"  [Auto-saved after {cat['name']} — {global_idx - 1} products]")

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
