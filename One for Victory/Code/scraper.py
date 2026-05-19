import json, os, sys, re, time
import requests
from pathlib import Path
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "One for Victory")
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
                        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL     = "https://www.oneforvictory.com"
US_HP_URL    = "https://www.ultrasuede.us/products/hp.html"
US_IMG_BASE  = "https://www.ultrasuede.us/products/"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
}

# ─── helpers ───────────────────────────────────────────────────────────────

def get_json(url: str) -> dict:
    sep = "&" if "?" in url else "?"
    r = requests.get(url + sep + "format=json", headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()


def polite_delay(secs: float = 1.0):
    time.sleep(secs)


def strip_inch(val: str) -> str:
    """Remove inch marks and surrounding whitespace."""
    return re.sub(r'[“”’‘"ââ]', "", val).strip()


# ─── excerpt parsers ────────────────────────────────────────────────────────

DIM_MAP = {
    "depth": "Depth",
    "width": "Width",
    "height": "Height",
    "heights": "Height",       # plural variant on some products
    "sh": "Seat Height",
    "sd": "Seat Depth",
    "seat height": "Seat Height",
    "seat depth": "Seat Depth",
    "ah": "Arm Height",
    "arm height": "Arm Height",
    "seat length": "Seat Length",
    "sl": "Seat Length",
    "dia": "Diameter",
    "diameter": "Diameter",
    "length": "Length",
}

SEPARATOR_RE = re.compile(
    r"^[-_——―●ââ]{2,}$"
)


def parse_furniture_excerpt(html: str) -> dict:
    """
    Parse furniture product HTML excerpt.

    Returns a flat dict with dimension fields, Finish, Description, etc.
    """
    soup = BeautifulSoup(html, "html.parser")
    raw = soup.get_text(separator="\n", strip=True)
    lines = [l.strip() for l in raw.split("\n") if l.strip()]

    data: dict = {}
    desc_lines: list[str] = []
    past_sep = False
    pending_finish = False

    SKIP_HEADERS = {"product information", "product dimensions", "color", "email us"}

    for line in lines:
        ll = line.lower().strip()

        # Section header lines -> skip
        if ll in SKIP_HEADERS:
            continue

        # Separator line -> description begins after
        if SEPARATOR_RE.match(line) or line in ("——", "——————", "____", "_____"):
            past_sep = True
            pending_finish = False
            continue

        # Handle "As Shown:" value that was on the previous (empty) line
        if pending_finish and not past_sep:
            if line:
                data["Finish"] = line
                pending_finish = False
            continue

        # "Key: Value" pattern — only before separator
        if not past_sep:
            colon_m = re.match(r"^([^:]+):\s*(.*)$", line)
            if colon_m:
                key = colon_m.group(1).strip()
                val = colon_m.group(2).strip()
                key_l = key.lower()

                if key_l in DIM_MAP:
                    data[DIM_MAP[key_l]] = strip_inch(val)
                    continue

                # Catch: "As Shown", "As Show" (typo), "Shown in", "Color"
                if (
                    "as shown" in key_l
                    or "as show" in key_l
                    or key_l.startswith("shown in")
                    or key_l.startswith("shown:")
                ):
                    val_clean = val.strip()
                    if val_clean:
                        data["Finish"] = val_clean
                    else:
                        pending_finish = True
                    continue

        # Description paragraphs (after separator)
        if past_sep:
            skip_phrases = {
                "email us", "@oneforvictory.com", "contact us", "our products do not contain"
            }
            if not any(p in ll for p in skip_phrases) and ll:
                desc_lines.append(line)

    if desc_lines:
        # Keep the first 3 paragraphs as description
        data["Description"] = " ".join(desc_lines[:3])

    return data


def parse_fabric_excerpt(html: str) -> dict:
    """Parse fabric item HTML excerpt -> flat dict."""
    soup = BeautifulSoup(html, "html.parser")
    data: dict = {}

    # Collection name + PC code from first header/strong
    header = soup.find(["h3", "h2"]) or soup.find("strong")
    if header:
        coll = header.get_text(strip=True)
        data["Collection"] = coll
        m = re.search(r"\(PC\s*([^)]+)\)", coll, re.I)
        if m:
            data["PC Code"] = "PC " + m.group(1).strip()

    # Bullet list items
    for li in soup.select("li"):
        txt = li.get_text(strip=True)
        if ":" not in txt:
            continue
        k, _, v = txt.partition(":")
        k, v = k.strip(), v.strip()
        if k == "Contents":
            data["Material"] = v
        elif k == "Double Rubs":
            data["Double Rubs"] = v
        elif k == "Cleaning Code":
            data["Cleaning Code"] = v
        elif k == "Cleaning Maintenance" and v:
            data["Cleaning Maintenance"] = v
        elif k == "Performance" and v:
            data["Performance"] = v

    return data


def parse_leather_excerpt(html: str) -> dict:
    """Parse leather item HTML excerpt -> flat dict."""
    soup = BeautifulSoup(html, "html.parser")
    data: dict = {}

    # Collection header
    header = soup.find(["h3", "h2"]) or soup.find("strong")
    if header:
        coll = header.get_text(strip=True)
        data["Collection"] = coll
        m = re.search(r"\(PC\s*([^)]+)\)", coll, re.I)
        if m:
            data["PC Code"] = "PC " + m.group(1).strip()

    header_text = header.get_text(strip=True) if header else ""

    # Description paragraphs + leather characteristics bullets
    desc_parts: list[str] = []
    chars: list[str] = []
    in_chars = False

    for el in soup.find_all(["p", "li", "strong"]):
        txt = el.get_text(strip=True)
        if not txt:
            continue
        if "leather characteristics" in txt.lower():
            in_chars = True
            continue
        if "variances in color" in txt.lower():
            continue
        if txt == header_text:
            continue
        if in_chars and el.name == "li":
            chars.append(txt)
        elif not in_chars and el.name == "p":
            desc_parts.append(txt)

    if desc_parts:
        data["Description"] = " ".join(desc_parts[:2])

    if chars:
        data["Leather Characteristics"] = "; ".join(chars)
        for c in chars:
            # e.g. "1.1-1.3 thickness" or "1.0-1.1 thickness"
            if re.search(r"\d[\d.]*\s*[-–]\s*\d[\d.]*", c):
                data["Thickness"] = c

    return data


# ─── ultrasuede scraper ─────────────────────────────────────────────────────

def get_ultrasuede_rows() -> list[dict]:
    """
    Fetch ultrasuede.us/products/hp.html.
    Returns one row per color swatch.
    """
    r = requests.get(US_HP_URL, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.content, "html.parser")

    # Build code -> image filename map from actual <img> tags
    img_map: dict[str, str] = {}
    for img in soup.select("img[src*='hp_']"):
        src = img.get("src", "")
        m = re.match(r"images/(hp_(\d+)_[^.]+\.jpg)", src)
        if m:
            img_map[m.group(2)] = m.group(1)  # code -> filename

    # Shared specs from <dl> elements
    specs: dict[str, str] = {}
    spec_vals: list[str] = []
    for dl in soup.select("dl"):
        dt = dl.select_one("dt")
        dd = dl.select_one("dd")
        if dt and dd:
            k = dt.get_text(strip=True).rstrip(":").strip()
            v = dd.get_text(strip=True)
            specs[k] = v
            spec_vals.append(f"{k}: {v}")

    # Features bullet list
    features = [li.get_text(strip=True) for li in soup.select(".productsFutureList li")]
    description = "; ".join(features[:6]) if features else ""

    # Composition may span two dd elements — merge if needed
    composition = specs.get("Composition", "")
    # Also grab second part from raw text if split across nodes
    comp_raw_parts = []
    for dl in soup.select("dl"):
        dt = dl.select_one("dt")
        if dt and "composition" in dt.get_text(strip=True).lower():
            for dd in dl.select("dd"):
                comp_raw_parts.append(dd.get_text(strip=True))
    if comp_raw_parts:
        composition = " ".join(comp_raw_parts)

    # Color palette: text lines alternate between 4-digit code and color name
    body_lines = [l.strip() for l in soup.get_text(separator="\n").split("\n") if l.strip()]
    colors: list[tuple[str, str]] = []  # (code, name)
    i = 0
    palette_start = False
    while i < len(body_lines):
        if "color palette" in body_lines[i].lower():
            palette_start = True
            i += 1
            continue
        if palette_start:
            code_line = body_lines[i].strip()
            if re.match(r"^\d{4}$", code_line) and i + 1 < len(body_lines):
                name_line = body_lines[i + 1].strip()
                # Only add if name is not another code
                if not re.match(r"^\d{4}$", name_line):
                    colors.append((code_line, name_line))
                    i += 2
                    continue
            # Stop if we've left the color section (hit a non-code/non-name block)
            if len(body_lines[i]) > 60:
                break
        i += 1

    # Build product rows
    rows: list[dict] = []
    for code, name in colors:
        # Use actual image filename from img_map, else construct it
        if code in img_map:
            img_filename = img_map[code]
        else:
            img_filename = f"hp_{code}_{name.replace(' ', '')}.jpg"

        row: dict = {
            "Product Name":       f"Ultrasuede HP {name}",
            "Product Family Id":  "Ultrasuede HP",
            "SKU":                f"HP{code}",
            "Source URL":         US_HP_URL,
            "Image URL":          US_IMG_BASE + "images/" + img_filename,
            "Collection":         "Ultrasuede HP",
            "Color":              name,
            "Style":              specs.get("Style", ""),
            "Description":        description,
            "Specifications":     "; ".join(spec_vals),
        }
        if specs.get("Width"):
            row["Width"] = specs["Width"]
        if specs.get("Thickness"):
            row["Thickness"] = specs["Thickness"]
        if specs.get("Weight"):
            row["Weight"] = specs["Weight"]
        if specs.get("Fiber Fineness"):
            row["Fiber Fineness"] = specs["Fiber Fineness"]
        if specs.get("Put up"):
            row["Put Up"] = specs["Put up"]
        if composition:
            row["Material"] = composition

        rows.append(row)

    return rows


# ─── listing fetchers ───────────────────────────────────────────────────────

def fetch_squarespace_items(url: str) -> list[dict]:
    """
    Fetch all items from a Squarespace collection via ?format=json.
    Works for product-line?category=X, /fabrics, /leathers.
    """
    data = get_json(url)
    return data.get("items", [])


# ─── main ───────────────────────────────────────────────────────────────────

def main():
    info = json.loads((Path(__file__).parent / "vendor_info.json").read_text(encoding="utf-8"))
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    if TEST_MODE:
        print(f"[TEST: max {TEST_MAX_PRODUCTS} products per category, all categories]")

    for cat in info["categories"]:
        if not cat["links"]:
            continue

        cat_name     = cat["name"]
        links        = cat["links"]
        studio_cols  = cat["studio_columns"]
        max_prod     = TEST_MAX_PRODUCTS if TEST_MODE else None

        writer.add_sheet(cat_name, links[0], studio_columns=studio_cols)
        print(f"\n[{cat_name}] ({len(links)} link(s))")

        global_idx = 1
        written    = 0

        # Determine collection type from links
        is_fabric  = any("fabrics" in l or "ultrasuede" in l for l in links)
        is_leather = any("leathers" in l for l in links)

        # ── FABRIC ──────────────────────────────────────────────────────────
        if is_fabric:
            for link in links:
                if max_prod and written >= max_prod:
                    break

                is_us_page = "ultrasuede" in link and "oneforvictory" in link

                if is_us_page:
                    print(f"  -> Fetching Ultrasuede HP colors …")
                    try:
                        us_rows = get_ultrasuede_rows()
                        print(f"     {len(us_rows)} colors found")
                        for row in us_rows:
                            if max_prod and written >= max_prod:
                                break
                            row["Manufacturer"] = VENDOR_NAME
                            if not row.get("SKU"):
                                row["SKU"] = generate_sku(VENDOR_NAME, cat_name, global_idx)
                            writer.write_row(row, category_name=cat_name)
                            global_idx += 1
                            written += 1
                    except Exception as e:
                        print(f"  [ERROR] Ultrasuede: {e}")
                    polite_delay()

                else:
                    print(f"  -> {link}")
                    try:
                        items = fetch_squarespace_items(link)
                        print(f"     {len(items)} fabric items")
                        for item in items:
                            if max_prod and written >= max_prod:
                                break
                            try:
                                product_name = item.get("title", "")
                                full_url     = item.get("fullUrl", "")
                                source_url   = (
                                    BASE_URL + full_url
                                    if full_url.startswith("/")
                                    else full_url
                                )
                                image_url    = item.get("assetUrl", "")
                                variants     = item.get("variants", [])
                                sku          = variants[0].get("sku", "") if variants else ""
                                if not sku:
                                    sku = generate_sku(VENDOR_NAME, cat_name, global_idx)

                                row: dict = {
                                    "Product Name":      product_name,
                                    "Product Family Id": extract_family_id(product_name),
                                    "Manufacturer":      VENDOR_NAME,
                                    "Source URL":        source_url,
                                    "Image URL":         image_url,
                                    "SKU":               sku,
                                }

                                exc_data = parse_fabric_excerpt(item.get("excerpt", ""))
                                row.update(exc_data)

                                # Fabric type from non-prefixed categories
                                item_cats = item.get("categories", [])
                                fabric_types = [c for c in item_cats if not c.startswith("-")]
                                if fabric_types:
                                    row["Fabric Type"] = ", ".join(fabric_types)

                                # Color: strip the first word (collection prefix) from the
                                # product name — works even when spelling differs slightly
                                # e.g. "Mate Sand" (collection "Matte"), "Towne Cloud" ("Town")
                                if " " in product_name:
                                    color_raw = product_name.split(" ", 1)[1].strip()
                                    # Strip emoji and non-printable chars
                                    color_clean = re.sub(r"[^\w\s,.'-]", "", color_raw).strip()
                                    if color_clean:
                                        row["Color"] = color_clean
                                else:
                                    # Single-word product — fall back to "-" prefixed categories
                                    color_cats = [
                                        c.lstrip("-").strip()
                                        for c in item_cats if c.startswith("-")
                                    ]
                                    if color_cats:
                                        row["Color"] = ", ".join(color_cats)

                                writer.write_row(row, category_name=cat_name)
                                global_idx += 1
                                written += 1
                            except Exception as e:
                                print(f"  [ERROR] fabric '{item.get('title','?')}': {e}")
                    except Exception as e:
                        print(f"  [ERROR] {link}: {e}")
                    polite_delay()

        # ── LEATHER ─────────────────────────────────────────────────────────
        elif is_leather:
            for link in links:
                if max_prod and written >= max_prod:
                    break
                print(f"  -> {link}")
                try:
                    items = fetch_squarespace_items(link)
                    print(f"     {len(items)} leather items")
                    for item in items:
                        if max_prod and written >= max_prod:
                            break
                        try:
                            product_name = item.get("title", "")
                            full_url     = item.get("fullUrl", "")
                            source_url   = (
                                BASE_URL + full_url
                                if full_url.startswith("/")
                                else full_url
                            )
                            image_url    = item.get("assetUrl", "")
                            variants     = item.get("variants", [])
                            sku          = variants[0].get("sku", "") if variants else ""
                            if not sku:
                                sku = generate_sku(VENDOR_NAME, cat_name, global_idx)

                            row = {
                                "Product Name":      product_name,
                                "Product Family Id": extract_family_id(product_name),
                                "Manufacturer":      VENDOR_NAME,
                                "Source URL":        source_url,
                                "Image URL":         image_url,
                                "SKU":               sku,
                            }

                            exc_data = parse_leather_excerpt(item.get("excerpt", ""))
                            row.update(exc_data)

                            writer.write_row(row, category_name=cat_name)
                            global_idx += 1
                            written += 1
                        except Exception as e:
                            print(f"  [ERROR] leather '{item.get('title','?')}': {e}")
                    polite_delay()
                except Exception as e:
                    print(f"  [ERROR] {link}: {e}")

        # ── FURNITURE / SEATING ──────────────────────────────────────────────
        else:
            seen_urls: set[str] = set()
            all_items: list[dict] = []

            for link in links:
                print(f"  -> {link}")
                try:
                    items = fetch_squarespace_items(link)
                    print(f"     {len(items)} products")
                    for item in items:
                        furl = item.get("fullUrl", "")
                        if furl not in seen_urls:
                            seen_urls.add(furl)
                            all_items.append(item)
                    polite_delay()
                except Exception as e:
                    print(f"  [ERROR] {link}: {e}")

            print(f"  Unique products: {len(all_items)}")

            for item in all_items:
                if max_prod and written >= max_prod:
                    break
                try:
                    product_name = item.get("title", "")
                    full_url     = item.get("fullUrl", "")
                    source_url   = (
                        BASE_URL + full_url if full_url.startswith("/") else full_url
                    )
                    image_url    = item.get("assetUrl", "")
                    variants     = item.get("variants", [])
                    sku          = variants[0].get("sku", "") if variants else ""
                    if not sku:
                        sku = generate_sku(VENDOR_NAME, cat_name, global_idx)

                    price_cents = item.get("priceCents", 0)
                    price = price_cents / 100 if price_cents else None

                    row = {
                        "Product Name":      product_name,
                        "Product Family Id": extract_family_id(product_name),
                        "Manufacturer":      VENDOR_NAME,
                        "Source URL":        source_url,
                        "Image URL":         image_url,
                        "SKU":               sku,
                    }
                    if price:
                        row["Price"] = price

                    exc_data = parse_furniture_excerpt(item.get("excerpt", ""))
                    row.update(exc_data)

                    writer.write_row(row, category_name=cat_name)
                    global_idx += 1
                    written += 1
                except Exception as e:
                    print(f"  [ERROR] '{item.get('title','?')}': {e}")

        print(f"  OK {written} rows written for [{cat_name}]")

    writer.save()
    print(f"\nSaved -> {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
