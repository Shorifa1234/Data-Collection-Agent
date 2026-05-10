"""
Mercana scraper — pure requests, no Playwright (avoids Cloudflare challenge
on product detail pages that headless browsers trigger).

Strategy:
  1. Listing page  → requests.get  → parse .product-item hrefs
  2. Product page  → requests.get  → extract parent_id, CSRF token, variant list
  3. Per variant   → requests.post → /shoppingcart/ProductDetails_iProductChange
                                     → JSON with all dynamic fields (name, price,
                                       dimensions, image, specs, etc.)
  4. Attachments   → requests.get  → /ProductTab/GetProductAttachmentsTabAjax/{id}
"""

import json
import os
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text,
    clean_price,
    generate_sku,
    extract_family_id,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Mercana")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"  # unused; kept for compat
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://marketplace.mercana.com"
_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)
POLITE_DELAY = 1.5   # seconds between product requests


# ---------------------------------------------------------------------------
# Session helper
# ---------------------------------------------------------------------------

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": _UA})
    return s


# ---------------------------------------------------------------------------
# Listing page
# ---------------------------------------------------------------------------

def get_product_links(session: requests.Session, listing_url: str) -> list[str]:
    """Return all product-family URLs from a category listing page."""
    r = session.get(listing_url, timeout=20)
    soup = BeautifulSoup(r.text, "html.parser")

    links: list[str] = []
    seen: set[str] = set()
    for item in soup.select(".product-item"):
        a = item.select_one(".picture a[href]") or item.select_one("a[href]")
        if not a:
            continue
        href = a["href"]
        if href.startswith("/"):
            href = BASE_URL + href
        if href not in seen:
            seen.add(href)
            links.append(href)

    return links


# ---------------------------------------------------------------------------
# Product page helpers
# ---------------------------------------------------------------------------

def parse_static_product_page(html: str) -> dict:
    """Parse static product page HTML for parent_id, CSRF token, variants."""
    soup = BeautifulSoup(html, "html.parser")

    # Parent product ID from the SKU element  id="sku-45907"
    parent_id = None
    sku_el = soup.select_one('[id^="sku-"]')
    if sku_el:
        m = re.search(r"sku-(\d+)", sku_el["id"])
        if m:
            parent_id = m.group(1)

    # Fall back: attribute_change_handler_XXXXX in script
    if not parent_id:
        m = re.search(r"attribute_change_handler_(\d+)", html)
        if m:
            parent_id = m.group(1)

    # CSRF token
    token_el = soup.select_one('input[name="__RequestVerificationToken"]')
    token = token_el["value"] if token_el else ""

    # Variants from swatch list
    variants: list[dict] = []
    for li in soup.select(".option-list.image-squares li"):
        vid = li.get("data-attr-value", "")
        tooltip = li.select_one(".tooltip-header")
        finish = clean_text(tooltip.text.strip()) if tooltip else ""
        sq = li.select_one(".attribute-square")
        bg = sq.get("style", "") if sq else ""
        m2 = re.search(r"url\('([^']+)'\)", bg)
        thumbnail = m2.group(1) if m2 else ""
        if vid:
            variants.append({"id": vid, "finish": finish, "thumbnail": thumbnail})

    # Family-level description (shared across variants)
    desc_el = soup.select_one(".short-description")
    family_desc = clean_text(desc_el.text.strip()) if desc_el else ""

    # Family name (plain, without variant detail)
    h1 = soup.select_one(".mercana-product-attribute-availability h1, h1")
    family_name = clean_text(h1.text.strip()) if h1 else ""

    return {
        "parent_id": parent_id,
        "token": token,
        "variants": variants,
        "family_desc": family_desc,
        "family_name": family_name,
    }


def fetch_variant_ajax(
    session: requests.Session,
    parent_id: str,
    variant_id: str,
    token: str,
    referer: str,
) -> dict:
    """POST to the Mercana product-change AJAX endpoint for variant data."""
    try:
        r = session.post(
            f"{BASE_URL}/shoppingcart/ProductDetails_iProductChange?productId={parent_id}",
            data={"iProductId": variant_id, "__RequestVerificationToken": token},
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "Referer": referer,
            },
            timeout=35,
        )
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return {}


def fetch_attachments(session: requests.Session, parent_id: str) -> dict:
    """Fetch Downloads tab via AJAX → assembly + care instruction links."""
    result: dict = {}
    try:
        r = session.get(
            f"{BASE_URL}/ProductTab/GetProductAttachmentsTabAjax/{parent_id}",
            timeout=30,
        )
        soup = BeautifulSoup(r.text, "html.parser")
        for a in soup.select("a[href]"):
            href = a.get("href", "")
            text = a.text.strip()
            if not href or not text:
                continue
            full = BASE_URL + href if href.startswith("/") else href
            tl = text.lower()
            if "assembly" in tl:
                result.setdefault("Assembly Instructions", full)
            elif "care" in tl or "clean" in tl:
                result.setdefault("Care Instructions", full)
    except Exception:
        pass
    return result


# ---------------------------------------------------------------------------
# Core parser: AJAX JSON → flat product dict
# ---------------------------------------------------------------------------

def parse_ajax_to_row(
    ajax: dict,
    source_url: str,
    finish: str = "",
    family_desc: str = "",
) -> dict:
    """Convert the AJAX JSON payload into a flat product row dict."""
    row: dict = {"Source": source_url}

    # SKU (product number, e.g. "71235")
    if ajax.get("sku"):
        row["SKU"] = ajax["sku"]

    # Finish / color from swatch tooltip
    if finish:
        row["Finish"] = finish

    # Product name + description from backinstocksubscriptionhtml
    avail_html = ajax.get("backinstocksubscriptionhtml") or ""
    if avail_html:
        s = BeautifulSoup(avail_html, "html.parser")
        h1 = s.select_one("h1")
        if h1:
            row["Product Name"] = clean_text(h1.text.strip())
        d = s.select_one(".short-description")
        if d:
            txt = clean_text(d.text.strip())
            if txt:
                row["Description"] = txt

    # Fall back to family description if per-variant not found
    if not row.get("Description") and family_desc:
        row["Description"] = family_desc

    # Price from productpriceinfohtml: "Suggested Retail Price US$1,249.99"
    price_html = ajax.get("productpriceinfohtml") or ""
    if price_html:
        m = re.search(r"US\$\s*([\d,]+\.?\d*)", price_html)
        if m:
            row["Price"] = clean_price(m.group(1))

    # Image URL from productcloudzoomnhtml
    img_html = ajax.get("productcloudzoomnhtml") or ""
    if img_html:
        s = BeautifulSoup(img_html, "html.parser")
        # Prefer data-full-image-url attribute on the anchor
        a = s.select_one("a[data-full-image-url]")
        if a:
            row["Image URL"] = a["data-full-image-url"]
        else:
            img = s.select_one("img.cloudzoom, img[src*='media.mercana.com']")
            if img:
                row["Image URL"] = img.get("src", "")

    # Dimensions, Weight, Box Size, Box Weight from productdetailsinfohtml
    detail_html = ajax.get("productdetailsinfohtml") or ""
    if detail_html:
        s = BeautifulSoup(detail_html, "html.parser")
        seen_labels: set = set()
        for d_el in s.select(".details"):
            lbl_el = d_el.select_one(".label")
            val_el = d_el.select_one(".value")
            if not (lbl_el and val_el):
                continue
            lbl = lbl_el.text.strip()
            val = val_el.text.strip()
            if lbl in seen_labels:
                continue
            seen_labels.add(lbl)

            if lbl == "Product Size (Inches)":
                row["Dimensions"] = val          # ExcelWriter auto-parses into L/W/H
            elif lbl == "Weight":
                m2 = re.search(r"([\d.]+)\s*lbs", val)
                if m2:
                    row["Weight"] = m2.group(1)
            elif lbl == "Box Size (Inches)":
                row["Box Size"] = val
            elif lbl == "Box Weight":
                m2 = re.search(r"([\d.]+)\s*lbs", val)
                if m2:
                    row["Box Weight"] = m2.group(1) + " lbs"

    # Feature bullet points: <li><p><strong>Title</strong></p><p><span>Body</span></p></li>
    bullet_html = ajax.get("productbulletpointhtml") or ""
    if bullet_html:
        s = BeautifulSoup(bullet_html, "html.parser")
        parts: list[str] = []
        for li in s.select("li"):
            strong = li.select_one("strong")
            span   = li.select_one("span")
            if strong:
                title = clean_text(strong.text.strip())
                body  = clean_text(span.text.strip()) if span else ""
                if title:
                    parts.append(f"{title}: {body}" if body else title)
        if parts:
            row["Features"] = " | ".join(parts)

    # Full description (Features tab body text)
    full_desc_html = ajax.get("productfulldescriptionhtml") or ""
    if full_desc_html:
        s = BeautifulSoup(full_desc_html, "html.parser")
        txt = clean_text(s.get_text(separator=" ").strip())
        if txt:
            row["Full Description"] = txt

    # Specifications from productspecificationshtml
    spec_html = ajax.get("productspecificationshtml") or ""
    if spec_html:
        s = BeautifulSoup(spec_html, "html.parser")
        for tr in s.select("tr"):
            name_td = tr.select_one("td.spec-name")
            val_td  = tr.select_one("td.spec-value")
            if name_td and val_td:
                name = clean_text(name_td.text.strip())
                val  = clean_text(val_td.text.strip())
                if name and val:
                    row.setdefault(name, val)

    # Spec Sheet URL: constructed from product number
    sku_num = ajax.get("sku", "")
    if sku_num:
        row.setdefault(
            "Spec Sheet",
            f"{BASE_URL}/files/product_spec_sheet/ProductSpecSheet_{sku_num}.pdf",
        )

    return row


# ---------------------------------------------------------------------------
# Top-level product scraper
# ---------------------------------------------------------------------------

def _get_with_retry(session: requests.Session, url: str, retries: int = 3) -> requests.Response | None:
    """GET with retries on timeout/connection errors."""
    for attempt in range(retries):
        try:
            return session.get(url, timeout=35)
        except Exception as e:
            if attempt < retries - 1:
                wait = (attempt + 1) * 5
                print(f"  Retry {attempt+1}/{retries} for {url} (waiting {wait}s): {e}")
                time.sleep(wait)
            else:
                print(f"  GET failed after {retries} attempts {url}: {e}")
    return None


def scrape_product(session: requests.Session, url: str) -> list[dict]:
    """Return one dict per variant for a product family URL."""
    r = _get_with_retry(session, url)
    if r is None:
        return []

    page_data = parse_static_product_page(r.text)
    parent_id   = page_data["parent_id"]
    token       = page_data["token"]
    variants    = page_data["variants"]
    family_desc = page_data["family_desc"]

    if not parent_id:
        print(f"  Cannot find parent_id for {url}")
        return []

    # Fetch attachment links once per product family
    attachments = fetch_attachments(session, parent_id)

    rows: list[dict] = []

    # If no variant swatches, call AJAX with iProductId=parent_id (single-variant product)
    if not variants:
        ajax = fetch_variant_ajax(session, parent_id, parent_id, token, url)
        time.sleep(POLITE_DELAY)
        row = parse_ajax_to_row(ajax, url, family_desc=family_desc)
        row.update(attachments)
        rows.append(row)
        return rows

    # One row per variant
    for variant in variants:
        vid    = variant["id"]
        finish = variant["finish"]
        ajax = fetch_variant_ajax(session, parent_id, vid, token, url)
        time.sleep(POLITE_DELAY)

        row = parse_ajax_to_row(ajax, url, finish=finish, family_desc=family_desc)
        # Attachments override spec-table filenames (attachments have full URLs)
        row.update(attachments)

        # Fallback image from swatch thumbnail if AJAX didn't provide one
        if not row.get("Image URL") and variant.get("thumbnail"):
            thumb = variant["thumbnail"]
            # Convert thumbnail to hires: /thumbs/010/0100039_71235_A_304.jpeg → /0100039_71235_A.jpeg
            hires = re.sub(r"/thumbs/\w+/", "/", thumb)
            hires = re.sub(r"_\d+\.jpeg", ".jpeg", hires)
            row["Image URL"] = hires

        rows.append(row)

    return rows


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])
    session = make_session()

    if TEST_MODE:
        print(f"[TEST: all categories, max {TEST_MAX_PRODUCTS} products each]")

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]

    for cat in categories:
        if not cat["links"]:
            continue

        writer.add_sheet(
            cat["name"],
            cat["links"][0],
            studio_columns=cat["studio_columns"],
        )

        # Collect product URLs from ALL listing links → deduplicate
        seen_urls: set[str] = set()
        all_product_urls: list[str] = []
        for listing_url in cat["links"]:
            for u in get_product_links(session, listing_url):
                if u not in seen_urls:
                    seen_urls.add(u)
                    all_product_urls.append(u)

        if TEST_MODE:
            all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

        print(f"  [{cat['name']}] {len(all_product_urls)} products")

        global_idx = 1
        for product_url in all_product_urls:
            try:
                variant_rows = scrape_product(session, product_url)
                for row in variant_rows:
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
                print(f"  ERROR [{cat['name']}] {product_url}: {e}")
                global_idx += 1

            time.sleep(POLITE_DELAY)

    writer.save()


if __name__ == "__main__":
    main()
