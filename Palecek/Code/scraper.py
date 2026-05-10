"""
scraper.py  —  Palecek
------------------------
Platform: palecek.com (custom ASP.NET CMS — fully server-rendered, no JS needed)

APPROACH: requests + BeautifulSoup only — NO Playwright.
  Both listing pages and product detail pages return complete HTML via a regular
  HTTP request. Playwright was causing VPS failures because Cloudflare blocked
  the headless browser; switching to requests bypasses that entirely.

Listing URL format : https://www.palecek.com/{category-slug}/
Product URL format : https://www.palecek.com/{product-slug}/{sku}-1822/iteminformation.aspx
Pagination         : next-page link in HTML (href contains "page=N")

Fields collected:
  Product Name, SKU, Image URL, Source, Description,
  Dimensions, Width/Depth/Height/Diameter, Weight,
  Materials, Finish, Collection, Designer, Origin, Lead Time,
  Seat Height, Seat Depth, Arm Height, COM, COL, Fabric,
  Wattage, Socket, Color Temperature, Chain Length,
  Canopy, Shade Details, Tearsheet Link

Run directly:
    python scraper.py

Or via orchestrator:
    python orchestrator.py "Palecek"
    python orchestrator.py "Palecek" --test
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse, urlencode

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
    parse_dimensions,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "Palecek")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(
    os.environ.get(
        "OUTPUT_PATH",
        str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx"),
    )
)
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.palecek.com"


# ---------------------------------------------------------------------------
# HTTP session
# ---------------------------------------------------------------------------

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;"
            "q=0.9,image/avif,image/webp,*/*;q=0.8"
        ),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Cache-Control": "max-age=0",
        "Upgrade-Insecure-Requests": "1",
    })
    return s


_SESSION = make_session()


def _get(url: str, retries: int = 3) -> BeautifulSoup | None:
    """GET a URL and return a BeautifulSoup object, or None on failure."""
    for attempt in range(retries):
        try:
            resp = _SESSION.get(url, timeout=30)
            resp.raise_for_status()
            return BeautifulSoup(resp.text, "html.parser")
        except Exception as e:
            print(f"    [WARN] GET failed (attempt {attempt + 1}): {e}")
            if attempt < retries - 1:
                time.sleep(3)
    return None


# ---------------------------------------------------------------------------
# Listing page — collect product URLs
# ---------------------------------------------------------------------------

def get_product_links(listing_url: str, max_products: int = 0) -> list[str]:
    """
    Collect all product URLs from a Palecek category listing page.
    Follows pagination via "next page" links in the HTML.
    Product URLs end with /iteminformation.aspx.
    """
    links: list[str] = []
    seen:  set[str]  = set()
    url = listing_url

    while url:
        print(f"  [Listing] {url}")
        soup = _get(url)
        if soup is None:
            break

        # Extract all product links
        for a in soup.select("a[href*='iteminformation.aspx']"):
            href = a.get("href", "")
            if not href:
                continue
            full = urljoin(BASE_URL, href).split("?")[0]
            if full not in seen:
                seen.add(full)
                links.append(full)

        print(f"  [Listing] {len(links)} total links so far")

        if max_products and len(links) >= max_products:
            break

        # Find next page link — look for href containing "page="
        next_url = None
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if "page=" in href:
                text = clean_text(a.get_text()).lower()
                # Accept "next", "→", ">", or a page number link
                if any(kw in text for kw in ("next", "›", ">", "»")):
                    next_url = urljoin(BASE_URL, href)
                    break
        # Fallback: look for incrementing page= link
        if not next_url:
            current_page = _current_page_num(url)
            for a in soup.find_all("a", href=True):
                href = a["href"]
                m = re.search(r"page=(\d+)", href)
                if m and int(m.group(1)) == current_page + 1:
                    next_url = urljoin(BASE_URL, href)
                    break

        url = next_url
        if url:
            time.sleep(1.0)

    return links[:max_products] if max_products else links


def _current_page_num(url: str) -> int:
    """Extract current page number from URL, default 1."""
    m = re.search(r"page=(\d+)", url)
    return int(m.group(1)) if m else 1


# ---------------------------------------------------------------------------
# Product detail page
# ---------------------------------------------------------------------------

# Text patterns → field names
_FIELD_PATTERNS: list[tuple[str, str]] = [
    (r"Overall Dimensions?\s*[:\-]\s*(.+)",         "Dimensions"),
    (r"\bWidth\s*[:\-]\s*([\d.\"\']+)",              "Width"),
    (r"\bDepth\s*[:\-]\s*([\d.\"\']+)",              "Depth"),
    (r"\bHeight\s*[:\-]\s*([\d.\"\']+)",             "Height"),
    (r"\bDiameter\s*[:\-]\s*([\d.\"\']+)",           "Diameter"),
    (r"\bWeight\s*[:\-]\s*(.+)",                     "Weight"),
    (r"\bFinish\s*[:\-]\s*(.+)",                     "Finish"),
    (r"\bMaterials?\s*[:\-]\s*(.+)",                 "Materials"),
    (r"\bCollection\s*[:\-]\s*(.+)",                 "Collection"),
    (r"\bDesigner\s*[:\-]\s*(.+)",                   "Designer"),
    (r"\bOrigin\s*[:\-]\s*(.+)",                     "Origin"),
    (r"\bCountry of Origin\s*[:\-]\s*(.+)",          "Origin"),
    (r"\bLead Time\s*[:\-]\s*(.+)",                  "Lead Time"),
    (r"\bSeat Height\s*[:\-]\s*(.+)",                "Seat Height"),
    (r"\bSeat Depth\s*[:\-]\s*(.+)",                 "Seat Depth"),
    (r"\bArm Height\s*[:\-]\s*(.+)",                 "Arm Height"),
    (r"\bCOM\s*[:\-]\s*(.+)",                        "COM"),
    (r"\bCOL\s*[:\-]\s*(.+)",                        "COL"),
    (r"\bFabric\s*[:\-]\s*(.+)",                     "Fabric"),
    (r"\bUpholstery\s*[:\-]\s*(.+)",                 "Fabric"),
    (r"\bWattage\s*[:\-]\s*(.+)",                    "Wattage"),
    (r"\bSocket\s*[:\-]\s*(.+)",                     "Socket"),
    (r"\bColor Temperature\s*[:\-]\s*(.+)",          "Color Temperature"),
    (r"\bChain Length\s*[:\-]\s*(.+)",               "Chain Length"),
    (r"\bCanopy\s*[:\-]\s*(.+)",                     "Canopy"),
    (r"\bShade\s*[:\-]\s*(.+)",                      "Shade Details"),
]


def scrape_product(url: str) -> list[dict]:
    """
    Scrape a Palecek product detail page via requests.
    Returns a single-element list (no purchasable variants on this site).
    """
    row: dict = {"Source": url}

    soup = _get(url)
    if soup is None:
        return [row]

    page_text = soup.get_text(separator="\n")

    # ── 1. SKU from URL ───────────────────────────────────────────────────
    # URL: /{slug}/{sku}-1822/iteminformation.aspx
    url_parts = urlparse(url).path.strip("/").split("/")
    for part in url_parts:
        # Remove store suffix "-1822"
        m = re.match(r"^(.+?)-1822$", part)
        if m:
            row["SKU"] = m.group(1)
            break

    # ── 2. Product Name ───────────────────────────────────────────────────
    h1 = soup.find("h1")
    if h1:
        row["Product Name"] = clean_text(h1.get_text())

    # Fallback: "SKU: {sku}" line for name from meta/title
    if not row.get("Product Name"):
        title = soup.find("title")
        if title:
            row["Product Name"] = clean_text(title.get_text().split("|")[0])

    # ── 3. SKU fallback from page text ────────────────────────────────────
    if not row.get("SKU"):
        m = re.search(r"\bSKU\s*[:\-]\s*([\w-]+)", page_text, re.I)
        if m:
            row["SKU"] = m.group(1).strip()

    # ── 4. Image URL ──────────────────────────────────────────────────────
    img = soup.find("img", src=re.compile(r"imgix\.net", re.I))
    if not img:
        img = soup.find("img", attrs={"data-src": re.compile(r"imgix\.net", re.I)})
    if img:
        src = img.get("data-src") or img.get("src") or ""
        if src.startswith("//"):
            src = "https:" + src
        # Upgrade to high-res
        src = re.sub(r"[?&]w=\d+", "", src)
        src = re.sub(r"[?&]h=\d+", "", src)
        row["Image URL"] = src + "?w=1200" if "?" not in src else src + "&w=1200"

    # Construct from SKU if no image found
    if not row.get("Image URL") and row.get("SKU"):
        sku = row["SKU"]
        row["Image URL"] = f"https://images2.imgix.net/p4dbimg/1822/images/{sku}a.jpg?w=1200"

    # ── 5. Tearsheet ──────────────────────────────────────────────────────
    ts = soup.find("a", href=re.compile(r"printtearsheet|tearsheet", re.I))
    if ts and ts.get("href"):
        href = ts["href"]
        row["Tearsheet Link"] = href if href.startswith("http") else urljoin(BASE_URL, href)

    # ── 6. Description ────────────────────────────────────────────────────
    # Look for main product description paragraph (longest non-trivial paragraph)
    best_desc = ""
    for p in soup.find_all("p"):
        txt = clean_text(p.get_text())
        if len(txt) > len(best_desc) and len(txt) > 30:
            best_desc = txt
    if best_desc:
        row["Description"] = best_desc[:1500]

    # ── 7. Structured spec fields via regex on page text ──────────────────
    for pattern, field in _FIELD_PATTERNS:
        if row.get(field):
            continue
        m = re.search(pattern, page_text, re.I | re.MULTILINE)
        if m:
            val = clean_text(m.group(1)).strip().rstrip(".")
            if val and len(val) < 200:
                row[field] = val

    # ── 8. Parse Dimensions string into sub-fields ────────────────────────
    if row.get("Dimensions") and not row.get("Width"):
        parsed = parse_dimensions(row["Dimensions"])
        for k, v in parsed.items():
            row.setdefault(k, v)

    # ── 9. Product Family Id ──────────────────────────────────────────────
    if not row.get("Product Family Id") and row.get("SKU"):
        row["Product Family Id"] = row["SKU"].split("-")[0]
    if not row.get("Product Family Id") and row.get("Product Name"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    return [row]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main() -> None:
    info       = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer     = ExcelWriter(OUTPUT_PATH, info["vendor_name"])
    categories = info["categories"]

    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    print(f"\n[Scraper] Vendor  : {info['vendor_name']}")
    print(f"[Scraper] Mode    : {'TEST' if TEST_MODE else 'FULL'}")
    print(f"[Scraper] Output  : {OUTPUT_PATH}")

    for cat in categories:
        if not cat["links"]:
            continue

        writer.add_sheet(
            cat["name"],
            cat["links"][0],
            studio_columns=cat["studio_columns"],
        )

        # Collect all product URLs via requests (no browser)
        seen_urls: set[str] = set()
        all_urls:  list[str] = []

        for listing_url in cat["links"]:
            max_p = TEST_MAX_PRODUCTS if TEST_MODE else 0
            for u in get_product_links(listing_url, max_products=max_p):
                if u not in seen_urls:
                    seen_urls.add(u)
                    all_urls.append(u)

        if TEST_MODE:
            all_urls = all_urls[:TEST_MAX_PRODUCTS]

        print(f"\n[Category] {cat['name']}: {len(all_urls)} products")

        for idx, url in enumerate(all_urls, 1):
            try:
                rows = scrape_product(url)
                for row in rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    row["Manufacturer"] = info["vendor_name"]
                    writer.write_row(row, category_name=cat["name"])
                print(f"  [{idx}] {url.split('/')[-2]}")
            except Exception as e:
                print(f"  [ERROR] {url}: {e}")
            time.sleep(0.8)

        time.sleep(1.0)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
