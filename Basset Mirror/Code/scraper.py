"""
scraper.py — Basset Mirror
---------------------------
Platform: bassettmirror.com (Custom ColdFusion CMS)

Site structure:
  Listing  : /shop-products.cfm?cat={Category Name}  — server-rendered
             Product cards in <div class="product-card"> with <a class="product-anchor">
             All products shown per category (no pagination observed)
  Product  : /detail.cfm?id={id}/{sku}/{name}

Product page fields (trade-only — no price shown without login):
  Product Name   : <h1>
  SKU            : sub-header text e.g. "7086-LR-140  |  52x24x16"
  Dimensions     : also from sub-header (52x24x16 → W x D x H)
  Color/Finish   : from SPECIFICATIONS section
  Material       : from SPECIFICATIONS section
  Collection     : from SPECIFICATIONS section
  Description    : bullet list in DESCRIPTION section
  Tearsheet      : /detail.cfm?id=...&action=tearsheet
  Image          : from files.plytix.com CDN

Multi-link categories: tracker may have ?cat=Cocktail Tables + ?cat=Coffee Tables
The scraper loops all cat["links"] with deduplication.

Run directly:
    python scraper.py
Or via orchestrator:
    python orchestrator.py "Basset Mirror"
    python orchestrator.py "Basset Mirror" --test
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlencode, urlparse, parse_qs

import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    ExcelWriter,
    clean_text,
    sentence_case,
    generate_sku,
    extract_family_id,
    parse_dimensions,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Basset Mirror")
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))

BASE_URL = "https://www.bassettmirror.com"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": BASE_URL,
})

# Spec label (lowercase) → output field
_SPEC_MAP: dict[str, str] = {
    "color/finish":       "Finish",
    "finish":             "Finish",
    "color":              "Color",
    "color details":      "Color Details",
    "material":           "Materials",
    "style":              "Style",
    "collection":         "Collection",
    "designer":           "Designer",
    "table shape":        "Shape",
    "weight capacity":    "Weight Capacity",
    "shipping weight":    "Weight",
    "shipping method":    "Shipping Method",
    "bulb type":          "Bulb Type",
    "socket type":        "Socket",
    "socket":             "Socket",
    "wattage":            "Wattage",
    "shade type":         "Shade Details",
    "shade color":        "Shade Color",
    "shade shape":        "Shade Shape",
    "shade details":      "Shade Details",
    "base":               "Base",
    "base type":          "Base",
    "seat height":        "Seat Height",
    "seat depth":         "Seat Depth",
    "arm height":         "Arm Height",
    "com available":      "COM Available",
    "com":                "COM",
    "col":                "COL",
    "cot":                "COT",
    "content":            "Content",
    "overall dimensions": "Dimensions",
    "seat number":        "Seat Number",
}


# ---------------------------------------------------------------------------
# Listing page
# ---------------------------------------------------------------------------

def get_product_links(listing_url: str, max_products: int = 0) -> list[str]:
    """
    Collect all product detail links from a Bassett Mirror listing page.
    All products appear on one page; no pagination.
    Links follow: /detail.cfm?id={id}/{sku}/{name}
    """
    links: list[str] = []
    try:
        resp = SESSION.get(listing_url, timeout=25)
        resp.raise_for_status()
    except Exception as e:
        print(f"  [WARN] listing {listing_url}: {e}")
        return []

    soup = BeautifulSoup(resp.text, "html.parser")
    for card in soup.find_all("div", class_="product-card"):
        a = card.find("a", class_="product-anchor")
        if a and a.get("href"):
            href = a["href"]
            full = href if href.startswith("http") else urljoin(BASE_URL, href)
            links.append(full)
            if max_products and len(links) >= max_products:
                break

    return links


# ---------------------------------------------------------------------------
# Product detail page
# ---------------------------------------------------------------------------

def scrape_product(detail_url: str) -> list[dict]:
    """
    Scrape a Bassett Mirror product detail page.
    Returns a list with one dict (no variants on Bassett Mirror).
    """
    row: dict = {"Source": detail_url}

    try:
        resp = SESSION.get(detail_url, timeout=25)
        resp.raise_for_status()
    except Exception as e:
        print(f"    [WARN] {detail_url}: {e}")
        return [row]

    soup = BeautifulSoup(resp.text, "html.parser")

    # ── 1. Product Name ───────────────────────────────────────────────────
    h1 = soup.find("h1")
    if h1:
        row["Product Name"] = sentence_case(clean_text(h1.get_text()))

    # ── 2. SKU + Dimensions from sub-header ──────────────────────────────
    # Pattern: "7086-LR-140  |  52x24x16"  or "7086-LR-140"
    page_text = soup.get_text(separator="\n")
    sku_dim_match = re.search(
        r"([A-Z0-9]{2,10}[-–][A-Z0-9\-]+(?:[-–][A-Z0-9]+)?)"  # SKU
        r"\s*[|｜]\s*"                                           # separator
        r"([\d.x×\s]+)",                                         # dimensions
        page_text
    )
    if sku_dim_match:
        row["SKU"] = sku_dim_match.group(1).strip()
        dim_raw = sku_dim_match.group(2).strip()
        # Bassett mirror dims are WxDxH (integers): "52x24x16"
        dim_parts = re.split(r"[x×]", dim_raw)
        if len(dim_parts) == 3:
            w, d, h = [p.strip() for p in dim_parts]
            row["Dimensions"] = f"{w}W x {d}D x {h}H"
            row["Width"]  = w
            row["Depth"]  = d
            row["Height"] = h
        elif len(dim_parts) == 2:
            w, h = [p.strip() for p in dim_parts]
            row["Dimensions"] = f"{w}W x {h}H"
            row["Width"]  = w
            row["Height"] = h
        elif len(dim_parts) == 1 and dim_raw:
            row["Dimensions"] = dim_raw
    else:
        # Try SKU only (no dimensions in sub-header)
        sku_only = re.search(r"\b([A-Z0-9]{2,10}[-–][A-Z0-9\-]+)\b", page_text)
        if sku_only:
            row["SKU"] = sku_only.group(1)

    # ── 3. Image URL ─────────────────────────────────────────────────────
    for img in soup.find_all("img"):
        src = img.get("src", "")
        if "plytix" in src or "bassettmirror" in src:
            row["Image URL"] = src
            break

    # ── 4. Specifications section ─────────────────────────────────────────
    spec_section = None
    for div in soup.find_all("div"):
        txt = div.get_text()
        if "Color/Finish" in txt or ("Color" in txt and "Material" in txt and len(txt) < 1000):
            spec_section = div
            break

    if spec_section:
        spec_text = spec_section.get_text(separator="\n")
        # Parse "Key:Value" or "Key:\nValue" patterns
        for line in spec_text.split("\n"):
            line = line.strip()
            if ":" in line:
                parts = line.split(":", 1)
                label = parts[0].strip().lower()
                value = parts[1].strip()
                field = _SPEC_MAP.get(label)
                if field and value and len(value) < 200:
                    row.setdefault(field, value)

    # ── 5. Description ───────────────────────────────────────────────────
    # DESCRIPTION section contains bullet list items
    desc_lines = []
    desc_header = soup.find(string=re.compile(r"^Description$", re.I))
    if desc_header:
        parent = desc_header.find_parent(["div", "section"])
        if parent:
            items = parent.find_all("li")
            if items:
                desc_lines = [clean_text(li.get_text()) for li in items if clean_text(li.get_text())]
            else:
                # plain text paragraphs
                for p in parent.find_all("p"):
                    txt = clean_text(p.get_text())
                    if txt and len(txt) > 10:
                        desc_lines.append(txt)
    if desc_lines:
        row["Description"] = " | ".join(desc_lines)

    # ── 6. Tearsheet ─────────────────────────────────────────────────────
    # Construct from detail URL: add &action=tearsheet
    if "detail.cfm" in detail_url:
        row["Tearsheet Link"] = detail_url + "&action=tearsheet"
    else:
        ts_el = soup.find("a", string=re.compile(r"tearsheet|download", re.I))
        if ts_el and ts_el.get("href"):
            href = ts_el["href"]
            row["Tearsheet Link"] = href if href.startswith("http") else urljoin(BASE_URL, href)

    # ── 7. Product Family Id ─────────────────────────────────────────────
    if not row.get("Product Family Id") and row.get("Product Name"):
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    row["Manufacturer"] = VENDOR_NAME
    return [row]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
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

        print(f"\n[Category] {cat['name']} — collecting links…")

        seen_urls: set[str] = set()
        all_urls: list[str] = []
        max_p = TEST_MAX_PRODUCTS if TEST_MODE else 0

        for listing_url in cat["links"]:
            for u in get_product_links(listing_url, max_products=max_p):
                if u not in seen_urls:
                    seen_urls.add(u)
                    all_urls.append(u)

        if TEST_MODE:
            all_urls = all_urls[:TEST_MAX_PRODUCTS]
        print(f"  {len(all_urls)} products")

        global_idx = 1
        for url in all_urls:
            try:
                rows = scrape_product(url)
                for row in rows:
                    if not row.get("SKU"):
                        row["SKU"] = generate_sku(info["vendor_name"], cat["name"], global_idx)
                    if not row.get("Product Family Id") and row.get("Product Name"):
                        row["Product Family Id"] = extract_family_id(row["Product Name"])
                    writer.write_row(row, category_name=cat["name"])
                    global_idx += 1
                name = url.rstrip("/").split("/")[-1][:50]
                print(f"  [{global_idx - 1}] {name}")
            except Exception as e:
                print(f"  [ERROR] {url}: {e}")
            time.sleep(0.5)

    writer.save()
    print(f"\n[Done] {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
