import asyncio
import html as html_module
import json
import os
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser,
    ExcelWriter,
    async_polite_delay,
    clean_text,
    sentence_case,
    generate_sku,
    extract_family_id,
    safe_float,
)

VENDOR_NAME = os.environ.get("VENDOR_NAME", "EQ3")
HEADLESS    = os.environ.get("HEADLESS", "true").lower() != "false"
OUTPUT_PATH = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

BASE_URL = "https://www.eq3.com"

TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _strip_html(html_str: str) -> str:
    """Strip HTML tags and decode entities into plain text."""
    text = re.sub(r"<[^>]+>", " ", html_str)
    text = html_module.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def _best_image(galleries: dict) -> str:
    """Return the highest-res image from a galleries dict (prefers COVE)."""
    for key in ("COVE", "LIFESTYLE", "SWATCH"):
        imgs = galleries.get(key) or []
        if imgs:
            img = imgs[0]
            return img.get("highResUrl") or img.get("url", "")
    # Any gallery as fallback
    for imgs in galleries.values():
        if imgs and isinstance(imgs, list):
            img = imgs[0]
            return img.get("highResUrl") or img.get("url", "")
    return ""


def _get_product_query_data(next_data: dict) -> dict:
    """
    Navigate __NEXT_DATA__ to the product query data.
    EQ3 structure:
      props.pageProps.dehydratedState.queries[0].state.data
        → { definition, instances, requestedInstance, instanceId }
    """
    try:
        queries = (
            next_data["props"]["pageProps"]["dehydratedState"]["queries"]
        )
    except (KeyError, TypeError):
        return {}

    # queries[0] has product data; confirm by checking for 'definition' key
    for q in queries:
        data = q.get("state", {}).get("data", {})
        if isinstance(data, dict) and "definition" in data:
            return data
    return {}


# ---------------------------------------------------------------------------
# Parse product data from __NEXT_DATA__
# ---------------------------------------------------------------------------

def parse_eq3_product(next_data: dict, url: str) -> list[dict]:
    """
    Parse EQ3 __NEXT_DATA__ JSON from a product page.
    Returns one dict per product instance (variant: finish / fabric / size).
    """
    q_data = _get_product_query_data(next_data)
    if not q_data:
        return [{"Source": url}]

    defn      = q_data.get("definition", {})
    instances = q_data.get("instances", [])

    if not defn:
        return [{"Source": url}]

    product_name = clean_text(defn.get("name", ""))

    # --- Description from textBlocks.DETAILS (HTML) ---
    text_blocks  = defn.get("textBlocks") or {}
    details_html = text_blocks.get("DETAILS", "")
    description  = _strip_html(details_html) if details_html else clean_text(
        defn.get("seoDescription", "") or defn.get("metaDescription", "")
    )

    # --- Collection / family ---
    collection = clean_text(defn.get("familyTitle", "") or "")

    # --- Shared dimensions from definition ---
    def_dims = defn.get("dimensions") or []

    # --- Product-level images (definition.images.galleries) ---
    def_galleries = (defn.get("images") or {}).get("galleries", {})
    main_image    = _best_image(def_galleries)

    # --- Fall back to single-instance if no instances list ---
    if not instances:
        instances = [{"selectedProductDefinition": defn, "sellingPrice": None, "regularPrice": None}]

    rows = []
    for inst in instances:
        row: dict = {
            "Source":            url,
            "Manufacturer":      VENDOR_NAME,
            "Product Name":      product_name,
            "Product Family Id": extract_family_id(product_name),
        }

        if description:
            row["Description"] = description
        if collection:
            row["Collection"] = collection

        # Price
        price_val = inst.get("sellingPrice") or inst.get("regularPrice")
        if price_val is not None:
            row["Price"] = safe_float(str(price_val))

        # Weight
        mass = inst.get("massPounds")
        if mass is not None:
            row["Weight"] = f"{safe_float(str(mass))} lbs"

        # Availability
        avail = inst.get("availability", "")
        if avail:
            row["Availability"] = sentence_case(avail.replace("_", " "))

        # selectedProductDefinition — variant-specific sub-definition
        spd = inst.get("selectedProductDefinition") or defn

        # --- SKU from spd.components[i].item.partnerItem.id ---
        sku = ""
        for comp in spd.get("components", []):
            item = comp.get("item") or {}
            pi   = item.get("partnerItem") or {}
            variant_sku = str(pi.get("id", "")).strip()
            if variant_sku:
                sku = variant_sku
                break
        if not sku:
            sku = spd.get("viewModelSku", "")
        if sku:
            row["SKU"] = sku

        # --- Component selections (Finish, Fabric, Size, etc.) ---
        for comp in spd.get("components", []):
            comp_name = comp.get("name", "")
            item      = comp.get("item") or {}
            alias     = item.get("alias", "")
            if comp_name and alias:
                row[comp_name] = alias

        # --- Dimensions (spd first, then shared definition dims) ---
        inst_dims = spd.get("dimensions") or def_dims
        if inst_dims:
            d = inst_dims[0] if isinstance(inst_dims, list) else inst_dims
            if isinstance(d, dict):
                w      = safe_float(str(d["widthInches"]))    if d.get("widthInches")    else None
                h      = safe_float(str(d["heightInches"]))   if d.get("heightInches")   else None
                dep    = safe_float(str(d["depthInches"]))    if d.get("depthInches")    else None
                diam   = safe_float(str(d["diameterInches"])) if d.get("diameterInches") else None
                length = safe_float(str(d["lengthInches"]))   if d.get("lengthInches")   else None

                parts = []
                if w:      row["Width"]    = str(w);      parts.append(f"W {w}")
                if dep:    row["Depth"]    = str(dep);    parts.append(f"D {dep}")
                if h:      row["Height"]   = str(h);      parts.append(f"H {h}")
                if diam:   row["Diameter"] = str(diam);   parts.append(f"Diam {diam}")
                if length: row["Length"]   = str(length); parts.append(f"L {length}")
                if parts:
                    row["Dimensions"] = " x ".join(parts)

        # --- Image URL (variant gallery → instance thumbnail → product gallery) ---
        spd_galleries = (spd.get("images") or {}).get("galleries", {})
        spd_image     = _best_image(spd_galleries)
        thumbnail_url = (inst.get("thumbnail") or {}).get("url", "")
        row["Image URL"] = spd_image or thumbnail_url or main_image

        rows.append(row)

    return rows if rows else [{"Source": url}]


# ---------------------------------------------------------------------------
# Listing page — collect all product URLs
# ---------------------------------------------------------------------------

def _build_product_url(prod: dict) -> str | None:
    """Construct a direct instance URL from a listing product entry."""
    slugs = prod.get("slugs") or {}
    pid   = prod.get("id", "")
    cat   = slugs.get("category", "")
    sub   = slugs.get("subcategory", "")
    pl    = slugs.get("productline", "")
    slug  = slugs.get("product", "")
    if pid and cat and sub and pl and slug:
        return f"{BASE_URL}/us/en/product/{pid}/{cat}/{sub}/{pl}/{slug}"
    return None


def _extract_nd_products(nd_json: dict) -> tuple[list[str], int]:
    """
    Parse __NEXT_DATA__ from a listing page.
    Returns (list_of_product_urls, total_server_count).
    Skips isCustomMade products.
    """
    try:
        queries = nd_json["props"]["pageProps"]["dehydratedState"]["queries"]
    except (KeyError, TypeError):
        return [], 0

    for q in queries:
        data = q.get("state", {}).get("data", {})
        if not isinstance(data, dict) or "pages" not in data:
            continue
        pages  = data["pages"] or []
        total  = (pages[0].get("data", {}) or {}).get("count", 0) if pages else 0
        urls: list[str] = []
        seen: set[str]  = set()
        for page in pages:
            for prod in (page.get("data", {}) or {}).get("products", []):
                if prod.get("isCustomMade"):
                    continue
                url = _build_product_url(prod)
                if url and url not in seen:
                    seen.add(url)
                    urls.append(url)
        return urls, total

    return [], 0


async def get_product_links(page, listing_url: str) -> list[str]:
    """
    Collect all individual product instance URLs for a category listing.

    Strategy:
    1. Parse __NEXT_DATA__ on initial load — gives clean instanceId URLs with
       no cross-category contamination and no custom/configurator products.
    2. If total server count > initial load, click 'Load More' and collect
       additional DOM links filtered by the listing's category path.
    """
    from urllib.parse import urlparse

    await page.goto(listing_url, wait_until="domcontentloaded", timeout=45_000)
    await page.wait_for_timeout(3_000)

    # Step 1 — __NEXT_DATA__ products (clean, definitive)
    nd_text: str = await page.evaluate(
        '() => document.getElementById("__NEXT_DATA__")?.textContent || ""'
    )
    nd_json   = json.loads(nd_text) if nd_text.strip() else {}
    nd_urls, total_count = _extract_nd_products(nd_json)

    seen: set[str]   = set(nd_urls)
    all_urls: list[str] = list(nd_urls)

    # Step 2 — if more products remain, click 'Load More' and collect DOM links
    if total_count > len(nd_urls):
        # Category path for DOM link filtering (prevents cross-category contamination)
        cat_path = urlparse(listing_url).path.split("/category/", 1)[-1]
        # e.g. "bedroom/storage/nightstands" → filter for /{cat_path}/ in URL

        prev_dom_count = 0
        for _ in range(60):
            clicked = await page.evaluate("""
                () => {
                    const btn = Array.from(document.querySelectorAll('button'))
                        .find(b => /load\\s*more/i.test(b.textContent) && !b.disabled);
                    if (btn) { btn.click(); return true; }
                    return false;
                }
            """)
            if not clicked:
                break
            await page.wait_for_timeout(2_500)
            dom_count = await page.evaluate(
                '() => document.querySelectorAll(\'a[href*="/us/en/product/"]\').length'
            )
            if dom_count <= prev_dom_count:
                break
            prev_dom_count = dom_count

        # Collect DOM links filtered by this category's path
        dom_urls: list[str] = await page.evaluate(f"""
            () => {{
                const catPath = "/{cat_path}/";
                const seen = new Set();
                const result = [];
                document.querySelectorAll('a[href*="/us/en/product/"]').forEach(a => {{
                    const href = a.href.split('?')[0].split('#')[0];
                    if (href.includes(catPath) && !seen.has(href)) {{
                        seen.add(href);
                        result.push(href);
                    }}
                }});
                return result;
            }}
        """)
        for url in dom_urls:
            if url not in seen:
                seen.add(url)
                all_urls.append(url)

    return all_urls


# ---------------------------------------------------------------------------
# Product page scraper
# ---------------------------------------------------------------------------

async def scrape_product(page, url: str) -> list[dict]:
    """
    Scrape a single EQ3 product page via __NEXT_DATA__ JSON.
    Returns one row per variant instance.
    """
    await page.goto(url, wait_until="domcontentloaded", timeout=45_000)
    await page.wait_for_timeout(1_500)

    try:
        next_data_text: str = await page.evaluate(
            '() => document.getElementById("__NEXT_DATA__")?.textContent || ""'
        )
        if not next_data_text.strip():
            print(f"    No __NEXT_DATA__ at {url}")
            return [{"Source": url}]

        next_data = json.loads(next_data_text)
        return parse_eq3_product(next_data, url)

    except json.JSONDecodeError as e:
        print(f"    JSON parse error for {url}: {e}")
        return [{"Source": url}]
    except Exception as e:
        print(f"    ERROR scraping {url}: {e}")
        return [{"Source": url}]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    categories = info["categories"]
    if TEST_MODE:
        categories = categories[:TEST_MAX_CATEGORIES]
        print(f"[TEST: max {TEST_MAX_CATEGORIES} categories, {TEST_MAX_PRODUCTS} products each]")

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        for cat in categories:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            # Collect all product URLs across all listing links for this category
            seen_urls: set[str] = set()
            all_product_urls: list[str] = []
            for listing_url in cat["links"]:
                for u in await get_product_links(page, listing_url):
                    if u not in seen_urls:
                        seen_urls.add(u)
                        all_product_urls.append(u)

            if TEST_MODE:
                all_product_urls = all_product_urls[:TEST_MAX_PRODUCTS]

            print(f"  {cat['name']}: {len(all_product_urls)} products found")

            global_idx = 1
            for url in all_product_urls:
                try:
                    variant_rows = await scrape_product(page, url)
                    for variant in variant_rows:
                        # Skip rows with no product data (failed parse / product-line pages)
                        if not variant.get("Product Name"):
                            continue
                        if not variant.get("SKU"):
                            variant["SKU"] = generate_sku(
                                info["vendor_name"], cat["name"], global_idx
                            )
                        if not variant.get("Product Family Id") and variant.get("Product Name"):
                            variant["Product Family Id"] = extract_family_id(
                                variant["Product Name"]
                            )
                        writer.write_row(variant, category_name=cat["name"])
                        global_idx += 1
                except Exception as e:
                    print(f"    ERROR on {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
