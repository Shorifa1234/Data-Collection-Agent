# -*- coding: utf-8 -*-
import asyncio, json, os, sys, re
from pathlib import Path
from urllib.parse import urlparse

PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from base_scraper import (
    PlaywrightBrowser, ExcelWriter,
    async_polite_delay, clean_text, sentence_case,
    generate_sku, extract_family_id,
)

VENDOR_NAME         = os.environ.get("VENDOR_NAME", "Helene Aumont")
HEADLESS            = os.environ.get("HEADLESS", "true").lower() != "false"
TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "2"))
TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))
BASE_URL            = "https://heleneaumont.com"
OUTPUT_PATH         = Path(os.environ.get("OUTPUT_PATH",
    str(PROJECT_ROOT / VENDOR_NAME / "Data" / f"{VENDOR_NAME}.xlsx")))

# Labels to skip when building description
_DESC_SKIP = re.compile(
    r'^(dimensions|lead times|fabrication|hardwood|bronze|leather|color cords?|'
    r'f\.o\.b\.|copyright|\d+[-–]\d+\s*weeks?)',
    re.IGNORECASE,
)


def _extract_sku_from_pdf(pdf_url):
    # e.g. /s/HA3042-16small.pdf -> HA304  /s/HA-107-2-16-small.pdf -> HA-107
    if not pdf_url:
        return None
    filename = pdf_url.rstrip('/').split('/')[-1]
    filename = re.sub(r'\.pdf$', '', filename, flags=re.IGNORECASE)
    filename = re.sub(r'-small$',   '', filename, flags=re.IGNORECASE)
    filename = re.sub(r'v\d+$',     '', filename)
    filename = re.sub(r'-?2-16$',   '', filename)
    return filename.strip() or None


def _parse_dim_lines(raw):
    # Normalise fraction characters and quote variants to ASCII
    raw = (raw
           .replace('½', '.5')
           .replace('¼', '.25')
           .replace('¾', '.75')
           .replace('″', '"')
           .replace('”', '"')
           .replace('’', '"')
           .replace('′', '"'))
    # Convert "N M/K" mixed fractions  e.g. 18 1/2 -> 18.5
    raw = re.sub(
        r'(\d+)\s+(\d+)/(\d+)',
        lambda m: str(round(int(m.group(1)) + int(m.group(2)) / int(m.group(3)), 2)),
        raw,
    )

    # Match: number + optional-quote + specific dimension keyword
    # Uses finditer on the whole text so inline and newline-separated both work
    _PAT = re.compile(
        r'([\d]+(?:\.\d+)?)\s*["\'`]?\s*'
        r'((?:(?:front|back)\s+)?seat\s+(?:width|depth|height)'
        r'|arm\s+height'
        r'|back\s+height'
        r'|(?:seat\s+)?(?:width|depth|height|length|diameter))',
        re.IGNORECASE,
    )

    result = {}
    dim_parts = []

    for m in _PAT.finditer(raw):
        val   = m.group(1)
        label = m.group(2).strip().lower()
        dim_parts.append(val + ' ' + label)

        if 'seat' in label and 'width'  in label:
            result.setdefault('Width', val)
            result['Seat Width'] = val
        elif 'seat' in label and 'depth'  in label:
            result.setdefault('Depth', val)
            result['Seat Depth'] = val
        elif 'seat' in label and 'height' in label:
            result.setdefault('Seat Height', val)
        elif 'front' in label and 'seat'  in label:
            result.setdefault('Seat Height', val)
        elif 'arm'   in label and 'height' in label:
            result['Arm Height'] = val
        elif 'back'  in label and 'height' in label:
            result.setdefault('Height', val)
        elif 'diameter' in label:
            result['Diameter'] = val
        elif 'length'   in label:
            result['Length'] = val
        elif 'width'    in label:
            result['Width'] = val
        elif 'depth'    in label:
            result['Depth'] = val
        elif 'height'   in label:
            result['Height'] = val

    if dim_parts:
        result.setdefault('Dimensions', ', '.join(dim_parts))
    return result


# ---------------------------------------------------------------------------
# Listing page cache
# ---------------------------------------------------------------------------

async def load_listing_cache(page, anchor_ids):
    # Navigate to /shop-by-style once; section IDs are on DIV elements, not headings
    await page.goto(f"{BASE_URL}/shop-by-style",
                    timeout=60_000, wait_until="networkidle")
    await page.wait_for_timeout(2000)
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await page.wait_for_timeout(1500)

    # Click through all carousel right-arrows until each is hidden (end of slides)
    # Squarespace uses: button.user-items-list-carousel__arrow-button--right
    # Hidden (disabled) state: adds class --hidden
    await page.evaluate('''
        async () => {
            const SEL = "button.user-items-list-carousel__arrow-button--right:not(.user-items-list-carousel__arrow-button--hidden)";
            for (let round = 0; round < 100; round++) {
                const btns = Array.from(document.querySelectorAll(SEL));
                if (btns.length === 0) break;
                for (const btn of btns) btn.click();
                await new Promise(r => setTimeout(r, 400));
            }
        }
    ''')
    await page.wait_for_timeout(800)

    ids_json = json.dumps(anchor_ids)
    return await page.evaluate('''
        (knownIds) => {
            const BASE = "''' + BASE_URL + '''";
            const SKIP = /\\/(shop-by-style|browse-by-collection|blog|about|contact|follow|legal|index|showrooms|press|la-maison|cart|burger)/;
            const headers = knownIds
                .map(id => document.getElementById(id))
                .filter(Boolean);

            if (headers.length === 0) return {};

            const anchors = Array.from(document.querySelectorAll("a[href]"))
                .filter(a => {
                    const h = a.getAttribute("href") || "";
                    return h.startsWith("/") && !h.includes("#") &&
                           h.length > 1 && !SKIP.test(h);
                });

            const result = {};
            anchors.forEach(a => {
                const href = a.getAttribute("href");
                let best = null;
                for (const h of headers) {
                    if (h.compareDocumentPosition(a) & 4) {
                        if (!best) {
                            best = h;
                        } else if (best.compareDocumentPosition(h) & 4) {
                            best = h;
                        }
                    }
                }
                if (!best) return;
                const full = BASE + href;
                if (!result[best.id]) result[best.id] = [];
                if (!result[best.id].includes(full)) result[best.id].push(full);
            });
            return result;
        }
    ''', anchor_ids)


# ---------------------------------------------------------------------------
# Product page scraper
# ---------------------------------------------------------------------------

async def scrape_product(page, url):
    row = {"Source URL": url, "Manufacturer": VENDOR_NAME}

    await page.goto(url, timeout=45_000, wait_until="domcontentloaded")
    await page.wait_for_timeout(2000)

    # Product name (Squarespace custom pages: name is in main content, not site header)
    name = await page.evaluate('''
        () => {
            const selectors = ["main h1", "#page h1", "article h1", "main h2", "#page h2", "h1"];
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                if (el) {
                    const t = el.textContent.trim();
                    if (t.length > 1) return t;
                }
            }
            return "";
        }
    ''')
    if name:
        row["Product Name"]      = sentence_case(clean_text(name))
        row["Product Family Id"] = extract_family_id(row["Product Name"])

    # Main image — look inside main content, skip header/nav/footer images
    img = await page.evaluate('''
        () => {
            // Try product content containers first
            const containers = [
                document.querySelector("main"),
                document.querySelector("#page"),
                document.querySelector("article"),
                document.body,
            ];
            for (const container of containers) {
                if (!container) continue;
                for (const img of container.querySelectorAll("img")) {
                    // Skip images inside header, nav, footer
                    if (img.closest("header, nav, footer, #header, #menu")) continue;
                    const src = img.getAttribute("data-src") || img.getAttribute("src") || "";
                    if (src.includes("squarespace-cdn") && src.length > 80) return src;
                }
            }
            return null;
        }
    ''')
    if img:
        row["Image URL"] = img if img.startswith("http") else "https:" + img

    # Description
    desc_parts = await page.evaluate('''
        () => {
            const seen  = new Set();
            const texts = [];
            for (const el of document.querySelectorAll("p, li")) {
                const t = el.textContent.trim();
                if (t.length < 20 || seen.has(t)) continue;
                if (/^(copyright|f\\.o\\.b\\.|\\d+[-\\u2013]\\d+\\s*weeks?|dimensions|lead times|fabrication)/i.test(t)) continue;
                seen.add(t);
                texts.push(t);
                if (texts.length >= 4) break;
            }
            return texts;
        }
    ''')
    if desc_parts:
        row["Description"] = clean_text(" ".join(desc_parts))

    # Open all Squarespace accordion sections
    # Squarespace uses: <button class="accordion-item__click-target" aria-expanded="false">
    await page.evaluate('''
        async () => {
            const btns = document.querySelectorAll("button.accordion-item__click-target");
            for (const btn of btns) {
                if (btn.getAttribute("aria-expanded") !== "true") {
                    btn.click();
                    await new Promise(r => setTimeout(r, 300));
                }
            }
        }
    ''')
    await page.wait_for_timeout(700)

    # Read accordion content via aria-controls pairing
    accordions = await page.evaluate('''
        () => {
            const out = {};
            const btns = document.querySelectorAll("button.accordion-item__click-target");
            btns.forEach(btn => {
                const titleEl = btn.querySelector(".accordion-item__title");
                if (!titleEl) return;
                const key     = titleEl.textContent.trim().toUpperCase();
                const panelId = btn.getAttribute("aria-controls");
                const panel   = panelId ? document.getElementById(panelId) : null;
                out[key]      = panel ? panel.innerText.trim() : "";
            });
            return out;
        }
    ''')

    # Dimensions
    dim_key = next((k for k in accordions if "DIMENSION" in k), None)
    if dim_key:
        for k, v in _parse_dim_lines(accordions[dim_key]).items():
            row[k] = v

    # Lead Time
    lt_key = next((k for k in accordions if "LEAD" in k), None)
    if lt_key:
        row["Lead Time"] = clean_text(accordions[lt_key])

    # Fabrication
    fab_key = next((k for k in accordions if "FABRICAT" in k), None)
    if fab_key:
        row["Fabrication"] = clean_text(accordions[fab_key])

    # Spec sheet PDF link
    pdf_href = await page.evaluate('''
        (BASE) => {
            for (const a of document.querySelectorAll("a[href]")) {
                const h = (a.getAttribute("href") || "").toLowerCase();
                if (h.endsWith(".pdf")) {
                    const full = a.getAttribute("href");
                    return full.startsWith("http") ? full : BASE + full;
                }
            }
            return null;
        }
    ''', BASE_URL)
    if pdf_href:
        row["Tearsheet Link"] = pdf_href
        sku = _extract_sku_from_pdf(pdf_href)
        if sku:
            row["SKU"] = sku

    # Finish section labels (HARDWOOD FINISHES, BRONZE FINISHES, LEATHER COLOR WAYS, etc.)
    finish_labels = await page.evaluate('''
        () => {
            const text = document.body.innerText || "";
            const raw  = text.match(
                /(HARDWOOD|BRONZE|LEATHER[\\s\\w]*|COLOR[\\s\\w]*)\\s*FINISHES?:|LEATHER COLOR WAYS?:|COLOR CORDS?:/gi
            ) || [];
            return [...new Set(raw.map(m => m.replace(/:$/, "").trim()))];
        }
    ''')
    if finish_labels:
        row["Finish"] = ", ".join(finish_labels)

    return [row]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def main():
    info   = json.loads((Path(__file__).parent / "vendor_info.json").read_text())
    writer = ExcelWriter(OUTPUT_PATH, info["vendor_name"])

    cats = info["categories"]
    if TEST_MODE:
        cats = cats[:TEST_MAX_CATEGORIES]

    all_anchor_ids = []
    for cat in info["categories"]:
        for link_url in cat["links"]:
            frag = urlparse(link_url).fragment
            if frag and frag not in all_anchor_ids:
                all_anchor_ids.append(frag)

    async with PlaywrightBrowser(headless=HEADLESS) as page:
        print(f"[{VENDOR_NAME}] Loading shop-by-style listing page...")
        listing = await load_listing_cache(page, all_anchor_ids)
        print(f"[{VENDOR_NAME}] Sections found: {sorted(listing.keys())}")

        for cat in cats:
            if not cat["links"]:
                continue

            writer.add_sheet(
                cat["name"],
                cat["links"][0],
                studio_columns=cat["studio_columns"],
            )

            seen = set()
            product_urls = []
            for link_url in cat["links"]:
                anchor_id = urlparse(link_url).fragment
                for u in listing.get(anchor_id, []):
                    base_u = u.split("?")[0]
                    if base_u not in seen:
                        seen.add(base_u)
                        product_urls.append(base_u)

            if TEST_MODE:
                product_urls = product_urls[:TEST_MAX_PRODUCTS]

            print(f"[{VENDOR_NAME}] [{cat['name']}] scraping {len(product_urls)} products")

            idx = 1
            for url in product_urls:
                try:
                    for row in await scrape_product(page, url):
                        if not row.get("SKU"):
                            row["SKU"] = generate_sku(info["vendor_name"], cat["name"], idx)
                        if not row.get("Product Family Id") and row.get("Product Name"):
                            row["Product Family Id"] = extract_family_id(row["Product Name"])
                        writer.write_row(row, category_name=cat["name"])
                        idx += 1
                except Exception as e:
                    print(f"  [WARN] {url}: {e}")
                await async_polite_delay()

    writer.save()
    print(f"[{VENDOR_NAME}] Done -> {OUTPUT_PATH}")


if __name__ == "__main__":
    asyncio.run(main())
