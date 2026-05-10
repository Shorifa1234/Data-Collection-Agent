# AVS Agent — Changelog

---

## V1.7 — 2026-04-19

### Curry scraper VPS fix + base_scraper empty-workbook guard

**Curry (`Curry\Code\scraper.py`) — full rewrite**
- **Root cause**: listing pages are 100% JS-rendered; Playwright could not extract product links even locally
- **Fix**: replaced Playwright listing scrape with `requests` + XML sitemap parse (`/sitemap.xml`); all ~1000+ product URLs fetched in one HTTP request, then filtered per category by path prefix (e.g. `/c/furniture/chests-nightstands/`)
- Replaced `PlaywrightBrowser` (blocks fonts) with custom `create_page()` that only blocks `media` — fonts/images now load, enabling Cloudflare JS challenges to complete on VPS/datacenter IPs
- Added `_is_cf_blocked()` and `safe_goto()` with 2 retries + 8s Cloudflare wait (same pattern as Palecek fix in V1.5)
- Added full stealth headers: `Sec-Ch-Ua`, `Sec-Fetch-*`, realistic `navigator.plugins` / `window.chrome` init script
- Playwright is now used only for product detail page scraping (JSON-LD + CSS spec table)

**Palecek (`Palecek\Code\scraper.py`) — full rewrite (requests only)**
- **Root cause**: site is fully server-rendered static HTML — Playwright was never needed; Cloudflare was blocking the headless browser on VPS
- **Fix**: complete rewrite using `requests` + `BeautifulSoup` only — no browser, no Cloudflare issues
- Listing pages: HTTP GET → parse `a[href*="iteminformation.aspx"]` links; pagination via next-page link detection
- Product pages: HTTP GET → regex extraction of fields from page text (`Overall Dimensions`, `SKU`, `Finish`, `Materials`, etc.)
- Image URL: from `img[src*="imgix.net"]` or constructed from SKU (`images2.imgix.net/p4dbimg/1822/images/{sku}a.jpg`)
- Tearsheet: from `a[href*="printtearsheet.aspx"]`
- SKU: extracted from URL by stripping `-1822` store suffix from path segment

**Alfonso Marina (`Alfonso Marina\Code\scraper.py`) — VPS fix**
- Same root cause as Curry: `PlaywrightBrowser` blocked fonts → Cloudflare JS challenge failed on VPS → 0 products → empty workbook crash
- Replaced `PlaywrightBrowser` with `create_page()` (media-only block) + stealth headers + `_is_cf_blocked()` / `safe_goto()` with retries
- `get_product_links` simplified to use `safe_goto` (removed manual retry loop)

**`base_scraper.py` — `ExcelWriter.save()`**
- Guard against saving an empty workbook (all sheets had 0 rows): now prints a warning and returns instead of crashing with `openpyxl.utils.exceptions.InvalidFileException`

---

## V1.6 — 2026-04-18

### New vendor scrapers: Sunpan, Villa & House, Basset Mirror, Alfonso Marina

| Vendor | Tracker name | Platform | Approach |
|---|---|---|---|
| Sunpan | `Sunpan` | Shopify (sunpan.com) | `requests` + Shopify `/products.json` API for listing; `requests` + BeautifulSoup for product detail spec table (Material, Dimensions, Designer, etc.); one row per variant (option1 = finish/color) |
| Villa & House | `Villa & House` | BigCommerce (vandh.com) | `PlaywrightBrowser`; article card links (`article.card a.card-figure__link`); `?page=N` pagination; SKU from `SKU: {value}` text, dimensions from body text pattern (`48W x 24D x 17.5H`); no price (trade-only) |
| Basset Mirror | `Basset Mirror` | Custom ColdFusion CMS (bassettmirror.com) | `requests` + BeautifulSoup; product cards `.product-card a.product-anchor`; detail at `/detail.cfm?id={id}/{sku}/{name}`; SKU + dimensions from sub-header (`7086-LR-140 \| 52x24x16`); specs parsed as key:value text; tearsheet via `&action=tearsheet` query param |
| Alfonso Marina | `Alfonso Marina` | WooCommerce/WordPress (alfonsomarina.com) | `PlaywrightBrowser` for listing (JS-rendered); `requests` for product detail; SKU extracted from product image filename (`509-318-02_PRODUCT.webp`); dimensions from CM and IN blocks; finish options from `<select>` dropdown; no price (trade-only) |

---

## V1.5 — 2026-04-18

### New vendor scrapers: Parker Southern, Gabby, Curry & Company, Allied Maker, Four Hands
### Palecek scraper rewrite (VPS Cloudflare fix + full product pages)
### Regina Andrew multi-link bug fix
### Parallel runner: run 2+ vendors simultaneously on VPS

---

#### Palecek — scraper rewrite

**Problem:** Scraper showed 0 products on VPS but worked locally. Also only scraped page 1 (~30 products) from listing cards — no product detail pages, no Description/Price/Dimensions.

**Root cause:** `PlaywrightBrowser` blocks `{"image", "media", "font", "other"}` via route handler. Fonts are required for Cloudflare's JS challenge to pass. On datacenter/VPS IPs, Cloudflare is stricter and serves a challenge page with no `.ProductThumbnail` elements → 0 products.

**Changes:**
- Custom browser context that blocks **only `media`** (video/audio) — fonts and images now load, enabling Cloudflare challenge to complete on VPS IPs
- Added full stealth headers: `Sec-Ch-Ua`, `Sec-Fetch-*`, `Sec-Ch-Ua-Platform`, realistic `navigator.plugins` and `window.chrome` init script
- **Cloudflare detection**: checks page title for "Just a moment" / challenge element; waits 8 s and retries
- **Pagination**: WordPress `/page/N/` pattern — now scrapes all pages, not just page 1
- **Product detail pages** (Phase 2): visits every product URL for full data — Price, Description, Dimensions, Materials, Finish, Collection, Designer, Origin, Lead Time, Seat Height, Arm Height, COM/COL, Wattage, Socket, Chain Length, Tearsheet Link
- **Listing fallback**: if product URLs still can't be collected (CF still blocking), falls back to card extraction from listing page
- Tearsheet link constructed from robots.txt-confirmed pattern: `/product/{slug}/tearsheet/pdf`

---

#### New vendor scrapers (5)

| Vendor | Tracker name | Platform | Approach |
|---|---|---|---|
| Parker Southern | `Parker Southern` | Custom ASP.NET, public | `PlaywrightBrowser`; all products on one page per listing; no JSON-LD — regex parsing of DIMENSIONS / DESCRIPTION / SERIES / FINISH SHOWN sections; high-res image from `/downloadit/ps/style/{sku}`; handles Furniture + Textiles (Fabric, Leather, Trim) |
| Gabby | `Gabby` | Shopify (`gabriellawhite.com`) | `PlaywrightBrowser`; gabby.com → 301 → gabriellawhite.com (Playwright follows automatically); JSON-LD Product schema; SKU format `SCH-XXXXX`; uses list price; Shopify CDN image upgraded to `?width=2048`; spec table for dimensions |
| Curry & Company | `Curry` | EPiServer CMS, JS-rendered | `PlaywrightBrowser`; waits for JS product grid; product URL format confirmed from sitemap: `/c/{category}/{subcategory}/{sku}/`; `?page=N` pagination + load-more button fallback; JSON-LD + CSS spec table |
| Allied Maker | `Allied Maker` | NetSuite / custom CMS | `PlaywrightBrowser`; all products on one page; product URLs are `/{Slug}`; scrapes DIMENSIONS / LAMPING / BRIGHTNESS / METAL FINISHES / GLASS FINISHES sections via heading-adjacent text extraction |
| Four Hands | `Four Hands` | React SPA (login-required) | Custom `create_page()` (no font blocking); login via `FH_EMAIL` / `FH_PASSWORD` env vars with 6 selector fallbacks for email/password/submit; auto-detects product card selector; load-more + `?page=N` pagination; JSON-LD + CSS spec fallbacks; tearsheet from `/product/{slug}/tearsheet/pdf` |

> **Note:** Curry & Company is listed as **`Curry`** in the SD tracker. Use `python orchestrator.py "Curry"` not `"Curry & Company"`.

> **Note:** Four Hands requires trade account credentials. Set `FH_EMAIL` and `FH_PASSWORD` environment variables before running.

---

#### Regina Andrew — multi-link bug fix

**Problem:** `main()` called `get_product_links(page, cat["links"][0])` — only ever scraped the first listing URL, ignoring Link 2 / Link 3 entries.

**Fix:** Replaced single-link call with a loop over all `cat["links"]` entries with a `seen_urls` deduplication set, matching the standard template in `CLAUDE.md`.

---

#### Parallel runner — `run_parallel.py`

New top-level script for running multiple scrapers simultaneously on multi-core VPS.

- Launches each vendor as a separate `asyncio` subprocess calling `orchestrator.py`
- Streams output from all vendors interleaved in real time, prefixed with `[VendorName]`
- `--batch N` flag for processing large vendor lists in chunks
- Timing summary at the end (per vendor + wall time)
- Recommended for 8-core / 16 GB VPS: **2 vendors at a time** is safe for all scraper types

```bash
python run_parallel.py "Gabby" "Curry"
python run_parallel.py --batch 2 "Parker Southern" "Gabby" "Curry" "Allied Maker" "Regina Andrew"
python run_parallel.py "Gabby" "Curry" --test
```

---

## V1.4 — 2026-04-13

### New vendor scrapers: Kravet, Highland House, Remains, Porta Romana, Palecek, Kannoa, Century, Asthetic Decor, Bunny Williams Home (BWH)

Added 9 new vendor scrapers. All follow the standard structure (vendor_info.json + scraper.py, test-mode support, dynamic columns, run_log.txt on completion).

| Vendor | Platform | Approach |
|---|---|---|
| Remains | Shopify (remains.com) | `requests` + Shopify `/products.json` API |
| Porta Romana | Shopify (portaromana.com) | `requests` + Shopify `/products.json` API |
| Kannoa | Shopify (kannoa.com) | `requests` + Shopify `/products.json` API |
| Bunny Williams Home (BWH) | Shopify (bunnywilliamshome.com) | `requests` + Shopify `/products.json` API |
| Century | Shopify (shop.centuryfurniture.com) | `requests` + Shopify API + client-side `product_type` filter; skips the external-domain Table Lamps category |
| Asthetic Decor | WooCommerce (aestheticdecor.com) | `requests` + BeautifulSoup; handles shared listing pages + direct product URLs |
| Highland House | Custom ASP.NET (highlandhousefurniture.com) | `requests` + BeautifulSoup; `ShowItemDetail.aspx?SKU=` pattern; no pricing (trade-only) |
| Palecek | Custom ASP.NET / OmniVue (palecek.com) | Playwright; `?page=N` pagination; JS-rendered finish variants via `?finish=N` URLs |
| Kravet | Magento 2 + Algolia (kravet.com) | Playwright; waits for Algolia-rendered product grid; Magento 2 product detail selectors + JSON-LD |

**Also updated:** The Future Perfect scraper unchanged — already production-ready.

---

## V1.3 — 2026-04-13

### Run log — automatic timing file saved after every full run

**Problem:** Scrape run durations were only printed to the terminal and lost after the session ended.

**Changes:**

- **`orchestrator.py` — `save_run_log()`** — new function that appends a timing entry to `<Vendor>/Data/run_log.txt` after every scrape run (full and test).
  - Format: `[2026-04-13 14:30:00]  |  Mode: FULL  |  Scrape: 1m 06s  |  Session: 1m 06s`
  - Multi-vendor sessions also record the full vendor list: `|  Vendors: Wesley Hall, Vanguard Designs`
  - Entries are separated by a `+` line so the file stays readable across repeated runs.
  - File is appended to — previous entries are never overwritten.
- **`orchestrator.py` — `main()`** — calls `save_run_log()` for every vendor whose `scrape_time` was captured, after the summary is printed.

---

## V1.2 — 2026-04-09

### Auto-derive dimension sub-fields from Dimensions string (all vendors)

**Problem:** The `Dimensions` field was being stored as a raw string (e.g. `"16.00" L x 10.75" W x 5.00" H"`) without populating the individual columns (`Width`, `Height`, `Length`, etc.).

**Changes:**

- **`base_scraper.py` — `parse_dimensions()`** — extended to handle two additional formats:
  - **Number-before-label** (e.g. `16.00" W x 5.00" H`, `36" dia`) — used by Hennepin Made Tech Specs
  - **Number-directly-followed-by-label** (e.g. `30H x 18.5Dia`) — used by SkLO
  - Existing label-before-number format (`W 25" x H 12"`) is unchanged and takes priority via a `seen` set that prevents duplicate axis overwrites.

- **`base_scraper.py` — `ExcelWriter.write_row()`** — added automatic sub-field derivation: when a row contains `Dimensions` but is missing any of `Width`, `Height`, `Depth`, `Diameter`, `Length`, `parse_dimensions` is called and the missing fields are filled in automatically. This applies to **all vendors** globally without requiring changes to individual scrapers.

- **`Hennepin Made/Code/scraper.py`** — removed the now-redundant secondary `parse_dimensions` call in `scrape_product_page()` and the unused import.

---

## V1.1 — 2026-04-07

### Multi-link categories

- **`vendor_parser.py`** — already collected all `Link`, `Link 2`, `Link 3` … rows into
  `cat["links"]`. No change needed; the parser was already correct.
- **`CLAUDE.md` scraper template** — `main()` now iterates every URL in `cat["links"]`
  and uses a `seen_urls` set to deduplicate products before scraping. Previously the
  template only visited `cat["links"][0]` (the first link).
- **`scrape-vendor.md`** — STEP 4 multi-link requirement documented; scraper writers
  must implement the dedup loop.

### Multi-vendor orchestrator

- **`orchestrator.py`** — `vendor` argument changed from a single string to
  `vendors` (`nargs="+"`) so one invocation can process any number of vendors in sequence.
  - New `process_vendor()` helper encapsulates single-vendor logic.
  - Timing tracked per vendor: `scrape_time` (wall-clock seconds the scraper process ran).
  - Structured summary printed at the end of every run showing status + timing per vendor
    plus total session time.
- **`scrape-vendor.md`** — updated description, STEP 2 shows multi-vendor invocation,
  STEP 7/8 ask Claude to report timing per vendor and overall session time.

### Useful commands updated

- `CLAUDE.md` commands section shows multi-vendor orchestrator examples.

---

## V1.0 — 2026-04-02

**Initial release as AVS Agent**

Project rebranded from "Vendor Scraping Agent" to **AVS Agent** (Accord Tech Solutions).

### Changes

- **Renamed `Product URL` → `Source`** across all scrapers and `base_scraper.py`
- **Renamed `List Price` → `Price`** across all scrapers and `base_scraper.py`
- **Added `Manufacturer` column** — auto-populated with the vendor name in every row via `ExcelWriter.write_row()`
- **Improved image quality** — image URL is now always taken from the product detail page (high-res `data-zoom-image` / `data-src` attributes preferred over `src`); listing-page thumbnails are no longer used
- **Created `Log/` folder** — this changelog file introduced for tracking version history
- Updated `CLAUDE.md` to reflect new column names, project name, and image rules
- Updated `orchestrator.py` spec output to use new column names
- Updated all vendor scrapers: The Future Perfect, Arteriors, Regina Andrew, Visual Comfort
- **Auto-hide empty columns** — `ExcelWriter.save()` now drops any column that has no data across all rows in a sheet (mandatory fields always kept); previously blank columns were written to the Excel making sheets look cluttered

---
