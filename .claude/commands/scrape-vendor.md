---
description: Scrape ALL product data for a vendor. Always tests first (all cats, min 5 products each), then asks before full run. Usage: /scrape-vendor <Vendor Name> [Vendor Name 2 ...]
---

You are a senior Python developer and web-scraping expert.
The user wants to collect product data for: **$ARGUMENTS**

Parse the argument as a space-separated list of vendor names (quoted names may contain spaces).
Process each vendor in sequence using the steps below.

Columns are **DYNAMIC** — collect every field on the product page.
Do not fix or limit the columns. `ExcelWriter` handles ordering automatically.

**Always run a test first. Never go straight to a full run.**

---

## STEP 1 — Verify each vendor exists in the tracker

```bash
python vendor_parser.py "$ARGUMENTS"
```

If multiple vendors: run `vendor_parser.py` for each one.
If any returns `ERROR: Vendor ... not found`, run `python vendor_parser.py --list`,
tell the user the closest match, and stop processing that vendor (continue others).

---

## STEP 2 — Run the orchestrator (check for existing scraper)

For a single vendor:
```bash
python orchestrator.py "$ARGUMENTS" --test --headless true
```

For multiple vendors (run them together — orchestrator handles the loop):
```bash
python orchestrator.py "Vendor A" "Vendor B" --test --headless true
```

| Exit code | Meaning | Action |
|---|---|---|
| 0 | Test done for all vendors | Jump to STEP 5 (inspect test output) |
| 2 | One or more scrapers crashed | Read traceback, fix scraper.py, retry |
| 3 | One or more scrapers not yet written | Continue to STEP 3 for those vendors |

When exit code 3, the orchestrator prints a timing summary and lists which vendors
need code generation. Process each "needs_code" vendor through STEPS 3–4, then
re-run STEP 2 to execute the test.

---

## STEP 3 — Analyse the website (only for vendors with exit code 3)

Read `<Vendor Name>/Code/vendor_info.json` for categories, links, and `studio_columns`.

**Multi-link categories**: if `cat["links"]` has more than one URL, all are listing
pages for the same category — the scraper must visit every link and deduplicate products.

Pick the **first category with links**. Fetch the listing page:
```bash
python -c "
import requests
r = requests.get('PASTE_CATEGORY_URL_HERE',
    headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'},
    timeout=20)
print(r.text[:10000])
"
```

Identify:
- **Product card selector** — CSS class wrapping each product, `<a>` href pattern
- **Pagination** — `?paged=N`, `?page=N`, "Load More" button, infinite scroll?

Fetch **2-3 product detail pages** and identify every available field:
- Product Name, SKU, Price, Image URL
- Description
- Specifications / attributes (table, `<dl>`, or inline pipe-separated text)
- Every attribute: Materials, Dimensions, Designer, Maker, Collection,
  Lead Time, Origin, Production, Date, Base, Canopy, Seat Height, COM, Finish,
  Illumination, Socket, Wattage, Tariff Disclaimer, Tearsheet Link, and more
- **Any field shown on the page — capture it all, nothing is too minor**

---

## STEP 4 — Write the scraper

Write `<Vendor Name>/Code/scraper.py` following the template in CLAUDE.md.

Critical requirements:
- `sys.path.insert(0, PROJECT_ROOT)` then `from base_scraper import ...`
- Use `PlaywrightBrowser` (handles JS rendering)
- `writer.add_sheet(name, link, studio_columns=cat["studio_columns"])`
- In `scrape_product()`: collect **every** field as plain dict keys — no fixed list
- Use `parse_spec_block()` for pipe/colon spec text
- Use `parse_dimensions()` to split dimension strings into individual fields
- **Multi-link support (required)**: iterate ALL links in `cat["links"]`, collect
  product URLs from each, deduplicate with a `seen_urls` set before scraping:
  ```python
  seen_urls: set[str] = set()
  all_product_urls: list[str] = []
  for listing_url in cat["links"]:
      for u in await get_product_links(page, listing_url):
          if u not in seen_urls:
              seen_urls.add(u)
              all_product_urls.append(u)
  ```
- Full pagination — ALL pages, not just page 1
- Wrap each product in try/except — one failure must not abort the whole category
- Read env vars: `HEADLESS`, `OUTPUT_PATH`, `VENDOR_NAME`
- **TEST MODE SUPPORT** (required):
  ```python
  TEST_MODE           = os.environ.get("TEST_MODE", "false").lower() == "true"
  TEST_MAX_CATEGORIES = int(os.environ.get("TEST_MAX_CATEGORIES", "999"))  # all categories
  TEST_MAX_PRODUCTS   = int(os.environ.get("TEST_MAX_PRODUCTS", "5"))      # min 5 per category
  ```
  When `TEST_MODE=true`:
  - Scrape **ALL categories** (no category limit — `TEST_MAX_CATEGORIES` defaults to 999)
  - Limit to first `TEST_MAX_PRODUCTS` products per category (across ALL links combined) — minimum 5
  - Log clearly: `[TEST: max N products per category]`

**Timing**: record the wall-clock time from when you start writing `scraper.py` to
when you finish. Report this as "Code generation time: Xs" in your STEP 5 report.

---

## STEP 5 — Run the test

```bash
python orchestrator.py "$ARGUMENTS" --test --headless true
```

For multiple vendors, pass all names together so the orchestrator prints a combined
timing summary. The summary shows `scrape=Xs` per vendor.

If it fails, read the traceback, fix the issue, retry. Common fixes:

| Problem | Fix |
|---|---|
| Selector not found | Inspect page source, update CSS selector |
| Pagination broken | Log page URLs to confirm `?paged=N` pattern |
| Timeout | Increase `TIMEOUT_MS`, add `wait_for_load_state("networkidle")` |
| Anti-bot block | Add longer delays, add `wait_for_timeout(3000)` |
| Missing fields | Check if JS-rendered, try scrolling into view first |
| Duplicate products | Ensure `seen_urls` dedup is applied across all listing links |

---

## STEP 6 — Inspect the test output

```bash
python -c "
import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('\"$ARGUMENTS\"/Data/\"$ARGUMENTS\"_TEST.xlsx')
for sheet in wb.sheetnames:
    ws = wb[sheet]
    headers = [ws.cell(4,c).value for c in range(1,ws.max_column+1) if ws.cell(4,c).value]
    rows = ws.max_row - 4
    print(f'Sheet: {sheet} | {rows} products | {len(headers)} cols')
    print('  Cols:', headers)
    row5 = {headers[c]: ws.cell(5,c+1).value for c in range(len(headers))}
    print('  Sample:', row5)
    print()
"
```

Repeat for each vendor when running multiple vendors.

Verify:
- Correct columns present (check against `studio_columns` in vendor_info.json)
- Data actually populated — not mostly blank rows
- Extra fields captured (Designer, Collection, parsed Dimensions, etc.)
- Tearsheet Links look correct
- Multi-link categories: products from ALL listing URLs appear (not just the first)

---

## STEP 7 — Ask before full run

Report the test results to the user for each vendor:
- Sheets tested, product count per sheet, column count
- A sample of what was captured
- Code generation time (if scraper was newly written)
- Test scrape time (from orchestrator timing summary)
- Any issues found

Then ask: **"The test looks good. Ready to run the full scrape for [vendor list]?"**

Only proceed to STEP 8 when the user confirms.

---

## STEP 8 — Full run

```bash
python orchestrator.py "$ARGUMENTS" --headless true
```

When done, report per vendor:
- Total categories scraped (and any skipped with reason)
- Total products per category
- Output file: `<Vendor Name>/Data/<Vendor Name>.xlsx`
- Total columns in each sheet (dynamically determined)
- **Scrape time** (from orchestrator timing summary `scrape=Xs`)
- Any warnings or errors encountered
- Combined session time (shown at bottom of orchestrator summary)
