---
description: Quick test scrape for a vendor — ALL categories, minimum 5 products each. Usage: /test-scrape <Vendor Name>
---

You are a senior Python developer and web-scraping expert.
Run a **quick test scrape** for the vendor: **$ARGUMENTS**

Test mode scrapes **all categories** with at least **5 products each**
and saves to `<Vendor Name>/Data/<Vendor Name>_TEST.xlsx`.
This lets you verify the scraper is working before committing to a full run.

---

## STEP 1 — Verify the vendor

```bash
cd "d:/workspace/playyeard/all scrapers/Agentic system"
python vendor_parser.py "$ARGUMENTS"
```

If not found, run `python vendor_parser.py --list` and tell the user the closest match.

---

## STEP 2 — Run the test

```bash
cd "d:/workspace/playyeard/all scrapers/Agentic system"
python orchestrator.py "$ARGUMENTS" --test --headless true
```

| Exit code | Meaning | Action |
|---|---|---|
| 0 | Test passed | Inspect output, report to user |
| 2 | Scraper crashed | Fix it, retry |
| 3 | No scraper yet | Follow STEPS 3-5 from /scrape-vendor to generate scraper, then re-run with --test |

---

## STEP 3 — Inspect the test output

Open and read the test Excel file:
```bash
cd "d:/workspace/playyeard/all scrapers/Agentic system"
python -c "
import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('\"$ARGUMENTS\"/Data/\"$ARGUMENTS\"_TEST.xlsx')
for sheet in wb.sheetnames:
    ws = wb[sheet]
    headers = [ws.cell(4, c).value for c in range(1, ws.max_column+1) if ws.cell(4,c).value]
    rows = ws.max_row - 4
    print(f'Sheet: {sheet} | {rows} products | {len(headers)} columns')
    print('  Columns:', headers)
    # Show first data row
    row5 = [ws.cell(5, c).value for c in range(1, ws.max_column+1)]
    print('  Row 1:', dict(zip(headers, row5)))
    print()
"
```

Check:
- Are the right columns present (matching the category's studio_columns)?
- Is the data actually populated (not all blank)?
- Are extra fields being captured (Designer, Collection, Dimensions parsed, etc.)?
- Is the Tearsheet Link correct?

---

## STEP 4 — Report to the user

Tell the user:
- Categories tested and product count per category
- Column count per sheet (and sample column list)
- Whether the data looks complete and correct
- The test output file path: `<Vendor Name>/Data/<Vendor Name>_TEST.xlsx`
- Whether you recommend proceeding with a full run (`/scrape-vendor $ARGUMENTS`)
- Any issues found (missing fields, wrong selectors, empty rows, etc.)

If the test looks good, suggest the user run:
```
/scrape-vendor $ARGUMENTS
```
