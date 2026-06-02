"""Debug page structure for Hooker Furniture."""
import asyncio, sys, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from base_scraper import PlaywrightBrowser, clean_text
from bs4 import BeautifulSoup

URL = "https://hookerfurnishings.com/driftwood-three-drawer-nightstand-6820-90116-99"

async def debug():
    async with PlaywrightBrowser(headless=True) as page:
        await page.goto(URL, timeout=60000, wait_until="domcontentloaded")
        await page.wait_for_timeout(4000)

        # Click Overall Dimensions tab
        btn = await page.query_selector("button:has-text('Overall Dimensions')")
        if btn:
            await btn.click()
            await page.wait_for_timeout(1500)
            print("Clicked Overall Dimensions")

        html = await page.content()
        soup = BeautifulSoup(html, "lxml")

        print("\n=== Dimension section (around <b>Width:</b>) ===")
        for b in soup.select("b"):
            txt = clean_text(b.get_text())
            if txt in ("Width:", "Depth:", "Height:", "Weight:", "Volume:"):
                parent = b.parent
                gp = parent.parent if parent else None
                print(f"  <b>{txt}</b> parent HTML: {str(parent)[:300]}")
                if gp:
                    print(f"  Grandparent HTML: {str(gp)[:400]}")
                print()

        print("\n=== productSpecification items ===")
        for label_el in soup.select("[class*='productSpecification-itemLabel']"):
            label = clean_text(label_el.get_text())
            sib = label_el.find_next_sibling()
            val = clean_text(sib.get_text()) if sib else ""
            print(f"  {repr(label)}: {repr(val[:80])}")

        print("\n=== All buttons/tabs on page ===")
        for btn_el in soup.select("button"):
            txt = clean_text(btn_el.get_text())
            if txt and len(txt) < 50 and not txt.startswith("["):
                print(f"  [{txt}]")

asyncio.run(debug())
