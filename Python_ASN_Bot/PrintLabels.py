import asyncio
import warnings
from datetime import datetime
from playwright.async_api import async_playwright
import requests
import pandas as pd

# URLs
target_site = "https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue"
shipment_detail_base = "https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shipmentdetail?rr="
log_file = "./PrintLabel_Log.xlsx"


def format_date(user_input):
    """Convert a date from 'MM/DD/YYYY' to 'Mon DD, YYYY'."""
    try:
        dt = datetime.strptime(user_input, "%m/%d/%Y")
        return f"{dt.strftime('%b')} {dt.day}, {dt.year}"
    except ValueError:
        return None


async def connect_browser():
    """Connect to an already running Chrome instance via CDP."""
    ws_url = "http://localhost:9222/json/version"
    try:
        response = requests.get(ws_url)
        response.raise_for_status()
        data = response.json()

        print(f"\n✅ Chrome DevTools detected at: {ws_url}")
        playwright = await async_playwright().start()
        browser = await playwright.chromium.connect_over_cdp(data["webSocketDebuggerUrl"])
        context = browser.contexts[0] if browser.contexts else await browser.new_context()

        target_page = None
        for p in context.pages:
            if "chrome://" not in p.url and "chatgpt.com" not in p.url:
                target_page = p
                break
        if not target_page:
            target_page = await context.new_page()
            await target_page.goto(target_site)
            print("✅ New page created.")

        await target_page.bring_to_front()
        print(f"✅ Connected to: {target_page.url}")
        return playwright, browser, target_page

    except Exception as e:
        print(f"❌ Could not connect to Chrome: {e}")
        return None, None, None


async def extract_pg_data(page, formatted_date_input):
    """Extract ARNs from current page based on the pickup date."""
    await page.wait_for_selector("div.rdt_TableRow", state="attached", timeout=20000)
    rows = await page.query_selector_all("div.rdt_TableRow")
    table_data = {}

    for row in rows:
        try:
            arn_el = await row.query_selector("kat-link[id^='sq-table-arn-link']")
            arn = await arn_el.get_attribute("label") if arn_el else None

            pickup_el = await row.query_selector("kat-label[id^='sq-table-sl2']")
            pickup_date = await pickup_el.get_attribute("text") if pickup_el else None

            if arn and pickup_date and formatted_date_input in pickup_date:
                table_data[arn] = pickup_date
                print(f"📦 Found ARN {arn} ({pickup_date})")

        except Exception as e:
            print(f"⚠️ Skipping row due to error: {e}")

    return table_data


async def paginate_and_extract(page, formatted_date_input):
    """Go through all pages and collect ARNs for the date."""
    all_arns = {}
    page_num = 1
    while True:
        print(f"\n📄 Extracting ARNs from page {page_num}...")
        new_data = await extract_pg_data(page, formatted_date_input)
        all_arns.update(new_data)

        next_button = await page.query_selector("div#sq-pag-next-div")
        if next_button and new_data:
            await next_button.click()
            await asyncio.sleep(3)
            page_num += 1
        else:
            break

    print(f"\n✅ Total ARNs found: {len(all_arns)}")
    return all_arns


async def click_print_sequence(page, arn):
    """
    Visit shipment detail page and:
    1️⃣ Click the first button (data-action="print")
    2️⃣ Then click the confirmation kat-button(label="Print carton labels")
    """
    try:
        shipment_url = f"{shipment_detail_base}{arn}"
        await page.goto(shipment_url)
        print(f"🔗 Opened: {shipment_url}")

        # Wait for the Print shipping labels button
        await page.wait_for_selector("kat-button[label='Print shipping labels']", timeout=15000)
        button = await page.query_selector("kat-button[label='Print shipping labels']")

        if button:
            await button.click()
            print(f"🖨️ Clicked 'Print shipping labels' for ARN {arn}\n")
            return True
        else:
            print(f"⚠️ 'Print shipping labels' button not found for ARN {arn}\n")
            return False

    except Exception as e:
        print(f"❌ Error while processing ARN {arn}: {e}\n")
        return False



async def run_script():
    warnings.filterwarnings("ignore", category=ResourceWarning)
    playwright, browser, page = await connect_browser()
    if not page:
        return

    try:
        await page.goto(target_site)
        await asyncio.sleep(2)

        date_input = input("Enter pickup date (MM/DD/YYYY): ").strip()
        formatted_date = format_date(date_input)
        if not formatted_date:
            print("❌ Invalid date format.")
            return

        # Step 1: Extract ARNs
        arn_data = await paginate_and_extract(page, formatted_date)

        # Step 2: Process each ARN
        log = []
        for arn in arn_data.keys():
            success = await click_print_sequence(page, arn)
            log.append([arn, f"{shipment_detail_base}{arn}", "Printed" if success else "Failed"])

        # Step 3: Save results
        df = pd.DataFrame(log, columns=["ARN", "Link", "Status"])
        df.to_excel(log_file, index=False, engine="openpyxl")
        print(f"✅ Log saved to {log_file}")

    except KeyboardInterrupt:
        print("\n🛑 Script manually stopped.")
    finally:
        if browser:
            await browser.close()
        if playwright:
            await playwright.stop()
        print("✅ Playwright closed.")


if __name__ == "__main__":
    try:
        asyncio.run(run_script())
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(run_script())
