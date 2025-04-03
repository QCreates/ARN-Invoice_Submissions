import asyncio
import warnings
import pandas as pd
from datetime import datetime
from playwright.async_api import async_playwright
import requests

input_file = "../invoices.xlsx"
output_file = "invoices_status.xlsx"
target_site = "https://vendorcentral.amazon.com/hz/vendor/members/invoice-creation/search-shipments"

def parse_date(date_input):
    if isinstance(date_input, pd.Timestamp):
        return date_input
    elif isinstance(date_input, str):
        return datetime.strptime(date_input, "%m/%d/%Y")
    else:
        raise ValueError(f"Invalid date format: {date_input}")

def excel_date_to_str(excel_date):
    try:
        date = pd.to_datetime(excel_date, unit='D', origin='1899-12-30')
        return date.strftime("%m/%d/%Y")
    except:
        return excel_date

async def connect_browser():
    ws_url = "http://localhost:9222/json/version"
    try:
        response = requests.get(ws_url)
        response.raise_for_status()
        data = response.json()
        print("✅ Connected to Chrome DevTools Protocol")
        playwright = await async_playwright().start()
        browser = await playwright.chromium.connect_over_cdp(data["webSocketDebuggerUrl"])
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        pages = context.pages
        page = pages[0] if pages else await context.new_page()
        if page.url != target_site:
            await page.goto(target_site)
        await page.wait_for_load_state("networkidle")
        await page.bring_to_front()
        return playwright, browser, page
    except Exception as e:
        print(f"❌ Error connecting to Chrome DevTools: {e}")
        return None, None, None

async def wait_for_loading(page):
    await page.wait_for_load_state("networkidle")
    await page.wait_for_selector("body", state="visible")

async def select_po_search(page):
    print("🔄 Selecting 'Purchase Order Number(s)' from dropdown...")
    await page.evaluate("""
        let dropdown = document.querySelector("#shipment-search-key");
        if (dropdown) {
            dropdown.value = "PURCHASE_ORDER";
            dropdown.dispatchEvent(new Event('change', { bubbles: true }));
        }
    """)
    await page.wait_for_timeout(500)
    selected_value = await page.evaluate("document.querySelector('#shipment-search-key').value")
    print("✅ Dropdown set to Purchase Order" if selected_value == "PURCHASE_ORDER" else "❌ Dropdown selection failed")
    await wait_for_loading(page)

async def process_invoices(page):
    df = pd.read_excel(input_file, engine="openpyxl")
    output_data = []
    for row_index, row in df.iterrows():
        try:
            await select_po_search(page)
            po_number = str(row.iloc[0]).strip()
            if not po_number:
                continue
            ship_date = parse_date(row["Invoice Date"])
            invoice_number = str(row.iloc[2]).strip()
            invoice_amount = float(row.iloc[3])
            print(f"Processing PO: {po_number}")
            await wait_for_loading(page)
            po_number_input = await page.wait_for_selector("#po-number", timeout=30000)
            await po_number_input.scroll_into_view_if_needed()
            await po_number_input.click()
            await po_number_input.fill(po_number)
            await page.click("#shipmentSearchTableForm-submit")
            await wait_for_loading(page)
            try:
                rows = await page.query_selector_all(".mt-row")
            except:
                print("⚠️ Could not query rows. Retrying...")
                await page.goto(target_site)
                await wait_for_loading(page)
                continue
            if not rows:
                print("⚠️ No PO found")
                output_data.append([po_number, invoice_number, invoice_amount, 0, "Not Available"])
                continue
            for r in range(len(rows)):
                try:
                    shipped_date_text = await page.locator(f"#r{r+1}-shipped_date").inner_text()
                    if parse_date(shipped_date_text.strip()) == ship_date:
                        await page.click(f"#r{r+1}-asn_checkbox-input-harmonic-checkbox ~ i")
                        await page.click("#create-inv-asn-po-toggle")
                        await wait_for_loading(page)
                        if await page.locator("input[data-asn-check='true']").is_checked():
                            await page.locator("input[data-asn-check='true']").click()
                            await page.locator(f"input[data-po-check='true'][value='{po_number}']").click()
                        await page.click("input.a-button-input[aria-labelledby='create-invoice-submit-announce']")
                        await wait_for_loading(page)
                        await page.wait_for_selector("#inv-total-amount-data")
                        await page.wait_for_function("""
                            () => {
                                const el = document.querySelector("#inv-total-amount-data");
                                return el && el.innerText.trim() !== "...";
                            }
                        """, timeout=5000)
                        total = await page.locator("#inv-total-amount-data").inner_text()
                        total = float(total.replace("$", "").replace(",", "").strip())
                        if abs(total - invoice_amount) < 0.01:
                            print("✅ USD Match")
                            await page.fill("#invoice-number", f"{int(invoice_number)}")
                            checkbox = await page.wait_for_selector("#inv-agree-checkbox")
                            await checkbox.scroll_into_view_if_needed()
                            if not await checkbox.is_checked():
                                await checkbox.check()
                            await page.wait_for_selector(".melodic-loading-overlay", state="hidden", timeout=5000)
                            await page.click("input.a-button-input[aria-labelledby='inv-submit-announce']")
                            await wait_for_loading(page)
                            await page.click("#inv-crt-redirect")
                            await wait_for_loading(page)
                            output_data.append([po_number, invoice_number, invoice_amount, total, "Submitted"])
                        else:
                            print("⚠️ Amount mismatch")
                            output_data.append([po_number, invoice_number, invoice_amount, total, "Price error"])
                            await page.goto(target_site)
                            await wait_for_loading(page)
                        break
                except Exception as e:
                    print(f"⚠️ Skipping row {r+1} due to error: {e}")
                    continue
        except Exception as outer:
            print(f"❌ Error on PO {row.iloc[0]}: {outer}")
            continue
    pd.DataFrame(output_data, columns=["PO Number", "Invoice Number", "Invoice Amount", "Total Amount", "Status"]).to_excel(output_file, index=False)
    print("✅ Invoice statuses saved.")

async def run_script():
    warnings.filterwarnings("ignore", category=ResourceWarning)
    playwright, browser, page = await connect_browser()
    if not page:
        print("❌ No valid page found. Exiting.")
        return
    try:
        await process_invoices(page)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
    finally:
        if browser:
            await browser.close()
        if playwright:
            await playwright.stop()
        print("✅ Playwright closed successfully.")

if __name__ == "__main__":
    try:
        asyncio.run(run_script())
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(run_script())
