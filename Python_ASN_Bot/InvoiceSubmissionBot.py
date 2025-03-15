import asyncio
import warnings
import pandas as pd
from datetime import datetime
from playwright.async_api import async_playwright
import requests

# Excel files
input_file = "../invoices.xlsx"
output_file = "invoices_status.xlsx"

# Vendor Central Target Site
target_site = "https://vendorcentral.amazon.com/hz/vendor/members/invoice-creation/search-shipments"

# Helper function to parse dates in M/D/YYYY format
def parse_date(date_input):
    if isinstance(date_input, pd.Timestamp):
        return date_input  # Return if it's already a datetime object
    elif isinstance(date_input, str):
        return datetime.strptime(date_input, "%m/%d/%Y")  # Convert string to datetime
    else:
        raise ValueError(f"Invalid date format: {date_input}")

# Convert Excel date format to MM/DD/YYYY
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

        target_page = None
        for page in context.pages:
            if "vendorcentral.amazon.com" in page.url:
                target_page = page
                break
        if not target_page:
            target_page = await context.new_page()
            await target_page.goto(target_site)
        await target_page.bring_to_front()
        return playwright, browser, target_page
    except Exception as e:
        print(f"❌ Error connecting to Chrome DevTools: {e}")
        return None, None, None

async def wait_for_loading(page):
    await page.wait_for_load_state("networkidle")
    await page.wait_for_selector("body", state="visible")

async def select_po_search(page):
    """Ensures the dropdown is set to 'Purchase Order Number(s)' before searching."""
    print("🔄 Selecting 'Purchase Order Number(s)' from dropdown...")

    # Set the dropdown value directly using JavaScript
    await page.evaluate("""
        let dropdown = document.querySelector("#shipment-search-key");
        if (dropdown) {
            dropdown.value = "PURCHASE_ORDER";
            dropdown.dispatchEvent(new Event('change', { bubbles: true }));
        }
    """)

    # Wait for the selection to be applied
    await page.wait_for_timeout(500)  # Small delay to allow selection update

    # Verify if selection was successful
    selected_value = await page.evaluate("document.querySelector('#shipment-search-key').value")
    if selected_value == "PURCHASE_ORDER":
        print("✅ 'Purchase Order Number(s)' successfully selected.")
    else:
        print("❌ Failed to select 'Purchase Order Number(s)'.")

    await wait_for_loading(page)

async def process_invoices(page):
    df = pd.read_excel(input_file, engine="openpyxl")
    output_data = []


    for row_index, row in df.iterrows():
        await select_po_search(page)  # Ensure correct search mode before searching POs
        po_number = str(row.iloc[0]).strip()
        ship_date = df.loc[row_index, "Invoice Date"]  # Get the date as a string
        ship_date = parse_date(ship_date)  # Convert to datetime object
        invoice_number = str(row.iloc[2]).strip()
        invoice_amount = float(row.iloc[3])

        if not po_number:
            print("Skipping empty PO Number.")
            continue

        print(f"Processing PO: {po_number}")
        await wait_for_loading(page)

        po_number_input = await page.wait_for_selector("#po-number", state="visible", timeout=30000)
        
        await po_number_input.scroll_into_view_if_needed()
        await po_number_input.click()
        await po_number_input.fill(po_number)

        await page.click("#shipmentSearchTableForm-submit")
        await wait_for_loading(page)

        rows = await page.query_selector_all(".mt-row")
        if not rows:
            print("⚠️ No PO number found")
            output_data.append([po_number, invoice_number, invoice_amount, total_amount, "Not Available"])
            continue
        
        for row_index, row in enumerate(rows):
            shipped_date_elem = await page.query_selector(f"#r{row_index+1}-shipped_date")
            shipped_date_text = await shipped_date_elem.inner_text()
            shipped_date_on_page = parse_date(shipped_date_text)
            shipped_date_invoice = parse_date(ship_date)

            if shipped_date_on_page == shipped_date_invoice:
                checkbox = await page.query_selector(f"#r{row_index+1}-asn_checkbox-input-harmonic-checkbox ~ i.a-icon.a-icon-checkbox")
                await checkbox.click()
                print(f"Clicked checkbox for {po_number} row {row_index+1}.")

                await page.click("#create-inv-asn-po-toggle")
                await wait_for_loading(page)

                asn_checkbox = await page.query_selector("input[type='checkbox'][data-asn-check='true']")
                if await asn_checkbox.is_checked():
                    await asn_checkbox.click()
                    po_checkbox = await page.query_selector(f"input[type='checkbox'][data-po-check='true'][value='{po_number}']")
                    await po_checkbox.click()

                await page.click("input.a-button-input[aria-labelledby='create-invoice-submit-announce']")
                await wait_for_loading(page)

                total_amount_elem = await page.query_selector("#inv-total-amount-data")
                if total_amount_elem:
                    total_amount_text = await total_amount_elem.inner_text()
                    total_amount_text = total_amount_text.replace("$", "").replace(",", "").strip()

                    # Ensure it's a valid number
                    try:
                        total_amount = float(total_amount_text)
                    except ValueError:
                        print(f"❌ ERROR: Could not convert total amount '{total_amount_text}' to float.")
                        output_data.append([po_number, invoice_number, invoice_amount, total_amount_text, "Invalid Amount"])
                        continue  # Skip this PO and proceed to the next one

                if abs(total_amount - invoice_amount) < 0.001:
                    print("✅ USD Match")
                    await page.fill("#invoice-number", invoice_number)
                    
                    checkbox_input = await page.wait_for_selector("#inv-agree-checkbox", timeout=30000)
                    await checkbox_input.scroll_into_view_if_needed()
                    if not await checkbox_input.is_checked():
                        await checkbox_input.check()
                    assert await checkbox_input.is_checked(), "❌ Checkbox was not successfully checked!"
                    print("✅ Invoice agreement checkbox successfully checked!")

                    await page.click("input.a-button-input[aria-labelledby='inv-submit-announce']")
                    await wait_for_loading(page)

                    create_another_invoice_btn = await page.wait_for_selector("#inv-crt-redirect", timeout=30000)
                    await create_another_invoice_btn.scroll_into_view_if_needed()
                    await create_another_invoice_btn.click()
                    await wait_for_loading(page)

                    print("✅ Clicked 'Create Another Invoice', ready for the next PO.")
                    output_data.append([po_number, invoice_number, invoice_amount, total_amount, "Submitted"])
                else:
                    print("⚠️ NO USD Match")
                    output_data.append([po_number, invoice_number, invoice_amount, total_amount, "Price error"])
                    await page.goto(target_site)
                    await wait_for_loading(page)
                    continue

    output_df = pd.DataFrame(output_data, columns=["PO Number", "Invoice Number", "Invoice Amount", "Total Amount", "Status"])
    output_df.to_excel(output_file, index=False, engine="openpyxl")
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

# ✅ Ensures the script runs properly
if __name__ == "__main__":
    try:
        asyncio.run(run_script())
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(run_script())
