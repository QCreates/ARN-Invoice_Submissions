import asyncio
import warnings

from datetime import datetime, timedelta

from playwright.async_api import async_playwright
import requests
import pandas as pd

target_site = "https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue?openid.assoc_handle=amzn_vc_us_v2&openid.claimed_id=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fid%2Famzn1.account.AERNEUPQAXWSM2DTBSEAINFHCFUA&openid.identity=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fid%2Famzn1.account.AERNEUPQAXWSM2DTBSEAINFHCFUA&openid.mode=id_res&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&openid.op_endpoint=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fsignin&openid.response_nonce=2025-03-14T22%3A56%3A06Z3064750971631980168&openid.return_to=https%3A%2F%2Fvendorcentral.amazon.com%2Fkt%2Fvendor%2Fmembers%2Fafi-shipment-mgr%2Fshippingqueue&openid.signed=assoc_handle%2Cclaimed_id%2Cidentity%2Cmode%2Cns%2Cop_endpoint%2Cresponse_nonce%2Creturn_to%2Cns.pape%2Cpape.auth_policies%2Cpape.auth_time%2Csigned&openid.ns.pape=http%3A%2F%2Fspecs.openid.net%2Fextensions%2Fpape%2F1.0&openid.pape.auth_policies=SinglefactorWithPossessionChallenge&openid.pape.auth_time=2025-03-14T22%3A55%3A39Z&openid.sig=c8PoFmf9ENHIP6yMONolJjf1GrheveoIWNyJJPz%2Fb68%3D&serial="
wrhs_file = "../Warehouse_Ship_Days.xlsx"
log_file = "./ASN_Status.xlsx"

def format_date(user_input):
    """
    Converts a date from 'MM/DD/YYYY' to 'Mon DD, YYYY'.
    
    Example:
    Input: "03/10/2025"
    Output: "Mar 10, 2025"
    """
    try:
        formatted_date = datetime.strptime(user_input, "%m/%d/%Y").strftime("%b %d, %Y")
        print(formatted_date)
        return formatted_date
    except ValueError:
        return "❌ Invalid date format. Please enter date as MM/DD/YYYY."

def get_eta(date, eta):
    """
    Calculates the estimated arrival date given a shipping date and ETA in days,
    skipping weekends.

    :param date: A string representing the shipping date in "MM/DD/YYYY" format.
    :param eta: An integer representing the estimated transit time in days.
    :return: A string representing the adjusted arrival date in "MM/DD/YYYY" format.
    """
    try:
        # Convert date string to datetime object
        ship_date = datetime.strptime(date, "%m/%d/%Y")
        days_added = 0

        while days_added < eta:
            ship_date += timedelta(days=1)
            if ship_date.weekday() not in [5, 6]:  # Skip weekends (Saturday=5, Sunday=6)
                days_added += 1

        # If the final arrival date lands on a weekend, push it to Monday
        while ship_date.weekday() in [5, 6]:
            ship_date += timedelta(days=1)

        return ship_date.strftime("%m/%d/%Y")
    except ValueError:
        return "❌ Invalid date format. Please use MM/DD/YYYY."

def extract_excel_data(file_path):
    # Load the Excel file (pandas uses openpyxl automatically for .xlsx)
    df = pd.read_excel(file_path, engine="openpyxl")

    # Ensure we have at least 3 columns
    if df.shape[1] < 3:
        print("❌ The Excel file must have at least 3 columns.")
        return

    # Extract first and third columns
    data_dict = {}
    for index, row in df.iterrows():
        key = str(row.iloc[0]).strip()  # First column as string
        value = row.iloc[2]  # Third column

        # Ensure value is an integer
        try:
            value = int(value)
            data_dict[key] = value
        except ValueError:
            print(f"⚠️ Skipping row {index+1}: '{value}' is not an integer.")

        # Stop when we hit an empty row
        if pd.isna(key) or pd.isna(value):
            break


    return data_dict

# Don't ask how this function works it just does
async def connect_browser():
    """Connect to an already running Chrome instance via CDP."""
    ws_url = "http://localhost:9222/json/version"

    try:
        response = requests.get(ws_url)
        response.raise_for_status()
        data = response.json()

        print(f"\n✅ Chrome DevTools Protocol detected at: {ws_url}")
        print(f"🔗 WebSocket Debugger URL: {data['webSocketDebuggerUrl']}")

        playwright = await async_playwright().start()
        browser = await playwright.chromium.connect_over_cdp(data["webSocketDebuggerUrl"])

        if not browser.contexts:
            print("⚠️ No browser contexts found! Creating a new one...")
            context = await browser.new_context()
        else:
            context = browser.contexts[0]

        target_page = None
        for page in context.pages:
            if "chrome://" not in page.url and "chatgpt.com" not in page.url:
                target_page = page
                break

        if not target_page:
            print("⚠️ No suitable open page found! Creating a new one...")
            target_page = await context.new_page()
            await target_page.goto(target_site)
            print("✅ New page created!")

        await target_page.bring_to_front()
        print(f"✅ Connected successfully to: {target_page.url}")

        # Confirm Playwright can interact with the page
        try:
            await target_page.evaluate("document.body.click()")
            print("✅ Permission check passed: Playwright can interact with the page.")
        except Exception as e:
            print(f"❌ Permission Error: {e}")

        return playwright, browser, target_page

    except requests.exceptions.RequestException as e:
        print(f"❌ Error connecting to Chrome DevTools: {e}")
        return None, None, None
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return None, None, None

async def extract_pg_data(page, formatted_date_input):
    """
    Extract all elements with class 'kat-label.kat-label-light-text'.
    This finds all ARN's and their links
    """
    await page.wait_for_selector("div.rdt_TableRow", state="attached", timeout=20000)
    table_rows = await page.query_selector_all("div.rdt_TableRow")

    if not table_rows:
        print("❌ No table rows found. The page structure may have changed or elements are not loading.")
        return {}

    print(f"✅ Found {len(table_rows)} table rows.")  # Debugging print

    table_data = {}

    for row in table_rows:
        # Extract ARN & Link
        arn_link_element = await row.query_selector("kat-link[id^='sq-table-arn-link']")
        arn = await arn_link_element.get_attribute("label") if arn_link_element else None
        arn_link = await arn_link_element.get_attribute("href") if arn_link_element else None
        arn_link = f"https://vendorcentral.amazon.com{arn_link}" if arn_link else None

        # Extract Pickup Date
        pickup_element = await row.query_selector("kat-label[id^='sq-table-sl2'][text^='Pickup:']")
        pickup_date = await pickup_element.get_attribute("text") if pickup_element else None

        # Extract Ship From & Ship To Location
        ship_from_to_elements = await row.query_selector_all("kat-label[id^='sq-table-st']")
        ship_location = " | ".join([await loc.get_attribute("text") for loc in ship_from_to_elements]) if ship_from_to_elements else None

        if arn and pickup_date:
            if formatted_date_input in pickup_date:
                arn_link = arn_link.replace("shipmentdetail?rr=", "asnsubmission?arn=")
                arn_link = arn_link.replace("&asn=", "&asnId=")
                print(f"{arn}: {arn_link}---{pickup_date}---{ship_location.split(",")[0]}")
                table_data[arn] = [arn_link, pickup_date, ship_location.split(",")[0]]

    if table_data:
        print(f"✅ Extracted {len(table_data)} rows. Sample: {list(table_data.items())[:3]}")
    else:
        print("🚨 No data extracted! Check previous logs for missing conditions.")

    return table_data

async def paginate_and_extract(page, formatted_date_input):
    """
    Extract ARN's using function above and click the 'Next' button until no more pages exist.
    """
    table_data = {}

    while True:
        try:
            new_data = await extract_pg_data(page, formatted_date_input)
            table_data.update(new_data)
            print(f"New Data Length: {len(new_data)}")
            next_button = await page.query_selector("div#sq-pag-next-div")
            if (next_button and len(new_data) > 0):
                print("➡️ Next page found! Clicking 'Next'...")
                await next_button.click()
                await asyncio.sleep(3)  # Wait for the next page to load
            else:
                print("No more pages with required pickup date...\n")
                break
        except:
            print("No data was found on this page...\n\n")
            break        
    
    return table_data

async def cont_to_step(page, step_num):
    await page.wait_for_selector(f"kat-button[label='Continue to step {step_num}']", timeout=20000)
    button = await page.query_selector(f"kat-button[label='Continue to step {step_num}']")
    if button:
        await button.click()
        print(f"Step {step_num} button clicked")
    else:
        print(f"Error: Continue to step {step_num} button not found")

async def fill_tracking_numbers(page):
    """
    Clicks the tracking number cell and selects the corresponding tracking number 
    for each row where cartonLabelBarcode starts with 'AMZN'.
    """
    await page.wait_for_selector("div[col-id='carrierTrackingNumber']", state="attached", timeout=20000)
    tracking_elements = await page.query_selector_all("div[col-id='carrierTrackingNumber']")
    carton_label_elements = await page.query_selector_all("div[col-id='cartonLabelBarcode']")

    if not tracking_elements or not carton_label_elements:
        print("❌ No tracking or carton label elements found.")
        return

    print(f"✅ Found {len(carton_label_elements)} rows. Checking each for 'AMZN' labels...")
    for i in range(len(carton_label_elements)):
        try:
            # Get carton label text
            label_text = await carton_label_elements[i].inner_text()
            if label_text.startswith("AMZN"):
                print(f"Row {i + 1}: Found carton label '{label_text}'. Selecting tracking number...")

                # Click tracking number cell
                await tracking_elements[i].dblclick()

                # **Wait for the dropdown to appear**
                await page.wait_for_selector(".ag-rich-select-list", state="visible", timeout=5000)

                # Get all available tracking numbers in the dropdown
                rich_select_rows = await page.query_selector_all(".ag-rich-select-row")

                if len(rich_select_rows) > i:  # Make sure there's a corresponding row
                    # Click the corresponding tracking number (i-th row)
                    row_text = await rich_select_rows[i-1].inner_text()
                    if row_text.strip():
                        print(f"Row {i + 1}: Selecting tracking number '{row_text}'")
                        await rich_select_rows[i-1].click()

                        # Wait until the selected value appears inside the tracking cell
                        await page.wait_for_selector(f"div[col-id='carrierTrackingNumber']:has-text('{row_text}')", timeout=5000)

                        print(f"✅ Row {i + 1}: Tracking number set to '{row_text}'")

                    else:
                        print(f"⚠️ Row {i + 1}: Found empty tracking number option, skipping.")

                else:
                    print(f"❌ Row {i + 1}: No corresponding tracking number found!")

        except Exception as e:
            print(f"⚠️ Skipping row {i + 1} due to error: {e}")

    print("✅ Finished filling tracking numbers.")
    await cont_to_step(page, "4")

async def set_ship_date(page, date):
    """
    Sets the given date (MM/DD/YYYY) inside the kat-date-picker#asnlabel-shipdate-picker element.
    """
    # Execute JavaScript inside the page to access the shadow DOM
    date_picker_element = await page.evaluate_handle("""
        () => {
            let datePicker = document.querySelector('kat-date-picker#asnlabel-shipdate-picker');
            return datePicker ? datePicker.shadowRoot.querySelector('kat-input') : null;
        }
    """)

    if not date_picker_element:
        print("❌ Ship date picker input field not found.")
        return

    # Find the input field inside the shadow root
    input_field = await page.evaluate_handle("""
        (datePicker) => {
            return datePicker ? datePicker.shadowRoot.querySelector('input[placeholder=\"MM/DD/YYYY\"]') : null;
        }
    """, date_picker_element)

    if not input_field:
        print("❌ Ship date input field inside the shadow DOM not found.")
        return

    # Fill the input field with the date
    await input_field.fill(date)
    print(f"✅ Ship Date set to: {date}")

async def set_arrival_date(page, date, eta):
    """
    Sets the given date (MM/DD/YYYY) inside the kat-date-picker#asnlabel-edd-picker element.
    """
    # Execute JavaScript inside the page to access the shadow DOM
    date_picker_element = await page.evaluate_handle("""
        () => {
            let datePicker = document.querySelector('kat-date-picker#asnlabel-edd-picker');
            return datePicker ? datePicker.shadowRoot.querySelector('kat-input') : null;
        }
    """)

    if not date_picker_element:
        print("❌ Date picker input field not found.")
        return

    # Find the input field inside the shadow root
    input_field = await page.evaluate_handle("""
        (datePicker) => {
            return datePicker ? datePicker.shadowRoot.querySelector('input[placeholder=\"MM/DD/YYYY\"]') : null;
        }
    """, date_picker_element)

    if not input_field:
        print("❌ Date input field inside the shadow DOM not found.")
        return

    # Fill the input field with the date
    await input_field.fill(get_eta(date, eta))
    print(f"✅ EDD Date set to: {get_eta(date, eta)}")

async def adjust_and_click_submit_button(page):
    """
    Adjusts the 'Confirm and submit shipment' button by removing the 'disabled' attribute and clicks it.
    """
    # Locate the button using the label
    button_selector = 'kat-button[label="Confirm and submit shipment"]'
    
    # Wait for the button to be available in the DOM
    await page.wait_for_selector(button_selector, state="attached", timeout=5000)

    # Remove the 'disabled' attribute
    await page.evaluate(f"""
        let button = document.querySelector('{button_selector}');
        if (button) {{
            button.removeAttribute('disabled');
            console.log("✅ Disabled attribute removed from the button.");
        }} else {{
            console.log("❌ Button not found.");
        }}
    """)

    # Click the button
    await page.click(button_selector)
    print("✅ 'Confirm and submit shipment' button clicked.")

async def asn_submission(page, link, date_input, eta):
    await page.goto(link)  # Navigate to the link
    print(link)
    await cont_to_step(page, "2")
    await cont_to_step(page, "3")
    await fill_tracking_numbers(page)
    await set_ship_date(page, date_input)
    await set_arrival_date(page, date_input, eta)
    await adjust_and_click_submit_button(page)
    await asyncio.sleep(3)

async def run_script():
    log_data = []
    """Get Warehouse ship days"""
    eta_by_wrhs = extract_excel_data(wrhs_file)

    """Runs the script and ensures proper cleanup."""
    warnings.filterwarnings("ignore", category=ResourceWarning)  # Suppress asyncio resource warnings

    playwright, browser, page = await connect_browser()
    if not page:
        print("❌ No valid page found. Exiting.")
        return

    print("✅ Playwright is running. Press CTRL+C to stop.")
    
    try:
        #######################################################################################
        ############################### ALL MAIN CODE RAN BELOW ###############################
        #######################################################################################
        # Navigate to below page before anything:
        await page.goto("https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue")
        await asyncio.sleep(2)

        """Get Pickup Date"""
        date_input = input("Enter a date (MM/DD/YYYY): ")
        formatted_date_input = format_date(date_input)
        
        """Find All Products and Store in a Dictionary"""        
        arn_data = await paginate_and_extract(page, formatted_date_input)
        print(f"\n🔎 Extracted {len(arn_data)} ARNs: {arn_data}\n")

        """Visit Each ASN Submission Page"""
        submission_status = "Error"  # Default status in case of failure
        print("\n\n**************************************************\n**************************************************\n**************************************************\n*************Now Beginning Submissions************\n**************************************************\n**************************************************\n**************************************************\n")
        for key, value in arn_data.items():
            try:
                print(f"{(value[2])}: {eta_by_wrhs.get(value[2])} day(s)")
                await asn_submission(page, value[0], date_input, eta_by_wrhs.get(value[2]))
                submission_status = "Submitted"
                print("\n")
            except TypeError as wrhsE:
                print(f"❌Error with warehouse {value[2]}... {wrhsE}\n\n")
                submission_status = "Warehouse Not Found"
            except e:
                print(f"❌Error with ARN {key}... {e}\n\n")
                submission_status = "Error"

            log_data.append([key, value[2], value[0], submission_status])
            
        df = pd.DataFrame(log_data, columns=["ARN", "Warehouse", "Link", "Status"])

        # Save DataFrame to Excel
        df.to_excel(log_file, index=False, engine="openpyxl")
        print(f"✅ Log saved to {log_file}")

        #######################################################################################
        ############################### ALL MAIN CODE RAN ABOVE ###############################
        #######################################################################################
    except TypeError as typeE:
        print(f"\nNo labels found... {typeE}")
    except KeyboardInterrupt:
        print("\n🛑 Shutting down gracefully...")
    finally:
        if browser:
            await browser.close()
        if playwright:
            await playwright.stop()
        print("✅ Playwright closed successfully.")

if __name__ == "__main__":
    try:
        asyncio.run(run_script())  # Ensures only one event loop is running
    except RuntimeError as e:
        print(f"❌ RuntimeError: {e}. Trying alternative method...")
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(run_script())
