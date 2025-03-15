import asyncio
import warnings
from playwright.async_api import async_playwright
import json
import requests
from datetime import datetime

target_site = "https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue?openid.assoc_handle=amzn_vc_us_v2&openid.claimed_id=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fid%2Famzn1.account.AERNEUPQAXWSM2DTBSEAINFHCFUA&openid.identity=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fid%2Famzn1.account.AERNEUPQAXWSM2DTBSEAINFHCFUA&openid.mode=id_res&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&openid.op_endpoint=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fsignin&openid.response_nonce=2025-03-14T22%3A56%3A06Z3064750971631980168&openid.return_to=https%3A%2F%2Fvendorcentral.amazon.com%2Fkt%2Fvendor%2Fmembers%2Fafi-shipment-mgr%2Fshippingqueue&openid.signed=assoc_handle%2Cclaimed_id%2Cidentity%2Cmode%2Cns%2Cop_endpoint%2Cresponse_nonce%2Creturn_to%2Cns.pape%2Cpape.auth_policies%2Cpape.auth_time%2Csigned&openid.ns.pape=http%3A%2F%2Fspecs.openid.net%2Fextensions%2Fpape%2F1.0&openid.pape.auth_policies=SinglefactorWithPossessionChallenge&openid.pape.auth_time=2025-03-14T22%3A55%3A39Z&openid.sig=c8PoFmf9ENHIP6yMONolJjf1GrheveoIWNyJJPz%2Fb68%3D&serial="

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
        print("❌ No table rows found.")
        return {}

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

        if arn and formatted_date_input in pickup_date:
            arn_link = arn_link.replace("shipmentdetail?rr=", "asnsubmission?arn=")
            arn_link = arn_link.replace("&asn=", "&asnId=")
            print(f"{arn}: {arn_link}---{pickup_date}---{ship_location.split(",")[0]}")
            table_data[arn] = [arn_link, pickup_date, ship_location.split(",")[0]]

    print(f"✅ Extracted {len(table_data)} rows:")

    return table_data

async def paginate_and_extract(page, formatted_date_input):
    """
    Extract ARN's using function above and click the 'Next' button until no more pages exist.
    """
    table_data = {}

    while True:
        table_data = await extract_pg_data(page, formatted_date_input)

        next_button = await page.query_selector("div#sq-pag-next-div")
        if (next_button and len(table_data) > 0):
            print("➡️ Next page found! Clicking 'Next'...")
            await next_button.click()
            await asyncio.sleep(3)  # Wait for the next page to load
        else:
            print("✅ No more 'Next' buttons found. Extraction complete.")
            break

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

async def run_script():
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

        user_date = input("Enter a date (MM/DD/YYYY): ")
        formatted_date_input = format_date(user_date)
        labels = await paginate_and_extract(page, formatted_date_input)
        print(f"\n✅ Extracted {len(labels)} total labels from all pages.")

        #######################################################################################
        ############################### ALL MAIN CODE RAN ABOVE ###############################
        #######################################################################################
    except TypeError:
        print("\nNo labels found...")
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
