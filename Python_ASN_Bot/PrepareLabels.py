import asyncio
import warnings

from datetime import datetime, timedelta

from playwright.async_api import async_playwright
import requests
import pandas as pd

target_site = "https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue?openid.assoc_handle=amzn_vc_us_v2&openid.claimed_id=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fid%2Famzn1.account.AERNEUPQAXWSM2DTBSEAINFHCFUA&openid.identity=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fid%2Famzn1.account.AERNEUPQAXWSM2DTBSEAINFHCFUA&openid.mode=id_res&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&openid.op_endpoint=https%3A%2F%2Fvendorcentral.amazon.com%2Fap%2Fsignin&openid.response_nonce=2025-03-14T22%3A56%3A06Z3064750971631980168&openid.return_to=https%3A%2F%2Fvendorcentral.amazon.com%2Fkt%2Fvendor%2Fmembers%2Fafi-shipment-mgr%2Fshippingqueue&openid.signed=assoc_handle%2Cclaimed_id%2Cidentity%2Cmode%2Cns%2Cop_endpoint%2Cresponse_nonce%2Creturn_to%2Cns.pape%2Cpape.auth_policies%2Cpape.auth_time%2Csigned&openid.ns.pape=http%3A%2F%2Fspecs.openid.net%2Fextensions%2Fpape%2F1.0&openid.pape.auth_policies=SinglefactorWithPossessionChallenge&openid.pape.auth_time=2025-03-14T22%3A55%3A39Z&openid.sig=c8PoFmf9ENHIP6yMONolJjf1GrheveoIWNyJJPz%2Fb68%3D&serial="
shipment_file = "../shipment_details.xlsx"
log_file = "./Label_Prep_Status.xlsx"

def format_date(user_input):
    """
    Converts a date from 'MM/DD/YYYY' to 'Mon DD, YYYY'.
    
    Example:
    Input: "03/10/2025"
    Output: "Mar 10, 2025"
    """
    try:
        dt = datetime.strptime(user_input, "%m/%d/%Y")
        mon = dt.strftime("%b")      # e.g. "Jul"
        day = dt.day                 # e.g. 4 instead of 04
        year = dt.year               # e.g. 2025
        formatted_date = f"{mon} {day}, {year}"
        print(formatted_date)
        return formatted_date
    except ValueError:
        return "❌ Invalid date format. Please enter date as MM/DD/YYYY."

def findNumCartons(pack, master):
    print()

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
        arn_link = f"https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/labelmapping?arn={arn}&isLegacy=false"

        # Extract Pickup Date
        pickup_element = await row.query_selector("kat-label[id^='sq-table-sl2'][text^='Pickup:']")
        pickup_date = await pickup_element.get_attribute("text") if pickup_element else None

        # Extract Ship From & Ship To Location
        ship_from_to_elements = await row.query_selector_all("kat-label[id^='sq-table-st']")
        ship_location = " | ".join([await loc.get_attribute("text") for loc in ship_from_to_elements]) if ship_from_to_elements else None

        if arn and pickup_date:
            if formatted_date_input in pickup_date:
                if len(ship_location.split(",")) > 0:
                    ship_location = ship_location.split(",")[0]
                print(f"{arn}: {pickup_date}---{ship_location}---{arn_link}")
                table_data[arn] = [pickup_date, ship_location, arn_link]

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
    await page.wait_for_selector(f"kat-button[label='Continue to step {step_num}']", timeout=5000)
    button = await page.query_selector(f"kat-button[label='Continue to step {step_num}']")
    try:
        await button.click()
        print(f"Step {step_num} button clicked")
    except:
        print(f"Error: Cannot press step {step_num} button")

async def extract_pack_info(page):

    total_packs = []
    elements = await page.query_selector_all("div.rdt_TableCell")
    asin_elements = await page.query_selector_all("div.sb-asinRow-detail-div")

    pack_i = 10

    for i in range(len(asin_elements)):
        packs = await elements[pack_i].inner_text()
        asin = await asin_elements[i].inner_text()
        print(asin.split("ASIN:")[1].split("Model:")[0].strip(), packs.split("/")[1].strip(), asin.split("Purchase order:")[1].split("ASIN:")[0].strip())
        total_packs.append([asin.split("ASIN:")[1].split("Model:")[0].strip(), packs.split("/")[1].strip(), asin.split("Purchase order:")[1].split("ASIN:")[0].strip()])
        pack_i += 6
    return total_packs

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
        # Navigate to below page before anything:
        await page.goto("https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue")
        await asyncio.sleep(2)

        """Get Pickup Date"""
        date_input = input("Enter a date (MM/DD/YYYY): ")
        formatted_date_input = format_date(date_input)

        """Find All Products and Store in a Dictionary"""        
        arn_data = await paginate_and_extract(page, formatted_date_input)
        print(f"\n🔎 Extracted {len(arn_data)} ARNs: {arn_data}\n")

        """Store in array and sort by warehouse"""
        arn_list = [
            [arn, data[1], data[2]]
            for arn, data in arn_data.items()
        ]
        arn_list = sorted(arn_list,key=lambda l:l[1])

        log_data = []
        shipment_data = pd.read_excel(
            shipment_file,
            skiprows=7, 
            usecols="A,C,D,M", 
            header=None, 
            names=["Wrhs", "ASIN", "PO", "Pack"]
        )
        print(shipment_data.head())
        shipment_dict = {
        f"{row['ASIN']}::{(row['Wrhs']).split()[0] if str(row['Wrhs']) != 'nan' else 'Default'}": {'PO': row['PO'], 'Pack': row['Pack']}
            for _, row in shipment_data.iterrows()
        }
 
        """Perform actions"""
        print(shipment_dict.keys())
        for arn, wrhs, link in arn_list:
            try:
                await asyncio.sleep(1)
                await page.goto(link)  # Navigate to the link
                print("\n_________________________________\n", wrhs)
                await cont_to_step(page, 2)
                # Extract the info from the cell after moving to step 2
            except Exception as e:
                print("error lolz", e)
                log_data.append([arn, wrhs, link, 0, "Error"])
            try:
                await page.wait_for_selector("input[name='packingMethod']", timeout=5000)
                radio_buttons = await page.query_selector_all("input[name='packingMethod']")
                if radio_buttons and len(radio_buttons) > 0:
                    await radio_buttons[0].click(force=True)
                
                all_pack_info = await extract_pack_info(page)

                """ FILL OUT PACK INFORMATION HERE TO CONFIRM THE LABEL"""
                kat_index = 2
                confirm_index = 1
                
                for asn, pack, po in all_pack_info:
                    dict_key = f"{asn}::{wrhs}"
                    print(f"VENDORCENTRAL ___ ASN: {asn}       Pack: {pack}    PO:{po}")

                    masterpack = int(shipment_dict.get(dict_key)['Pack'])
                    pack = int(pack)
                    unit = 0
                    cartons = 0
                    if (pack < masterpack):
                        unit = 1
                        cartons = pack
                    elif (pack % masterpack != 0):
                        print("INDIVISIBLE PACK")
                    else:
                        unit = masterpack
                        cartons = pack/masterpack

                    if (shipment_dict.get(dict_key)['PO'] == po):
                        print(f"SHEET         ___ ASN: {asn} Master Pack: {shipment_dict.get(dict_key)['Pack']}   PO{shipment_dict.get(dict_key)['PO']}")
                        print(f"UnitPerCartons: {unit}          Cartons: {cartons}")
                        # Wait for at least 2 kat-inputs to be present in the DOM
                        await page.wait_for_function(
                            "() => document.querySelectorAll('kat-input').length >= 2",
                            timeout=15000
                        )
                        # Wait until kat-inputs are available
                        await page.wait_for_function("() => document.querySelectorAll('kat-input').length >= 2", timeout=10000)

                        # Typing into Units Per Carton
                        unit_input = await page.evaluate_handle("""
                        (index) => {
                            const katInput = document.querySelectorAll('kat-input')[index];
                            return katInput?.shadowRoot?.querySelector('input');
                        }
                        """, kat_index)

                        if unit_input:
                            await unit_input.evaluate("el => el.value = ''")
                            await unit_input.type(str(unit), delay=50)
                        else:
                            print("❌ Could not find Units Per Carton input field")

                        # Typing into Cartons
                        carton_input = await page.evaluate_handle("""
                        (index) => {
                            const katInput = document.querySelectorAll('kat-input')[index];
                            return katInput?.shadowRoot?.querySelector('input');
                        }
                        """, kat_index + 1)

                        if carton_input:
                            await carton_input.evaluate("el => el.value = ''")
                            await carton_input.type(str(cartons), delay=50)
                        else:
                            print("❌ Could not find Cartons input field")


                    else:
                        print(f"SHEET         NOT FOUND... ASN: {asn}, PO: {po}, S_PO: {shipment_dict.get(dict_key)['PO']}")
                    print(kat_index)
    #                await conf_buttons[confirm_index].click(force=True)

                    kat_index += 3
                    confirm_index += 1
                try:
                    print("submitting...")
                    await page.screenshot(path=f"screenshot_{arn}.png")
                    
                    conf_button = await page.query_selector('kat-button[label="Confirm all SKUs"]')
                    await conf_button.click(force=True)

                    new_conf_button = await page.query_selector('kat-button[label="Confirm and print labels"]')
                    await new_conf_button.click(force=True)


                    log_data.append([arn, wrhs, link, len(all_pack_info), "Complete"])
                    await asyncio.sleep(2)  # Wait to allow submission process

                except Exception as e:
                    log_data.append([arn, wrhs, link, len(all_pack_info), "Error"])
                    print("Couldn't Find a Submit Button", e)

                if(len(all_pack_info) <= 0):
                    log_data.append([arn, wrhs, link, len(all_pack_info), "Already Completed"])
                    
            except Exception as e:
                print(f"Couldn't find radio button input\n{e}")

        """ Save DataFrame to Excel """
        df = pd.DataFrame(log_data, columns=["ARN", "Warehouse", "Link", "# Of Packs", "Status"])
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
    asyncio.run(run_script())