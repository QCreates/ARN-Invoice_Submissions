const { Builder, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const { JSDOM } = require('jsdom');
const fs = require('fs');
const XLSX = require('xlsx');
const { time } = require('console');

// Load the workbook and read the data
const workbook = XLSX.readFile('Warehouse_Ship_Days.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];

const readline = require('readline');
const { ConsoleLogEntry } = require('selenium-webdriver/bidi/logEntries');

// Create an interface to get user input
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Initialize an empty array to store the extracted data
let shipDates = [];

// Loop through each row in the sheet and populate shipDates array
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
rows.slice(1).forEach(row => {
    shipDates.push([row[0], row[2]]);
});


// Set Chrome options to connect to the existing Chrome session
let chromeOptions = new chrome.Options();
chromeOptions = chromeOptions.addArguments("--remote-debugging-port=9222");  // Use the correct port where Chrome is running
chromeOptions = chromeOptions.debuggerAddress('localhost:9222');  // Connect to the existing Chrome session on this port

// Initialize WebDriver and connect to the existing Chrome session
let driver = new Builder().forBrowser('chrome')
    .setChromeOptions(chromeOptions)
    .build();

let arnArr = [[]];
let outputWorkbook = XLSX.utils.book_new();
let outputSheet = [['ARN', 'ASN', 'Amazon Label', 'UPS Tracking', 'Warehouse Name']];
XLSX.utils.book_append_sheet(outputWorkbook, XLSX.utils.aoa_to_sheet(outputSheet), 'ARN and ASN');

async function setDateInDatePicker(driver, date) {
    // Wait for the kat-date-picker element to be available
    let datePicker = await driver.executeScript(`
        let datePicker = document.querySelector('kat-date-picker#asnlabel-shipdate-picker');
        if (datePicker) {
            let shadowRoot = datePicker.shadowRoot;
            if (shadowRoot) {
                return shadowRoot.querySelector('kat-input');
            }
        }
        return null;
    `);

    if (datePicker) {
        // Access the shadow root of the kat-input and input the date
        let inputField = await driver.executeScript(`
            let input = arguments[0].shadowRoot.querySelector('input[placeholder="MM/DD/YYYY"]');
            return input;
        `, datePicker);

        if (inputField) {
            // Clear any existing value in the input field
            await inputField.clear();

            // Mimic typing the desired date (formatted as MM/DD/YYYY)
            await inputField.sendKeys(date);

            console.log(`Date set to: ${date}`);
        } else {
            console.log("Input field not found inside the shadow DOM.");
        }
    } else {
        console.log("Date picker not found on the page.");
    }
    
}

async function setDateInEDDDatePicker(driver, date) {
    // Wait for the kat-date-picker element to be available
    let datePicker = await driver.executeScript(`
        let datePicker = document.querySelector('kat-date-picker#asnlabel-edd-picker');
        if (datePicker) {
            let shadowRoot = datePicker.shadowRoot;
            if (shadowRoot) {
                return shadowRoot.querySelector('kat-input');
            }
        }
        return null;
    `);

    if (datePicker) {
        // Access the shadow root of the kat-input and input the date
        let inputField = await driver.executeScript(`
            let input = arguments[0].shadowRoot.querySelector('input[placeholder="MM/DD/YYYY"]');
            return input;
        `, datePicker);

        if (inputField) {
            // Clear any existing value in the input field
            await inputField.clear();

            // Mimic typing the desired date (formatted as MM/DD/YYYY)
            await inputField.sendKeys(date);

            console.log(`EDD Date set to: ${date}`);
        } else {
            console.log("Input field not found inside the shadow DOM.");
        }
    } 
    
}


// Function to extract ARNs and ASNs from the page's DOM
async function getARN(dom, sheet, pickupDate) {
    const document = dom.window.document;
    const katLabels = document.querySelectorAll("kat-label.kat-label-light-text");

    katLabels.forEach(katLabel => {
        let textContent = katLabel.getAttribute('text');
        if (textContent && textContent.includes(pickupDate)) {
            let labelId = katLabel.getAttribute('id');
            let parts = labelId.split("-");
            if (parts.length >= 5) {
                let firstId = parts[3];
                let secondId = parts[4];
                console.log(`Extracted ARN: ${firstId}, ASN: ${secondId}`);
                arnArr.push([firstId, secondId]);
            }
        }
    });
}

async function pressConfirmAndSubmitButton(driver) {
    try {
        // Locate the kat-button by its label
        let confirmButton = await driver.wait(until.elementLocated(By.css('kat-button[label="Confirm and submit shipment"]')), 15000);
        
        if (confirmButton) {

            // Click the button using WebElement.click()
            await confirmButton.click();
            console.log("Confirm and Submit Shipment button clicked using WebElement.click().");
        } else {
            console.log("Confirm and Submit Shipment button not found.");
        }
    } catch (error) {
        console.log("Error clicking the Confirm and Submit Shipment button: ", error);
    }
}


// Function to click buttons for "Continue to Step 2" and "Continue to Step 3"
async function continueToSteps(driver) {
    try {
        // Click on "Continue to step 2"
        let step2Button = await driver.executeScript(`
            return document.querySelector('kat-button[label="Continue to step 2"]');
        `);
        // Click on "Continue to step 3"
        let step3Button = await driver.executeScript(`
            return document.querySelector('kat-button[label="Continue to step 3"]');
        `);

        if (step2Button) {
            await driver.executeScript("arguments[0].click();", step2Button);
            console.log("Clicked Continue to step 2");
        }
        if (step3Button) {
            await driver.executeScript("arguments[0].click();", step3Button);
            console.log("Clicked Continue to step 3");
        }

        // Wait for page to load after clicking
        await driver.sleep(20);

    } catch (error) {
        console.log("Error clicking continue buttons: ", error);
    }
}


function hasWeekendsBetweenDates(startDateStr, endDateStr) {
    // Parse the input strings as UTC dates
    let startDate = new Date(`${startDateStr}T00:00:00Z`);  // Force UTC by appending 'T00:00:00Z'
    let endDate = new Date(`${endDateStr}T00:00:00Z`);      // Force UTC

    // Boolean variable to track if there is a weekend
    let hasWeekend = false;

    console.log(`Checking weekends between ${startDate.toUTCString()} and ${endDate.toUTCString()}`);

    // Loop through each day between the start and end dates (inclusive)
    while (startDate <= endDate) {
        let dayOfWeek = startDate.getUTCDay();  // getUTCDay() returns 0 for Sunday, 6 for Saturday

        console.log(`Checking date: ${startDate.toUTCString()}, Day of Week: ${dayOfWeek}`);

        // Check if the day is a weekend (Saturday or Sunday)
        if (dayOfWeek === 0 || dayOfWeek === 6) {
            hasWeekend = true;  // Set to true if a weekend is found
            console.log("Weekend found on: " + startDate.toUTCString());
            break;  // No need to check further if we already found a weekend
        }

        // Move to the next day in UTC
        startDate.setUTCDate(startDate.getUTCDate() + 1);
    }

    return hasWeekend;
}

// Main loop to scrape pages and navigate using the "Next" button
async function main() {
    try {
        var shipDate = "";
        let newDates = "";
        var pickupInpt = "";
        rl.question('Please enter the pickup example: "Pickup: Thu, Sep 19, 2024 CDT" ', (pickupInput) => {
            pickupInpt = pickupInput;
            rl.question('Please enter the ship date (MM/DD/YYYY): " ', (input) => {
                let inputs = input.split("/")
                shipDate = inputs[2] + "-" + inputs[0] + "-" + inputs[1];
                newDates = input.split('-');
                rl.close();
            });
        });
        
        // Use the existing Chrome instance to navigate to the URL
        await driver.get('https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/shippingqueue');

        // Extract ARNs and ASNs from the current page            
        await driver.sleep(10000);  // Wait for the page to load
        while (true) {
            const pageSource = await driver.getPageSource();
            const dom = new JSDOM(pageSource);
            
            
            await getARN(dom, outputSheet, pickupInpt);
            // Check the display style of the "sq-pag-next-div"
            const isNextButtonHidden = await driver.executeScript(`
                const nextDiv = document.querySelector('#sq-pag-next-div');
                return window.getComputedStyle(nextDiv).display === 'none';
            `);

            if (isNextButtonHidden) {
                console.log("Pages have come to an end.");
                break;
            }

            try {
                const nextButton = await driver.wait(until.elementLocated(By.xpath("//div[@id='sq-pag-next-div']//kat-label[@class='kat-label-link-text']//span[contains(text(), 'next >')]")), 10000);
                console.log("Next page found. Going to it...");

                await driver.executeScript("arguments[0].click();", nextButton);    
                await driver.sleep(2000);  // Wait for the page to load
            } catch (error) {
                console.log("No more pages.");
                break;
            }
        }

        // Process each ARN and ASN
        for (const arn of arnArr.slice(1)) {
            let trckNumbers = [];
            let amznLbls = [];
            let wrhs = "";
            
            await driver.get(`https://vendorcentral.amazon.com/kt/vendor/members/afi-shipment-mgr/asnsubmission?arn=${arn[0]}&asnId=${arn[1]}`);
            
            const image = await driver.wait(until.elementLocated(By.xpath("//img[@height='45']")), 10000);
            const newPageSource = await driver.getPageSource();
            const newDom = new JSDOM(newPageSource);

            // Get warehouse number (example using regex)
            const pattern = /^[A-Za-z0-9]{4},/;
            let katLinkElements = newDom.window.document.querySelectorAll("kat-link[slot='trigger']");
            if (katLinkElements.length > 1) {
                let labelValue = katLinkElements[1].getAttribute('label');
                if (pattern.test(labelValue)) {
                    wrhs = labelValue;
                    console.log(`Matching warehouse: ${labelValue}`);
                }
            }

            // Click "Continue to step 2" and "Continue to step 3"
            await continueToSteps(driver);

            // Change the date if necessary
            /*let matchingDate = shipDates.find(data => wrhs.split(',')[0] === data[0]);
            if (matchingDate) {
                await changeDate(driver, matchingDate[1]);  // Change the date
            } else {
                console.log(`Couldn't find shipdate for warehouse: ${wrhs}`);
            }*/

            // Get tracking numbers and labels
            let trackingElements = await driver.findElements(By.css('div[col-id="carrierTrackingNumber"]'));
            let amazonLabelElements = await driver.findElements(By.css('div[col-id="cartonLabelBarcode"]'));
            
            for (let i = 0; i < amazonLabelElements.length; i++) {
                let labelText = await amazonLabelElements[i].getText();
                if (labelText.startsWith("AMZN")) {
                    // Get tracking numbers
                    let actions = driver.actions();
                    await actions.doubleClick(trackingElements[i]).perform();
                    let richSelectRows = await driver.findElements(By.css('.ag-rich-select-row'));
                    if (richSelectRows.length > 1){
                        let rowText = await richSelectRows[i-1].getText(); 
                        if (rowText) {
                            console.log(rowText);
                            await richSelectRows[i-1].click();
                        }
                        // Print the details and append to output
                        console.log(arn[0], arn[1], labelText, rowText, wrhs);
                        outputSheet.push([arn[0], arn[1], labelText, rowText, wrhs]);
                    }
                }
                
            }

            let matchingDate = shipDates.find(data => wrhs.split(',')[0] === data[0]);
            if (matchingDate) {
                var deliveryDate = "";
                let addedDays = matchingDate[1];  // Extract the second element (data[1]) from the matching entry
                    // Pass the user input to the function  
                let newDate = (newDates[0] + '-' + newDates[1] + '-' + (parseInt(newDates[2])+addedDays));
            
                console.log(shipDate, newDate);
                let result = hasWeekendsBetweenDates(shipDate, newDate);
                if (result){
                    addedDays += 2;
                }
                // Click on "Continue to step 4"
                let step4Button = await driver.executeScript(`
                    return document.querySelector('kat-button[label="Continue to step 4"]');
                `);
                if (step4Button) {
                    await driver.executeScript("arguments[0].click();", step4Button);
                    console.log("Clicked Continue to step 4");
                }
                console.log(addedDays);
                let shipDateArray = shipDate.split("-");
                let deliveryDateArray = shipDate.split("-");
                
                shipDate = (shipDateArray[1] + "/" + shipDateArray[2] + "/" + shipDateArray[0])
                deliveryDate = (deliveryDateArray[1] + "/" + (parseInt(deliveryDateArray[2])+addedDays) + "/" + deliveryDateArray[0])
                await setDateInDatePicker(driver, shipDate);//
                
                await setDateInEDDDatePicker(driver, deliveryDate);
                await driver.sleep(3000);  // Wait for the next page to load

            } else {
                console.log(`Couldn't find shipdate for warehouse: ${wrhs}`);
            }
            
            /*
            try {
                const pageSource = await driver.getPageSource();
                const dom = new JSDOM(pageSource);
                const nextButton = await driver.wait(until.elementLocated(By.xpath('kat-button[label="Confirm and submit shipment"]')), 10000);
                console.log("Next button found. Clicking it...");
                await driver.executeScript("arguments[0].click();", nextButton);
                await driver.sleep(3000);  // Wait for the next page to load
            } catch (error) {
                console.log("Couldn't find submit button");
            }
            Ask user for pickup. example- Pickup: Tue, Sep 17, 2024 CDT
            Ask user for ship date    
            */
        }

        // Save workbook after all pages are processed
        XLSX.writeFile(outputWorkbook, 'ARN_ASN_Data.xlsx');
        console.log("Data saved to ARN_ASN_Data.xlsx");

    } finally {
        // Optionally, you can leave the Chrome session open
        // driver.quit();  // Comment this out if you want to keep Chrome open
    }
}

main();
