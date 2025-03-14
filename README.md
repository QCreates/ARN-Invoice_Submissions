This is a bot which submits the ASN's of items for your Vendor Central Account. Much work needs to be done still and there's a lot of dead code. For now both the shipment date of the item and way to find out when the ETA is, is manual. 

# Run the chrome driver:<br />
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222<br />
Make sure you're logged in to your amazon vendor central account.<br />
<br />
# Run this series of commands in VSCode terminal:<br />
npm init<br />
npm install selenium-webdriver<br />
npm install jsdom<br />
npm install xlsx<br />
<br />
# To run the code, run:<br />
node .\ASN_Submition.js\
