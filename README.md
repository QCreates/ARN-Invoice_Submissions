This is a bot which submits the ASN's and Invoices of items for your Vendor Central Account. 

# ALL COMMANDS BASED ON POWERSHELL<br />
MAKE SURE TO INSTALL CORRESPONDING VERSION OF CHROMEDRIVE TO YOUR CHROME BROWSER USING BELOW LINK:<br />
https://googlechromelabs.github.io/chrome-for-testing/#stable<br />
CHROMEDRIVER SHOULD BE LOCATED HERE: “C:/chromedriver/chromedriver.exe”<br />
ENV VARIABLE SHOULD POINT TO: “C:/chromedriver”<br />
<br />
-Activate the environment after navigating to the Python Directory<br />
venv\Scripts\activate<br />
<br />
-If not installed, install (Use pip list to check):<br />
pip install playwright<br />
playwright install<br />
pip install ipython<br />
pip install requests<br />
pip install pandas<br />
pip install openpyxl<br />
<br />

# How to run <br />
<br />
-Close all chrome tabs with task manager then run (Remove ‘&’ for cmd):<br />
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome-debug"<br />

# For ASN Submissions: <br />
Navigate to https://vendorcentral.amazon.com and login.<br />
Run “./python ASNBot.py” and input the pickup date in dd/mm/yyyy format.<br />
Wait until complete and the status of each submission will be updated on the excel sheet.<br />

# For Invoice Submissions: <br />
Make sure that you update invoices.xlsx with the right invoices. Date's may need to be adjusted.<br />
Navigate to https://vendorcentral.amazon.com/hz/vendor/members/invoice-creation/search-shipments and make sure your logged in.<br />
Click “Purchase Order Number(s)” in the second dropdown<br />
Run “python ./InvoiceSubmissionBot.py”<br />
If it doesn’t work, double-check that your date is the same as shown on the purchase order. Make sure to repeat step 1 if you restart the code.<br />
