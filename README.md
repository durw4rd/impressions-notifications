# Optimizely Impression Notification App

A connector between Optimizely and Google Spreadsheets through which you can monitor the impression usage in Optimizely. You can set a daily threshold and get proactively notifed of any experiments that consume more than the threshold amount.

## Instructions for use

1. Create a copy of the Spreadsheet template located here (ADD LINK). In the copied document, head to the *Configuration* tab and enter the Optimizely **Account ID**.
2. In the main navigation of the copied document, select *Extensions > Apps Script* to open the Apps Script in a new tab. Head to the (Apps Script) *Project Settings* and copy the **Script ID**.
2. In Optimizely, head to *Account Settings > Registered Apps* and Register a new application.
3. Set the application name & redirect URI. The Redirect URI should be in the format of 'https://script.google.com/macros/d/SCRIPT_ID/usercallback' where you replace the 'SCRIPT_ID' with the value retrieved in Step 2.
4. Upon successful registration, the app will get assigned a **Client ID** and **Client Secret**. You will need these in the next steps to modify the provided Apps Script template.
5. Head back to the tab with the spreadsheet and authenticate the application: *Optimizely Menu > Authorize*.
6. Upon the successfull authentication, you should now see new options under the *Optimizely Menu* tab: 
  a. **List Experiments** will list experiments that have consumed more impressions than is the threshold set in the *Configuration* tab (default is 1000).
  b. **Log Out** will break the connection between the Apps Script and the linked Optimizely Account in case you need to switch Optimizely accounts connected to your script.

## Editing values in the Configuration tab

This tab is where you can customize the filters used by the script. Namely, you can modify the following properties:

1. **Email** (optional): email address to which a notification will be sent, listing the experiments that consumed impressions above the set threshold.
2. **Date Range:** Controls the number of days included in the query by updating the *startDate* of the request. The number entered here is used to deduct the number of days from the current date. I.e. the default value (2) will run a query ranging from the day before yesterday (today - 2 days) until today.
  a. Note the job that generates the impression data runs once a day and generally is expected to finish by **9 pm GMT of the following day** (i.e. impressions numbers for 1/1/2022 will become available for queriying at 9pm GMT of 2/1/2022).
3. **Impressions Threshold:** Set the threshold for including individual experiments in the report. This threshold is for the entire duration of the query.
4. **Optimizely Account ID:** The account ID of your Optimizely account. You can find it under *Account Settings > Plan*.