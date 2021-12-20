# Optimizely Impression Notification App

A connector between Optimizely and Google Spreadsheets through which you can monitor the impression usage in Optimizely. You can set a daily threshold and get proactively notifed of any experiments that consume more than the threshold amount.

## Instructions for use

1. In Optimizely, head to *Account Settings > Registered Apps* and Register a new application.
2. Set the application name & Redirect URI (which one to use?). Upon successful registration, the app will get assigned a Client ID and Client Secret. You will need these later on to authenticate the requests made by the Google Apps Script to Optimizely REST API.
3. Create a copy of the Spreadsheet template located here (add link). In the main navigation of the copied document, select *Extensions > Apps Script* to open the script in a new tab.
4. Find the lines where the Client ID and Client Secret are set (41-42 for now) and enter the values assigned to your application from steps 1-2. Save your changes (Cmd/Ctrl + S).
5. Head back to the tab with the spreadsheet and authenticate the application: *Optimizely Menu > Authorize*.
6.