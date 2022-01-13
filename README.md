# Optimizely Impression Notification App

A connector between Optimizely and Google Spreadsheets through which you can monitor the impression usage in Optimizely. You can set a daily threshold and get proactively notifed of any experiments that consume more than the threshold amount.

## Instructions for use

1. Create a copy of the Spreadsheet template located [here](https://docs.google.com/spreadsheets/d/1g97TBPWWYl0kMm-UjTfvNqNnJjm7IS_S6Vm-dMDUuO0/edit?usp=sharing).
2. In Optimizely, head to *Account Settings > Registered Apps* and Register a new application. Enter the application name & the **Redirect URI**. You can find the Redirect URI in the *Configuration* tab in the cell B12. The URI should be in the following format 'https://script.google.com/macros/d/SCRIPT_ID/usercallback'. Set the **Client Type** as *Confidential*.
3. Upon successful registration, the app will get assigned a **Client ID** and **Client Secret**. Copy these values into the corresponding fields on the *Configuration* tab of the spreadsheet (B9 & B10).
4. Authenticate the application: *Optimizely Menu > Authorize*.
5. Upon the successfull authentication, you should now see new options under the *Optimizely Menu* tab: 
    1. **List Experiments** will list experiments that have consumed more impressions than is the threshold set in the *Configuration* tab. If the email notification is enabled (via the corresponding field in the *Configuration* tab) this function will also send the email notification.
    2. **Log Out** will break the connection between the Apps Script and the linked Optimizely Account in case you need to switch Optimizely accounts connected to your script.
6. Adjust the properties of the script via the *Configuration* tab (details for all fields in the following section).
7. If desired, set a script trigger to execute the script automatically on a regular basis. This can be done under *Extensions > Apps Script > Triggers > Add Trigger*.
    1. When setting up an automated trigger, select **listExperimentImpressions** as the function to run.
    2. If you want the script to be executing daily and retrieve impression numbers that will include the previous day, the script should execute after 9:00 pm GMT.

## Editing values in the Configuration tab

This tab is where you can customize the filters used by the script. Namely, you can modify the following properties:

1. **Optimizely Account ID:** The account ID of your Optimizely account. You can find it under *Account Settings > Plan*. 
2. **Impressions Threshold:** Set the threshold for including individual experiments in the report. This threshold is for the entire duration of the query.
3. **Date Range:** Controls the number of days included in the query by updating the *startDate* of the request. The number entered here is used to deduct the number of days from the current date. I.e. the default value (2) will run a query ranging from the day before yesterday (today - 2 days) until today.
    1. Note the job that generates the impression data runs once a day and generally is expected to finish by **9 pm GMT of the following day** (i.e. impressions numbers for 1/1/2022 will become available at 9pm GMT on 2/1/2022).
4. **Send Notification Emails:** If you'd like to be receiving emails, set this to 'Y'.
5. **Email:** Enter the email address where the automatic notificaton with experiments above the impression threshold should be sent. You can add multiple email addresses, separated by a comma.