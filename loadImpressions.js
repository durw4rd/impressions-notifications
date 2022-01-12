// Read configuration details from the spreadsheet
var acountId = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B2").getValue();
var impressionThreshold = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B3").getValue() || 1000;
var dateRange = parseInt(SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B4").getValue() || 2);
var notificationsEnabled = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B6").getValue();

// Check if the data is not stale
function getImpressionSummary() {
  
  var url = 'https://api.optimizely.com/v2/billing/usage/' + acountId + '/summary';
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'authorization': 'Bearer ' + getOptiService().getAccessToken(),
    }
  });
  var data = JSON.parse(response);
  var executionDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  var lastUpdateDate = data.last_update_date;
  SpreadsheetApp.getActive().getSheetByName('Log').appendRow([executionDate, lastUpdateDate]);
}

// Core function: lists all experiments that consumed impressions above the defined threshold
function listExperimentImpressions() {
  var endDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd") // Today's date
  /* 
  * Start date of the query. 
  * Can be adjusted in the 'Configuration' tab. 
  * Using 2 days by default (yesterday's data becomes available ~1pm PST/9pm GMT).
  */ 
  var startDateRaw = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - dateRange, new Date().getHours()); // NOTE: The time will be in local timezone.
  var startDate = Utilities.formatDate(startDateRaw, "GMT", "yyyy-MM-dd"); // NOTE: GMT is used here. Mind that the timezone offset can put you into a different day compared to the raw date.
  
  // Read the impressions details for all experiments
  var complete = false;
  var page = 1;
  var experimentsAboveLimit = [];
  while (!complete) {
    var url = 'https://api.optimizely.com/v2/billing/usage/' + acountId + '?usage_date_from=' + startDate + '&usage_date_to=' + endDate + '&per_page=50' + '&page=' + page;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'authorization': 'Bearer ' + getOptiService().getAccessToken(),
      }
    });

    var data = JSON.parse(response);
    var experimentsOnPage = [];
    data.forEach(function (elem) {
      experimentsOnPage.push(elem);
    });

    for (i in experimentsOnPage) {
      if (experimentsOnPage[i].impression_count > impressionThreshold) {
        experimentsAboveLimit.push([experimentsOnPage[i].project_name, experimentsOnPage[i].experiment_id, experimentsOnPage[i].experiment_name, experimentsOnPage[i].experiment_status, experimentsOnPage[i].platform, experimentsOnPage[i].impression_count]);
      } 
    }
    page++;
    if (data == '') {
      complete = true;
    }
  }

  // Print the date range of the executed query
  SpreadsheetApp.getActive().getSheetByName('Impressions')
    .getRange("A2")
    .setValue("Query Date Range: " + startDate + " to " + endDate);

  // Clear old experiment data & print new experiments above the selected threshold in the sheet
  if (experimentsAboveLimit.length > 0) {
    var lastRow = SpreadsheetApp.getActive().getSheetByName('Impressions').getLastRow();
    var oldExperimentData = SpreadsheetApp.getActive().getSheetByName('Impressions').getRange(4, 1, lastRow - 3,  6);
    oldExperimentData.clear({contentsOnly: true});
    SpreadsheetApp.getActive().getSheetByName('Impressions')
      .getRange(4, 1, experimentsAboveLimit.length, experimentsAboveLimit[0].length)
      .setValues(experimentsAboveLimit);

    // Send notification email
    if (notificationsEnabled === 'Y') {
      sendEmailNotification();
    }
  } else {
    var htmlModal = HtmlService.createHtmlOutput("No experiments above the threshold in the selected date range! <br>"
    +'Impressions Threshold: ' + impressionThreshold + '<br>'
    +'Query Date Range: ' + startDate + ' to ' + endDate + '<br><br>'
    +'<button onclick="google.script.host.close()" style="background-color: #0037ff; border-color: #1a4bff; color: #fff; display: inline-block; vertical-align: middle; white-space: nowrap; font-family: Inter,sans-serif; cursor: pointer; line-height: 32px; border-width: 1px; border-style: solid; font-size: 13px; font-weight: 400; border-radius: 4px; height: 34px; padding: 0 15px; display: flex; margin: auto;">Close'
    +'</button></div></html>'
    );
    SpreadsheetApp.getUi().showModalDialog(htmlModal, "Something's off");
  }

  // Check last impressions usage update date vs. the current date
  getImpressionSummary();
}