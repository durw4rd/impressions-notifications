// Read configuration details from the spreadsheet
var acountId = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B5").getValue();
var impressionLimit = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B4").getValue() || 1000;
var dateRange = parseInt(SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B3").getValue() || 2);

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
    Logger.log(url);
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
      if (experimentsOnPage[i].impression_count > impressionLimit) {
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
    var oldExperimentData = SpreadsheetApp.getActive().getSheetByName('Impressions').getRange(4, 1, lastRow - 2,  6);
    oldExperimentData.clear({contentsOnly: true});
    SpreadsheetApp.getActive().getSheetByName('Impressions')
      .getRange(4, 1, experimentsAboveLimit.length, experimentsAboveLimit[0].length)
      .setValues(experimentsAboveLimit);
  }

  // Check last impressions usage update date vs. the current date
  getImpressionSummary();
}