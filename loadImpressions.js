configSheet = SpreadsheetApp.getActive().getSheetByName('Configuration');
dataSheet = SpreadsheetApp.getActive().getSheetByName('Impressions');

// Read configuration details from the spreadsheet
var acountId = configSheet.getRange("B2").getValue();
var impressionThreshold = configSheet.getRange("B3").getValue() || 1000;
var dateRange = parseInt(configSheet.getRange("B4").getValue() || 2);
var endDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd") // Today's date
/* 
* Start date of the query. 
* Can be adjusted in the 'Configuration' tab. 
* Using 2 days by default (yesterday's data becomes available ~1pm PST/9pm GMT).
*/ 
var startDateRaw = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - dateRange, new Date().getHours()); // NOTE: The time will be in local timezone.
var startDate = Utilities.formatDate(startDateRaw, "GMT", "yyyy-MM-dd"); // NOTE: GMT is used here. Mind that the timezone offset can put you into a different day compared to the raw date.

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
function loadImpressionData() {
   
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
    Logger.log(response.getResponseCode());
    var data = JSON.parse(response);
    Logger.log(data);

    var experimentsOnPage = [];
    data.forEach(function (elem) {
      experimentsOnPage.push(elem);
    });

    function createEmulationLink(id) {
      return '=HYPERLINK("https://app.optimizely.com/s/redirect-to-emulate?input=' + id + '&version=2","' + id + '")'
    }

    for (i in experimentsOnPage) {
      if (experimentsOnPage[i].impression_count > impressionThreshold) {
        experimentsAboveLimit.push([
          experimentsOnPage[i].project_name, 
          createEmulationLink(experimentsOnPage[i].experiment_id.toString()), 
          experimentsOnPage[i].experiment_name, 
          experimentsOnPage[i].experiment_status, 
          experimentsOnPage[i].platform, 
          experimentsOnPage[i].impression_count
        ]);
      } 
    }
    page++;
    if (data == '') {
      complete = true;
    }
  }
  
  return experimentsAboveLimit;
}

function printToSheet() {
  var dataToPrint = loadImpressionData();

  // Print the date range of the executed query in the sheet
  dataSheet
    .insertRowsBefore(3,2)
    .getRange(3,1,1,6)
    .mergeAcross()
    .setValue("Query Date Range: from " + startDate + " to " + endDate);
  dataSheet
    .getRange(4,1,1,6)
    .clearFormat();

  // Print new experiments above the selected threshold in the sheet
  if (dataToPrint.length > 0) {
    var lastRow = dataSheet.getLastRow();
    dataSheet
      .insertRowsBefore(4, dataToPrint.length)
      .getRange(4, 1, dataToPrint.length, dataToPrint[0].length)
      .setValues(dataToPrint);
  } else { // No experiments above the threshold returned
    dataSheet
      .insertRowBefore(4)
      .getRange(4,1,1,6)
      .mergeAcross()
      .setValue("No experiments consumed more than the set threshold of " + impressionThreshold + " impressions in the selected date range.");
  }

  // Check last impressions usage update date vs. the current date
  getImpressionSummary();
}