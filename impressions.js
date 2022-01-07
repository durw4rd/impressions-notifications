// Check if the data is not stale
function getImpressionSummary() {
  var acountID = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange(4, 2).getValue();
  var url = 'https://api.optimizely.com/v2/billing/usage/' + acountID + '/summary';
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
  // Get Account ID
  var acountID = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange(4, 2).getValue();
  // Get Weekly Impression Limit
  var impressionLimit = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange(3, 2).getValue();
  // Today's date
  var endDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd")
  // Start date of the query. 
  // Can be adjusted in the 'Configuration' tab. Using 2 days by default (yesterday's data becomes available ~1pm PST/9pm GMT).
  var daysAgo = parseInt(SpreadsheetApp.getActive().getSheetByName('Configuration').getRange(2, 2).getValue() || 2);
  var startDateRaw = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - daysAgo, new Date().getHours()); // NOTE: The time will be in local timezone
  var startDate = Utilities.formatDate(startDateRaw, "GMT", "yyyy-MM-dd"); // NOTE: GMT is used here. Mind that the timezone offset can put you into a different date.
  
  // Read the impressions details for all experiments
  // TODO: handle pagination
  var url = 'https://api.optimizely.com/v2/billing/usage/' + acountID + '?usage_date_from=' + startDate + '&usage_date_to=' + endDate;
  Logger.log(url);
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'authorization': 'Bearer ' + getOptiService().getAccessToken(),
    }
  });

  var data = JSON.parse(response);
  var experiments = [];
  var experimentsAboveLimit = [];

  data.forEach(function (elem) {
    experiments.push(elem);
  });

  Logger.log(experiments);

  for (i in experiments) {
    if (experiments[i].impression_count > impressionLimit) {
      experimentsAboveLimit.push([experiments[i].project_name, experiments[i].experiment_id, experiments[i].experiment_name, experiments[i].experiment_status, experiments[i].platform, experiments[i].impression_count])
    }
  }

  // Print experiments above the threshold in the sheet
  if (experimentsAboveLimit.length > 0) {
    Logger.log(experimentsAboveLimit);
    var lastRow = SpreadsheetApp.getActive().getSheetByName('Results').getLastRow();
    SpreadsheetApp.getActive().getSheetByName('Results')
      .getRange(lastRow + 1, 1, experimentsAboveLimit.length, experimentsAboveLimit[0].length)
      .setValues(experimentsAboveLimit);
    // notifyUser(experimentsAboveLimit);
  }

  // Check last impressions usage update date vs. the current date
  getImpressionSummary();
}