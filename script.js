function getMenu(ui) {
  if (!getOptiService().hasAccess()) {
    ui.createMenu('Optimizely Menu')
    .addItem('Authorize', 'showSidebar')
    .addToUi();
  } else {
    ui.createMenu('Optimizely Menu')
    .addItem('Get Projects', 'getProjects')
    .addItem('Print Running experiments', 'printRunningExperimentResults')
    .addItem('Get Impressions', 'getImpressions')
    .addToUi();
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  getMenu(ui);
}

function getOptiService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth2.createService('opti')

    // Set the endpoint URLs for Optimizely 2.0.
    .setAuthorizationBaseUrl('https://app.optimizely.com/oauth2/authorize')
    .setTokenUrl('https://app.optimizely.com/oauth2/token')

    /*
     * NEED TO CHANGE 
     * Set the client ID and secret, from the Optimizely Account Settings Page.
     */
    .setClientId('21007060557')
    .setClientSecret('pxHi0N9_9lMcF2Y7izoyM8uB-4zyhvFXSfuLauN3M3s')

    // Set the name of the callback function in the script referenced
    // above that should be invoked to complete the OAuth flow.
    .setCallbackFunction('authCallback')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties())

    // Set the scopes to request (space-separated for Google services).
    .setScope('all')

    // Below OAuth2 parameters.

    // Response type
    .setParam('response_type', 'code')

  // Account ID
  //.setParam('account_id', 7954810296);

}

function showSidebar() {
  var ui = SpreadsheetApp.getUi();
  if (!getOptiService().hasAccess()) {
    var authorizationUrl = getOptiService().getAuthorizationUrl();
    var template = HtmlService.createTemplate(
      '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
      'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    getMenu(ui);
  }
}

function authCallback(request) {
  var ui = SpreadsheetApp.getUi();
  var isAuthorized = getOptiService().handleCallback(request);
  if (isAuthorized) {
    getMenu(ui);
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    getMenu(ui);
    return HtmlService.createHtmlOutput('Access Denied. Double-check that you have provided valid credentials.');
  }
}

// function getProjects() {
//   var url = 'https://api.optimizely.com/v2/projects';
//   var response = UrlFetchApp.fetch(url, {
//     headers: {
//       'authorization': 'Bearer ' + getOptiService().getAccessToken(),
//     }
//   });

//   var data = JSON.parse(response);
//   var project = [];
//   data.forEach(function (elem) {
//     project.push([elem["name"], elem["id"]]);
//   });

//   Logger.log(project);

//   var template = HtmlService.createTemplate(project);
//   var page = template.evaluate();
//   SpreadsheetApp.getUi().showSidebar(page);
//   Logger.log('get project: ' + response);
// }

function getImpressions() {
  // Get Account ID:
  var acountID = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange(4, 2).getValue();
  Logger.log(acountID);
  // Get Weekly Impression Limit:
  var limit = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange(3, 2).getValue();
  Logger.log(limit);

  var endDateRaw = new Date();
  var endDate = Utilities.formatDate(endDateRaw, "GMT", "yyyy-MM-dd")
  // Read the date window from the sheet. If no value provided, use the default 7-day window.
  var numberOfDays = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange(2, 2).getValue() - 1 || 6;
  var startDateRaw = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - numberOfDays);
  var startDate = Utilities.formatDate(startDateRaw, "GMT", "yyyy-MM-dd")
  
  // Read (old) impression numbers from the sheet
  var numberOfRows = SpreadsheetApp.getActive().getSheetByName('Results').getLastRow() - 2;
  // var lastRow = SpreadsheetApp.getActive().getSheetByName('Results').getLastRow();
  var oldImpressions = SpreadsheetApp.getActive().getSheetByName('Results').getRange(2, 6, numberOfRows).getValues();
  Logger.log('OLD IMPRESSIONS');
  Logger.log(oldImpressions);
  
  // Call the billing api
  var url = 'https://api.optimizely.com/v2/billing/usage/' + acountID + '?usage_date_from=' + startDate + '&usage_date_to=' + endDate;
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'authorization': 'Bearer ' + getOptiService().getAccessToken(),
    }
  });
  var data = JSON.parse(response);
  //Object containing metrics
  var experiments = [];
  var experimentsAboveLimit = [];
  var newImpressions = [];
  var project_name
  var experiment_id
  var experiment_name
  var experiment_status
  var platform
  var impression_count

  data.forEach(function (elem) {
    experiments.push(elem);
  });

  for (i in experiments) {
    // TODO move this into the push method (to save 12 rows)
    project_name = experiments[i].project_name;
    experiment_id = experiments[i].experiment_id;
    experiment_name = experiments[i].experiment_name;
    experiment_status = experiments[i].experiment_status;
    platform = experiments[i].platform;
    impression_count = experiments[i].impression_count;
    
    newImpressions.push(experiments[i].impression_count);

    if (impression_count > limit) {
      experimentsAboveLimit.push([project_name, experiment_id, experiment_name, experiment_status, platform, impression_count])
    }
    //Append Row
    SpreadsheetApp.getActive().getSheetByName('Results').appendRow([project_name, experiment_id, experiment_name, experiment_status, platform, impression_count]);
  }

  Logger.log('NEW IMPRESSIONS');
  Logger.log(newImpressions);

  Logger.log(oldImpressions - newImpressions);

  if (experimentsAboveLimit.length > 0) {
    notifyUser(experimentsAboveLimit);
  }
}

// TODO This one needs a bit of work
function notifyUser(arrayOfExperimentsAboveLimit) {
  Logger.log(arrayOfExperimentsAboveLimit);
}

// function clearSheet(sheetName, rowStart) {
//   var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
//   var rowEnd = sheet.getLastRow() - 1;//Number of last row with content
//   sheet.deleteRows(rowStart, rowEnd);
// }

function getResults(exp_id) {
  var optiServices = getOptiService();
  var url = 'https://api.optimizely.com/v2/experiments/' + exp_id + '/results';
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'authorization': 'Bearer ' + optiServices.getAccessToken(),
    }
  });
  var data = JSON.parse(response);

  //Object containing metrics
  var metrics = [];
  data["metrics"].forEach(function (elem) {
    metrics.push(elem);
  });

  for (i in metrics) {
    for (j in metrics[i].results) {
      var experiment_id = data.experiment_id
      var dt = new Date(getDateFromIso(data.start_time));
      var start_time = Date.parse(data.start_time)
      var metric_name = metrics[i].name
      var variation_name = metrics[i].results[j].name
      var conversion_rate = metrics[i].results[j].rate ? 'yes' : 'no'
      var reached_significance = metrics[i].results[j].is_baseline ? ' - ' : metrics[i].results[j].lift.is_significant
      var lift = metrics[i].results[j].is_baseline ? ' - ' : metrics[i].results[j].lift.value;
      //Append Row
      SpreadsheetApp.getActiveSheet().appendRow([experiment_id, dt, metric_name, variation_name, conversion_rate, lift, reached_significance, data.start_time]);
    }
  }
  Logger.log(data);
}

function getRunningExperiments() {
  Logger.log("getRunningExperiments is running");
  var optiServices = getOptiService();
  var complete = false;
  var page = 1;
  var projectId = SpreadsheetApp.getActiveSheet().getRange(2, 2).getValue();
  Logger.log("Project ID: " + projectId);
  var runningExperiments = [];

  while (!complete) {
    var url = 'https://api.optimizely.com/v2/experiments?project_id=' + projectId + '&per_page=50&page=' + page;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'authorization': 'Bearer ' + optiServices.getAccessToken(),
      }
    });
    var data = JSON.parse(response);

    //Object containing metrics
    var experiments = [];
    data.forEach(function (elem) {
      experiments.push(elem);
    });

    for (i in experiments) {
      if (experiments[i].status == 'running') {
        runningExperiments.push(experiments[i].id);
      }
    }
    page++;
    if (data == '') {
      complete = true;
    }
  }
  Logger.log("Running experiments:");
  Logger.log(runningExperiments);
  return runningExperiments;
}

function printRunningExperimentResults() {
  var runningExp = getRunningExperiments();
  var sheet = SpreadsheetApp.getActiveSheet()
  var end = sheet.getLastRow() - 1;//Number of last row with content
  //blank rows after last row with content will not be deleted
  sheet.deleteRows(3, end);
  SpreadsheetApp.getActiveSheet().appendRow(['Experiment ID', 'Start Date', 'Metric Name', 'Variation name', 'Conversion rate', 'lift', 'Is Significant']);
  for (i in runningExp) {
    getResults(runningExp[i]);
  }
}

function getDateFromIso(string) {
  try {
    var aDate = new Date();
    var regexp = "([0-9]{4})(-([0-9]{2})(-([0-9]{2})" +
      "(T([0-9]{2}):([0-9]{2})(:([0-9]{2})(\\.([0-9]+))?)?" +
      "(Z|(([-+])([0-9]{2}):([0-9]{2})))?)?)?)?";
    var d = string.match(new RegExp(regexp));

    var offset = 0;
    var date = new Date(d[1], 0, 1);

    if (d[3]) { date.setMonth(d[3] - 1); }
    if (d[5]) { date.setDate(d[5]); }
    if (d[7]) { date.setHours(d[7]); }
    if (d[8]) { date.setMinutes(d[8]); }
    if (d[10]) { date.setSeconds(d[10]); }
    if (d[12]) { date.setMilliseconds(Number("0." + d[12]) * 1000); }
    if (d[14]) {
      offset = (Number(d[16]) * 60) + Number(d[17]);
      offset *= ((d[15] == '-') ? 1 : -1);
    }

    offset -= date.getTimezoneOffset();
    time = (Number(date) + (offset * 60 * 1000));
    return aDate.setTime(Number(time));
  } catch (e) {
    return;
  }
}