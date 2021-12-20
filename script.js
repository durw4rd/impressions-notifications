function getMenu(ui) {
  if (!getOptiService().hasAccess()) {
    ui.createMenu('Optimizely Menu')
    .addItem('Authorize', 'showSidebar')
    .addToUi();
  } else {
    ui.createMenu('Optimizely Menu')
    .addItem('Get Projects', 'getProject')
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
    .setClientId('14086940272')
    .setClientSecret('ysO5KOCnlTXrGJknmbI_vHhABo2i8Qe0OHg9hAsgxt4')

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
  var optiServices = getOptiService();
  if (!optiServices.hasAccess()) {
    var authorizationUrl = optiServices.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
      '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
      'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    getFullMenu(ui);
  }
}

function authCallback(request) {
  var ui = SpreadsheetApp.getUi();
  var optiServices = getOptiService();
  var isAuthorized = optiServices.handleCallback(request);
  if (isAuthorized) {
    getAuthMenu(ui);
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    getFullMenu(ui);
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

function getProject() {
  var optiServices = getOptiService();

  var url = 'https://api.optimizely.com/v2/projects';
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'authorization': 'Bearer ' + optiServices.getAccessToken(),
    }
  });

  var data = JSON.parse(response);
  var project = [];
  data.forEach(function (elem) {
    project.push([elem["name"], elem["id"]]);
  });

  Logger.log(project);

  var template = HtmlService.createTemplate(project);
  var page = template.evaluate();
  SpreadsheetApp.getUi().showSidebar(page);
  Logger.log('get project: ' + response);
}

function sevenDaysAgo() {
  return new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - 6);
}

function getImpressions() {
  // Get Account ID:
  var acountID = SpreadsheetApp.getActiveSheet().getRange(2, 2).getValue();
  Logger.log(acountID);
  // Get Weekly Impression Limit:
  var limit = SpreadsheetApp.getActiveSheet().getRange(2, 4).getValue();
  Logger.log(limit);

  // Get the 7 day window dates 
  todaytemp = new Date();
  today = Utilities.formatDate(todaytemp, "GMT", "yyyy-MM-dd")
  Logger.log(today)
  var tempyesterday = sevenDaysAgo();
  yesterday = Utilities.formatDate(tempyesterday, "GMT", "yyyy-MM-dd")
  Logger.log(yesterday)
  // Call the billing api
  var optiServices = getOptiService();
  var url = 'https://api.optimizely.com/v2/billing/usage/' + acountID + '?usage_date_from=' + yesterday + '&usage_date_to=' + today;
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'authorization': 'Bearer ' + optiServices.getAccessToken(),
    }
  });
  var data = JSON.parse(response);
  //Object containing metrics
  var experiments = [];
  var experimentsAboveLimit = [];

  data.forEach(function (elem) {
    experiments.push(elem);
  });

  for (i in experiments) {
    var project_name = experiments[i].project_name;
    var experiment_id = experiments[i].experiment_id;
    var experiment_name = experiments[i].experiment_name;
    var experiment_status = experiments[i].experiment_status;
    var platform = experiments[i].platform;
    var impression_count = experiments[i].impression_count;

    if (impression_count > limit) {
      experimentsAboveLimit.push([project_name, experiment_id, experiment_name, experiment_status, platform, impression_count])
    }
    //Append Row
    SpreadsheetApp.getActiveSheet().appendRow([project_name, experiment_id, experiment_name, experiment_status, platform, impression_count]);
  }

  if (experimentsAboveLimit.length > 0) {
    notifyUser(experimentsAboveLimit);
  }
}

function notifyUser(arrayOfExperimentsAboveLimit) {
  Logger.log(arrayOfExperimentsAboveLimit);
}

function clearSheet() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var end = sheet.getLastRow() - 1;//Number of last row with content
  //blank rows after last row with content will not be deleted
  sheet.deleteRows(3, end);
}

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