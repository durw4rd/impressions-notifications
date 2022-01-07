// If not authorized yet, create menu item for starting the authorization flow.
// Else, create other menu items.
function getMenu(ui) {
  if (!getOptiService().hasAccess()) {
    ui.createMenu('Optimizely Menu')
    .addItem('Authorize', 'showSidebar')
    .addToUi();
  } else {
    ui.createMenu('Optimizely Menu')
    .addItem('List Experiments', 'listExperimentImpressions')
    .addItem('Log Out', 'logout')
    .addToUi();
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  getMenu(ui);
}

// Using the following library: https://github.com/googleworkspace/apps-script-oauth2
function getOptiService() {
  return OAuth2.createService('OptimizelyNotifier')
    .setAuthorizationBaseUrl('https://app.optimizely.com/oauth2/authorize')
    .setTokenUrl('https://app.optimizely.com/oauth2/token')

    /*
     * NEED TO CHANGE 
     * Set the client ID and secret, from the Optimizely Account Settings Page.
     */
    .setClientId('')
    .setClientSecret('')

    // Set the name of the callback function in the script referenced
    // above that should be invoked to complete the OAuth flow.
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('all')
    .setParam('response_type', 'code')
}

// Apps Script UI's are not allowed to redirect the user's window to a new URL, 
// so you'll need to present the authorization URL as a link for the user to click.
function showSidebar() {
  var ui = SpreadsheetApp.getUi();
  if (!getOptiService().hasAccess()) {
    var authorizationUrl = getOptiService().getAuthorizationUrl();
    var template = HtmlService.createTemplate(
      '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. \n' +
      'You can close the sidebar when the authorization is complete.');
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
    return HtmlService.createHtmlOutput('Access Denied. Double-check that you have provided valid credentials. (You can close this tab)');
  }
}

function logout() {
  var ui = SpreadsheetApp.getUi();
  var service = getOptiService();
  service.reset();
  getMenu(ui);
}