// If not authorized yet, create menu item for starting the authorization flow.
// Else, create other menu items.
function getMenu(ui) {
  if (!getOptiService().hasAccess()) {
    ui.createMenu('Optimizely Menu')
    .addItem('Authorize', 'openAuthScreen')
    .addToUi();
  } else {
    ui.createMenu('Optimizely Menu')
    .addItem('List Experiments', 'printToSheet')
    .addItem('Send Data via Email', 'sendEmail')
    .addItem('Log Out', 'logout')
    .addToUi();
  }
}

// Retrieve the Apps Script ID
function getScriptID(){ return ScriptApp.getScriptId() }

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  getMenu(ui);

  getScriptID();
}

// Using the following library: https://github.com/googleworkspace/apps-script-oauth2
function getOptiService() {
  var clientId = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B8").getValue().toString();
  var clientSecret = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B9").getValue(); 

  return OAuth2.createService('OptimizelyNotifier')
    .setAuthorizationBaseUrl('https://app.optimizely.com/oauth2/authorize')
    .setTokenUrl('https://app.optimizely.com/oauth2/token')

    .setClientId(clientId)
    .setClientSecret(clientSecret)

    // Set the name of the callback function in the script referenced
    // above that should be invoked to complete the OAuth flow.
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('all')
    .setParam('response_type', 'code')
}

function openAuthScreen(){
  var ui = SpreadsheetApp.getUi();
  if (!getOptiService().hasAccess()) {
    var authorizationUrl = getOptiService().getAuthorizationUrl();
    var html = HtmlService.createHtmlOutput('<!DOCTYPE html><html><script>'
    +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
    +'var a = document.createElement("a"); a.href="'+authorizationUrl+'"; a.target="_blank";'
    +'if(document.createEvent){'
    +'  var event=document.createEvent("MouseEvents");'
    +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
    +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
    +'}else{ a.click() }'
    +'close();'
    +'</script>'
    // Offer URL as clickable link in case above code fails.
    +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically.<br/><a href="'+authorizationUrl+'" target="_blank" onclick="window.close()">Click here to proceed to the Optimizely auth dialog</a>.</body>'
    +'<script>google.script.host.setHeight(55);google.script.host.setWidth(410)</script>'
    +'</html>')
    .setWidth( 90 ).setHeight( 1 );
    ui.showModalDialog(html, "Opening Optimizely Authorization Dialog...");
  } else {
    getMenu(ui);
    // TODO: show a warning that Optimizely cannot connect
  }
}

function authCallback(request) {
  var ui = SpreadsheetApp.getUi();
  var isAuthorized = getOptiService().handleCallback(request);
  if (isAuthorized) {
    getMenu(ui);
    return HtmlService.createHtmlOutput('<!DOCTYPE html><html>'
    +'<div style="margin: auto;">'
    +'<h2 style="font-family: Inter,sans-serif; text-align: center;">Authorization Successful. You can close this tab.</h2>'
    +'<button onclick="window.top.close()" style="background-color: #0037ff; border-color: #1a4bff; color: #fff; display: inline-block; vertical-align: middle; white-space: nowrap; font-family: Inter,sans-serif; cursor: pointer; line-height: 32px; border-width: 1px; border-style: solid; font-size: 13px; font-weight: 400; border-radius: 4px; height: 34px; padding: 0 15px; display: flex; margin: auto;">Close'
    +'</button></div></html>');
  } else {
    getMenu(ui);
    return HtmlService.createHtmlOutput('Access Denied. Double-check that you have provided valid credentials. (You can close this tab)');
    // TODO: make this a bit more presentable
  }
}

function logout() {
  var ui = SpreadsheetApp.getUi();
  var service = getOptiService();
  service.reset();
  getMenu(ui);
}