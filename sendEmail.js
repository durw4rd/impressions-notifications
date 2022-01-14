var impressionThreshold = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B3").getValue() || 1000;
var emailAddress = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B7").getValue();

var lastRow = SpreadsheetApp.getActive().getSheetByName('Impressions').getLastRow();
var expData = SpreadsheetApp.getActive().getSheetByName('Impressions').getRange(3, 1, lastRow - 2,  6).getValues();

var endDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
var startDateRaw = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - dateRange, new Date().getHours());
var startDate = Utilities.formatDate(startDateRaw, "GMT", "yyyy-MM-dd");

var spreadsheetURL = SpreadsheetApp.getActive().getUrl();
var linkTemplate = HtmlService.createTemplate("<a href='<?= spreadsheetURL ?>'>connected Google Sheet</a>");
var spreadheetLink = linkTemplate.evaluate();

function emailExperimentsAboveThreshold(){
  // TODO: Handle missing email address
  // TODO: Add notification to the UI
  var tableFormatting = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  var htmltable = ['<table ' + tableFormatting +' ">'];
  
  for (row = 0; row < expData.length; row++) {
    htmltable += '<tr>';
    for (col = 0; col < expData[row].length; col++) {
      if (expData[row][col] === "" || 0) { htmltable += '<td>' + 'N/A' + '</td>'; }
      else
        if (row === 0) {
          htmltable += '<th>' + expData[row][col] + '</th>';
        }
        else { htmltable += '<td>' + expData[row][col] + '</td>'; }
    }
    htmltable += '</tr>';
  }
  htmltable += '</table>';

  MailApp.sendEmail({
    to: emailAddress, 
    subject: 'Optimizely Impressions Notification [' + endDate + ']',
    noReply: true,
    name: "Optimizely Impression Notifier", 
    htmlBody: "Hello there, <br>Here's a list of experiments exceeding the set threshold of " + impressionThreshold + " impressions. <br>"
    + "Query Date Range: " + startDate + " to " + endDate + "<br><br>"
    + htmltable
    + "<br>See the " + spreadheetLink.getContent() + ".<br>"
    + "<br>Be good, <br>"
    + "Your Impressions Notifier App"
  });
}

function emailNothingToReport() {
  MailApp.sendEmail({
    to: emailAddress, 
    subject: 'Optimizely Impressions Notification [' + endDate + ']',
    noReply: true,
    name: "Optimizely Impression Notifier", 
    htmlBody: "Hello there, <br>No experiments exceeded the set threshold of " + impressionThreshold + " impressions. <br>"
    + "Query Date Range: " + startDate + " to " + endDate + "<br>"
    + "<br>See the " + spreadheetLink.getContent() + ".<br>"
    + "<br>Be good, <br>"
    + "Your Impressions Notifier App"
  });
}