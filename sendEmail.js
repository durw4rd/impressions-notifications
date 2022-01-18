var emailAddress = configSheet.getRange("B6").getValue();

var spreadsheetURL = SpreadsheetApp.getActive().getUrl();
var linkTemplate = HtmlService.createTemplate("<a href='<?= spreadsheetURL ?>'>connected Google Sheet</a>");
var spreadheetLink = linkTemplate.evaluate();

function createEmailTable(data) {
  var tableFormatting = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'
  var htmltable = ['<table ' + tableFormatting +' ">'];

  // Populate the table header
  var tableHeader = SpreadsheetApp.getActive().getSheetByName('Impressions').getRange(2,1,1,6).getValues();
  for (col = 0; col < tableHeader[0].length; col++) {
    htmltable += '<th>' + tableHeader[0][col] + '</th>';
  }
  cleanedData = data.map(function(experiment) {
    var originalLink = experiment[1];
    var expId = originalLink.split(',')[1].substring(1, originalLink.split(',')[1].length - 2);
    var htmlLink = "<a href='https://app.optimizely.com/s/redirect-to-emulate?input=" + expId + "&version=2'>" + expId + "</a>"
    experiment[1] = htmlLink;
    return experiment;
  });
  // Populate table body
  for (row = 0; row < cleanedData.length; row++) {
    htmltable += '<tr>';
    for (col = 0; col < cleanedData[row].length; col++) {
      if (cleanedData[row][col] === "" || 0) { htmltable += '<td>' + 'N/A' + '</td>'; }
      else { htmltable += '<td>' + cleanedData[row][col] + '</td>'; }
    }
    htmltable += '</tr>';
  }
  htmltable += '</table>';

  return htmltable;
}

function sendEmail() {
  var expData = loadImpressionData();
  if (expData && expData.length > 0) {
    MailApp.sendEmail({
      to: emailAddress, 
      subject: 'Optimizely Impressions Notification [' + endDate + ']',
      noReply: true,
      name: "Optimizely Impression Notifier", 
      htmlBody: "Hello there, <br><br>Here's a list of experiments exceeding the set threshold of " + impressionThreshold + " impressions. <br>"
      + "Query Date Range: " + startDate + " to " + endDate + "<br><br>"
      + createEmailTable(expData)
      + "<br>See the " + spreadheetLink.getContent() + ".<br>"
      + "<br>Be good, <br>"
      + "Your Impressions Notifier App"
    });
  } else {
    MailApp.sendEmail({
      to: emailAddress, 
      subject: 'Optimizely Impressions Notification [' + endDate + ']',
      noReply: true,
      name: "Optimizely Impression Notifier", 
      htmlBody: "Hello there, <br><br>No experiments exceeded the set threshold of " + impressionThreshold + " impressions. <br>"
      + "Query Date Range: " + startDate + " to " + endDate + "<br>"
      + "<br>See the " + spreadheetLink.getContent() + ".<br>"
      + "<br>Be good, <br>"
      + "Your Impressions Notifier App"
    });
  }
}