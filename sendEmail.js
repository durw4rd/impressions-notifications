function sendEmailNotification(){
  // TODO: Handle missing email address
  // TODO: Handle multiple email addresses
  // TODO: Add notification to the UI
  var emailAddress = SpreadsheetApp.getActive().getSheetByName('Configuration').getRange("B7").getValue();
  var subject = 'Optimizely Impressions Notification';
  var message = 'Test body';
  MailApp.sendEmail(emailAddress, subject, message);
}