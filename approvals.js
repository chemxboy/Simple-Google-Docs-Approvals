function onOpen() {
  var mySheet = SpreadsheetApp.getActiveSpreadsheet();  
  var menuEntries = [ 
	{name: "Submit for Approval", functionName: "submitForApproval"},
	{name: "Approve", functionName: "markApproved"},
	{name: "Reject", functionName: "markRejected"},
	];
  mySheet.addMenu("Approvals", menuEntries);
}

function submitForApproval(theType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
  var r = mySheet.getActiveRange();
  var currentCol = r.getColumn();
  var currentRow = r.getRow();
  var firstName = mySheet.getRange(currentRow, 3).getValue();
  var lastName = mySheet.getRange(currentRow, 4).getValue();
  var statusType = mySheet.getRange(currentRow, 1).getValue();
  var fullName = firstName+" "+lastName;
  ss.toast('Sending for Approval...', 'Please Wait');
  var emailSubject    = "Approval Required: "+statusType+" - "+fullName;
  var emailBody = "Employee Change: "+statusType+"<br />Employee: "+fullName+"<br />Link: http://www.your.spreadsheet.com/";
  var advancedArgs = {htmlBody:emailBody};
  MailApp.sendEmail("the.cfo@yourcompany.com", emailSubject, emailBody , advancedArgs);
  mySheet.getRange(currentRow, 6).setValue("Waiting Approval");
  ss.toast('Approval Request Sent', "Complete", -6);
  SpreadsheetApp.flush();
}
   
function markApproved(theType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();    
  var app = UiApp.createApplication();
  var vPanel = app.createVerticalPanel();
  app.setWidth( 400 );
  app.setHeight( 200 );
  var addressesLabel = app.createLabel("Approval Notes:");
  vPanel.add( addressesLabel );
  var addressesTextBox = app.createTextArea().setId('addresses').setName('addresses');
  addressesTextBox.setSize( '400px', '150px' );
  vPanel.add( addressesTextBox );
  var submitButton = app.createButton("Approve");
  vPanel.add( submitButton );
  var submitHandler = app.createServerClickHandler('finishApproval');
  submitHandler.addCallbackElement(vPanel);
  submitButton.addClickHandler(submitHandler);
  app.add( vPanel );
  ss.show( app );  
}

function finishApproval(e) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
  var r = mySheet.getActiveRange();
  var currentCol = r.getColumn();
  var currentRow = r.getRow();
  var firstName = mySheet.getRange(currentRow, 3).getValue();
  var lastName = mySheet.getRange(currentRow, 4).getValue();
  var statusType = mySheet.getRange(currentRow, 1).getValue();
  var fullName = firstName+" "+lastName;
  var addresses = e.parameter.addresses;
  mySheet.getRange(currentRow, 7).setValue(addresses);
  ss.toast('Sending Approved Message...', 'Please Wait');
  var emailSubject    = "Request Approved: "+statusType+" - "+fullName;
  var emailBody = "Employee Change: "+statusType+"<br />Employee: "+fullName+"<br />http://www.your.spreadsheet.com/";
  var advancedArgs = {htmlBody:emailBody};
  MailApp.sendEmail("your.hr.employee@yourcompany.com", emailSubject, emailBody , advancedArgs);
  mySheet.getRange(currentRow, 6).setValue("Approved");
  ss.toast('Approval Sent', "Complete", -1);
  SpreadsheetApp.flush();
  return closeUiApp();
}
    
    
function markRejected(theType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();   
  var app = UiApp.createApplication();
  var vPanel = app.createVerticalPanel();
  app.setWidth( 400 );
  app.setHeight( 200 );
  var addressesLabel = app.createLabel("Rejection Notes:");
  vPanel.add( addressesLabel );
  var addressesTextBox = app.createTextArea().setId('addresses').setName('addresses');
  addressesTextBox.setSize( '400px', '150px' );
  vPanel.add( addressesTextBox );
  var submitButton = app.createButton("Reject");
  vPanel.add( submitButton );
  var submitHandler = app.createServerClickHandler('finishRejection');
  submitHandler.addCallbackElement(vPanel);
  submitButton.addClickHandler(submitHandler);
  app.add( vPanel );
  ss.show( app );   
}

function finishRejection(e) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
  var r = mySheet.getActiveRange();
  var currentCol = r.getColumn();
  var currentRow = r.getRow();
  var firstName = mySheet.getRange(currentRow, 3).getValue();
  var lastName = mySheet.getRange(currentRow, 4).getValue();
  var statusType = mySheet.getRange(currentRow, 1).getValue();
  var fullName = firstName+" "+lastName;
  var addresses = e.parameter.addresses;
  mySheet.getRange(currentRow, 7).setValue(addresses);
  ss.toast('Sending Rejection Message...', 'Please Wait');
  var emailSubject    = "Request Rejected: "+statusType+" - "+fullName;
  var emailBody = "Employee Change: "+statusType+"<br />Employee: "+fullName+"<br />Link: http://www.your.spreadsheet.com/";
  var advancedArgs = {htmlBody:emailBody};
  MailApp.sendEmail("your.hr.employee@yourcompany.com", emailSubject, emailBody , advancedArgs);
  mySheet.getRange(currentRow, 6).setValue("Rejected");
  ss.toast('Rejection Sent', "Complete", -1);
  SpreadsheetApp.flush();
  return closeUiApp();
}
    
function closeUiApp() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}
