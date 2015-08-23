function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
      .addItem('1. Validate Emails', 'validateEmails')
      .addItem('2. Get Draft', 'getDraft')
      .addItem('3. Start Mail Schedule', 'scheduleSending')
      .addSeparator()
      .addItem('Resume Sending', 'resumeSending')
      .addItem('Resume Labeling', 'resumeLabeling')
      .addSeparator()
      .addSubMenu(ui.createMenu('Other Functions')
          .addItem('Delete All Triggers', 'deleteAllTriggers')
          .addItem('Check for Responses', 'labelResponses'))
      .addToUi();
}

function validateEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var customInfoSheet = ss.getSheetByName('Custom Info');
  var rowDataIndex = getRowDataIndex(customInfoSheet);
  var emailRange = customInfoSheet.getRange(2, 1, rowDataIndex, 1);
  var emails = emailRange.getValues();
  
  //set formula in Cell D2 for email validation 
  var formulaCell = customInfoSheet.getRange("D2");
  formulaCell.setFormula("=ArrayFormula(ISEMAIL(A2:A" + (rowDataIndex + 1) + "))");
  var statusRange = customInfoSheet.getRange(2, 4, rowDataIndex, 1);
  var statuses = statusRange.getValues();
  statusRange.clearFormat();
  
  //cycle through, highlighting validation failures
  for (var i = 0; i < statuses.length; i++) {
    var index = i + 1;
    if (statuses[i][0] == false) {
      var statusCell = statusRange.getCell(index, 1);
      var emailCell = emailRange.getCell(index, 1);
      emailCell.setValue(emailCell.getValue().toLowerCase());
      statuses = statusRange.getValues();
      if (statuses[i][0] == false) {
        statusCell.setBackground('red');
      }
    } 
  }
}

//**********************************************************************//

function getDraft() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateGrabberSheet = ss.getSheetByName("Template Grabber");
  var correctSubjectLine = templateGrabberSheet.getRange("A2").getValue();
  var drafts = GmailApp.getDraftMessages();
  
  //cycle through drafts, looking for correct subject line
  for (var i = 0; i < drafts.length; i++) {
    var subjectLine = drafts[i].getSubject();
    if (subjectLine == correctSubjectLine) {
      
      //put body of draft in sheet (for reference)
      var body = drafts[i].getBody();
      templateGrabberSheet.getRange("B2").setValue(body);
      return;
    }
  }
}
