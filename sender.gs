function scheduleSending() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mailSenderSheet = ss.getSheetByName('Mail Sender');
  var rowDataIndex = getRowDataIndex(mailSenderSheet);
  var sendingStatusColumn = mailSenderSheet.getRange(2, 4, rowDataIndex, 1);
  var sendingStatusEntries = sendingStatusColumn.getValues();
  var scheduleTime = mailSenderSheet.getRange("A2").getValue();
  var scheduleDate = new Date(scheduleTime);
  var currentTime = new Date().getTime();
  
  deleteAllTriggers();
  
  //check to ensure a schedule time is added
  if (scheduleTime == "") {
    Browser.msgBox("Please add a send time.");
    return;
  }
    
  //check for a future send time                                   
  if (scheduleTime > currentTime) {
    //create a trigger to send at that time
    ScriptApp.newTrigger("sendMail")
      .timeBased()
      .at(scheduleDate)
      .inTimezone(ss.getSpreadsheetTimeZone())
      .create();
  } 
  else {
    Browser.msgBox("Your proposed send time is in the past.");
    return;
  }
  
  //maintain instances of "Delivered"; change all others to "Scheduled"
  for (var i = 0; i < sendingStatusEntries.length; i++) {
    if (sendingStatusEntries[i][0] == "Delivered") {
      sendingStatusEntries[i][0] = ("Delivered");
    }
    else {
      sendingStatusEntries[i][0] = ("Scheduled");
    }
  }
  
  //update sending status column
  sendingStatusColumn.setValues(sendingStatusEntries);
}

//**********************************************************************//

function sendMail() {
  //error handling wrap
  var maxRetries = 5;
  for (var currentRetry = 0; currentRetry < maxRetries; currentRetry++) {   
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var templateGrabberSheet = ss.getSheetByName("Template Grabber");
      var subjectLine = templateGrabberSheet.getRange("A2").getValue();
      var body = templateGrabberSheet.getRange("B2").getValue();
      var mailSenderSheet = ss.getSheetByName('Mail Sender');
      var rowDataIndex = getRowDataIndex(mailSenderSheet);
      var sendingStatusColumn = mailSenderSheet.getRange(2, 4, rowDataIndex, 1);
      var sendingStatusEntries = sendingStatusColumn.getValues();
      var startTime= new Date().getTime();
      var cutoffTime = (1000 * 60 * 4);
      var objects = getWebsiteData();
  
      //cycle through
      for (var i = 0; i < objects.length; i++) {                             //assumes parity between objects.length & rowDataIndex (safe assumption)
        var websiteDataObject = objects[i]; 
        var currentTime = new Date().getTime();
        //set trigger if cutoff time is reached
        if (currentTime - startTime >= cutoffTime) {
          deleteFirstTrigger();
          ScriptApp.newTrigger("sendMail")
            .timeBased()
            .at(new Date(currentTime + (1000 * 60)))                                    
            .inTimezone(ss.getSpreadsheetTimeZone())
            .create();
          break;
        }     
        //otherwise, send customized emails one-by-one & update sending status
        else if (sendingStatusEntries[i][0] == "Scheduled") {       
          var recipient = websiteDataObject.emailAddress;
          var myName = templateGrabberSheet.getRange("E2").getValue();
          var cc = templateGrabberSheet.getRange("F2").getValue();
          var emailBody = customizeMessage(body, websiteDataObject);    
          //base options
          var options = {
            htmlBody: emailBody, 
          };      
          //optional options
          var imageFileName = templateGrabberSheet.getRange("C2").getValue();
          if (imageFileName) {
            var imageBlob = DriveApp.getFilesByName(imageFileName).next().getBlob();
            emailBody = emailBody + "<br><img src='cid:myImage'>";
            options = {
              htmlBody: emailBody,
              inlineImages: {
                myImage: imageBlob
              }
            };   
          }
          var attachmentFileName = templateGrabberSheet.getRange("D2").getValue();
          if (attachmentFileName) {
            var attachmentBlob = DriveApp.getFilesByName(attachmentFileName).next().getBlob();
            options["attachments"] = attachmentBlob;
          }     
          var sendingName = templateGrabberSheet.getRange("E2").getValue();
          if (sendingName) {
            options["name"] = myName;
          }      
          var cc = templateGrabberSheet.getRange("F2").getValue();
          if (cc) {
            options["cc"] = cc;
          }      
          var cellIndex = (i + 2);
          var dataIndex = (i + 1);
          GmailApp.sendEmail(recipient, subjectLine, emailBody, options);
          mailSenderSheet.getRange("D" + cellIndex).setValue("Delivered");                
          //check for completion of job; move to labeling if so
          if ((dataIndex) == rowDataIndex) {
            var labelName = mailSenderSheet.getRange("B2").getValue();
            if(labelName) {
              deleteFirstTrigger();
              ScriptApp.newTrigger("labelMail")
                .timeBased()
                .at(new Date(currentTime + (1000 * 60)))                                   
                .inTimezone(ss.getSpreadsheetTimeZone())
                .create();
              var labelingStatusColumn = mailSenderSheet.getRange(2, 5, rowDataIndex, 1);
              labelingStatusColumn.setValue("Scheduled");
              var responsesSheet = ss.getSheetByName('Responses');
              responsesSheet.getRange("A2").setValue(rowDataIndex);
              return;
            }
            else {
              deleteFirstTrigger();
              Browser.msgBox("Please specify a label, then click 'Resume Labeling'.");
              return;
            }
          }
        }
      }
    }
    //error scenario  
    catch(e) {
      MailApp.sendEmail("YOUR EMAIL HERE", "Error report", e.message);        //SET YOUR EMAIL
      Utilities.sleep(1000);
    }
  }
}

function resumeSending() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mailSenderSheet = ss.getSheetByName('Mail Sender');
  var rowDataIndex = getRowDataIndex(mailSenderSheet);
  var sendingStatusColumn = mailSenderSheet.getRange(2, 4, rowDataIndex, 1);
  var sendingStatusEntries = sendingStatusColumn.getValues();
  
  deleteAllTriggers();
  
  //maintain instances of "Delivered"; change all others to "Scheduled"
  for (var i = 0; i < sendingStatusEntries.length; i++) {
    if (sendingStatusEntries[i][0] == "Delivered") {
      sendingStatusEntries[i][0] = ("Delivered");
    }
    else {
      sendingStatusEntries[i][0] = ("Scheduled");
    }
  }
  
  //update sending status column
  sendingStatusColumn.setValues(sendingStatusEntries);
  //resume sendMail
  sendMail();  
}

//**********************************************************************//

function labelMail() {
  //error handling wrap
  var maxRetries = 5;
  for (var currentRetry = 0; currentRetry < maxRetries; currentRetry++) {   
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var templateGrabberSheet = ss.getSheetByName("Template Grabber");
      var subjectLine = templateGrabberSheet.getRange("A2").getValue();
      var mailSenderSheet = ss.getSheetByName('Mail Sender');
      var rowDataIndex = getRowDataIndex(mailSenderSheet);
      var emailAddressColumn = mailSenderSheet.getRange(2, 3, rowDataIndex, 1);
      var emailAddresses = emailAddressColumn.getValues();
      var labelingStatusColumn = mailSenderSheet.getRange(2, 5, rowDataIndex, 1);
      var labelingStatusEntries = labelingStatusColumn.getValues();
      var startTime= new Date().getTime();
      var cutoffTime = (1000 * 60 * 4);
      var labelName = mailSenderSheet.getRange("B2").getValue();
  
      //create label if designated but not extant
      var label = GmailApp.getUserLabelByName(labelName);
      if (!label) label = GmailApp.createLabel(labelName);
  
      //cycle through
      for (var i = 0; i < rowDataIndex; i++) {
        var currentTime = new Date().getTime();
        //set trigger if cutoff time is reached
        if (currentTime - startTime >= cutoffTime) {
          deleteFirstTrigger();
          ScriptApp.newTrigger("labelMail")
            .timeBased()
            .at(new Date(currentTime + (1000 * 60)))                                 
            .inTimezone(ss.getSpreadsheetTimeZone())
            .create();
          break;
        } 
        //otherwise, label sent emails one-by-one & update labeling status
        else if (labelingStatusEntries[i][0] == "Scheduled") { 
          var threads = GmailApp.search("subject: " + subjectLine + ", to: " + emailAddresses[i][0]);          
          for (var j = 0; j < threads.length; j++) {
            threads[j].addLabel(label);
          }
          var cellIndex = (i + 2);
          var dataIndex = (i + 1);
          mailSenderSheet.getRange("E" + cellIndex).setValue("Labeled");       
      
          //check for completion of job; move to forwarding or response count triggering
          if ((dataIndex) == rowDataIndex) {
            deleteFirstTrigger();
            ScriptApp.newTrigger("labelResponses")                               
              .timeBased()
              .at(new Date(currentTime + (1000 * 60 * 60 * 24)))                                 
              .inTimezone(ss.getSpreadsheetTimeZone())
              .create();       
            return;
          }
        }
      }
    }
    //error scenario  
    catch(e) {
      MailApp.sendEmail("YOUR EMAIL HERE", "Error report", e.message);        //SET YOUR EMAIL
      Utilities.sleep(1000);
    }
  } 
}
    
function resumeLabeling() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mailSenderSheet = ss.getSheetByName('Mail Sender');
  var rowDataIndex = getRowDataIndex(mailSenderSheet);
  var labelingStatusColumn = mailSenderSheet.getRange(2, 5, rowDataIndex, 1);
  var labelingStatusEntries = labelingStatusColumn.getValues();
  var labelName = mailSenderSheet.getRange("B2").getValue();
  
  if(labelName) {
    deleteAllTriggers();
  
    //maintain instances of "Labeled"; change all others to "Scheduled"
    for (var i = 0; i < labelingStatusEntries.length; i++) {
      if (labelingStatusEntries[i][0] == "Labeled") {
        labelingStatusEntries[i][0] = ("Labeled");
      }
      else {
        labelingStatusEntries[i][0] = ("Scheduled");
      }
    }
  
    //update labeling status column
    labelingStatusColumn.setValues(labelingStatusEntries);
    //resume labelMail
    labelMail(); 
  }
  else {
    Browser.msgBox("Please specify a label, then click 'Resume Labeling'.");
  }
}
