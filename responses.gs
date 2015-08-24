function labelResponses() {                                   
  //error handling wrap
  var maxRetries = 5;
  for (var currentRetry = 0; currentRetry < maxRetries; currentRetry++) {   
    try {
      deleteAllTriggers();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var mailSenderSheet = ss.getSheetByName('Mail Sender');
      var countersSheet = ss.getSheetByName("Counters (Ignore)");
      var labelName = mailSenderSheet.getRange("B2").getValue();
      var label = GmailApp.getUserLabelByName(labelName);
      var totalThreads = threadCounter(label, "A2");
      var startTime= new Date().getTime();
      var cutoffTime = (1000 * 60 * 4);
  
      //create reply label for campaign
      if(!(GmailApp.getUserLabelByName(labelName + "/Reply"))) {                               
        GmailApp.createLabel(labelName +"/Reply");
      }
      var replyLabel = GmailApp.getUserLabelByName(labelName + "/Reply");
  
      //determine batch index
      if (countersSheet.getRange("A2").getValue()) {
        var batchIndex = countersSheet.getRange("A2").getValue();
      }
      else {
        batchIndex = 1;
        countersSheet.getRange("A2").setValue(batchIndex);
      }
  
      //get threads in groups and cycle through
      var threadStart = (Math.round(batchIndex * 500) - 500); 
      var threads = label.getThreads(threadStart,500);
  
      for (var i = 0; i < threads.length; i++) {
        var currentTime = new Date().getTime();
        //set trigger if cutoff time is reached
        if (currentTime - startTime >= cutoffTime) {
          deleteFirstTrigger();
          ScriptApp.newTrigger("labelResponses")
            .timeBased()
            .at(new Date(currentTime + (1000 * 60)))                                 
            .inTimezone(ss.getSpreadsheetTimeZone())
            .create();
          batchIndex = batchIndex + (i / 500);
          countersSheet.getRange("A2").setValue(batchIndex);
          break;
        } 
        //otherwise, add reply label 
        else {
          var messageCount = threads[i].getMessageCount();
          if (messageCount > 1) {
            threads[i].addLabel(replyLabel);
          }
      
          //check for completion of job; move to printing responses if so
          if ((batchIndex * 500) - 500 + (i + 1) == totalThreads) {                      
            deleteFirstTrigger();
            ScriptApp.newTrigger("printResponses")                               
              .timeBased()
              .at(new Date(currentTime + (1000 * 60)))                                 
              .inTimezone(ss.getSpreadsheetTimeZone())
              .create();
            countersSheet.getRange("A2").setValue(1);
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
    if (currentRetry == (maxRetries - 1)) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      countersSheet = ss.getSheetByName("Counters (Ignore)");
      countersSheet.getRange("A2").setValue(1);
    }
  } 
}

function printResponses() {  
  //error handling wrap
  var maxRetries = 5;
  for (var currentRetry = 0; currentRetry < maxRetries; currentRetry++) {   
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var mailSenderSheet = ss.getSheetByName('Mail Sender');
      var countersSheet = ss.getSheetByName("Counters (Ignore)");
      var labelName = mailSenderSheet.getRange("B2").getValue() + "/Reply";
      var label = GmailApp.getUserLabelByName(labelName);
      var totalThreads = threadCounter(label, "C2");
      var startTime= new Date().getTime();
      var cutoffTime = (1000 * 60 * 4);
      var templateGrabberSheet = ss.getSheetByName("Template Grabber");
      var name = templateGrabberSheet.getRange("E2").getValue();
      var responsesSheet = ss.getSheetByName("Responses");
      var responsesRowDataIndex = getRowDataIndex(responsesSheet);
      var responderEmailData = responsesSheet.getRange(3, 3, responsesRowDataIndex, 2).getValues();
  
      //determine batch index
      if (countersSheet.getRange("A2").getValue()) {
        var batchIndex = countersSheet.getRange("A2").getValue();
      }
      else {
        batchIndex = 1;
        countersSheet.getRange("A2").setValue(batchIndex);
      }
  
      //get threads in groups and cycle through
      var threadStart = (Math.round(batchIndex * 500) - 500); 
      var threads = label.getThreads(threadStart,500);     
  
      for (var i = 0; i < threads.length; i++) {
        var currentTime = new Date().getTime();
        //set trigger if cutoff time is reached
        if (currentTime - startTime >= cutoffTime) {
          deleteFirstTrigger();
          ScriptApp.newTrigger("printResponses")
            .timeBased()
            .at(new Date(currentTime + (1000 * 60)))                                 
            .inTimezone(ss.getSpreadsheetTimeZone())
            .create();
          batchIndex = batchIndex + (i / 500);
          countersSheet.getRange("A2").setValue(batchIndex);
          break;
        } 
        //otherwise, add email address, reply to sheet 
        else {
          for (var k = 0; k < threads[i].getMessageCount(); k++) {
            responsesRowDataIndex = getRowDataIndex(responsesSheet);
            responderEmailData = responsesSheet.getRange(3, 3, responsesRowDataIndex, 2).getValues();
            var added = false;
            var sendername = Session.getActiveUser().getEmail();
            if (name) {
              sendername = name + " <" + Session.getActiveUser().getEmail() + ">";
            }                    
            if (threads[i].getMessages()[k].getFrom() !== sendername) {
              var responderEmailAddress = threads[i].getMessages()[k].getFrom();
              var responderEmailBody = threads[i].getMessages()[k].getPlainBody();

              for (var j = 0; j < responderEmailData.length; j++) {
                if(responderEmailData[j][0] == responderEmailAddress) {
                  added = true;
                  break;
                }
              }
              if(!added) {
                responderEmailData[responderEmailData.length - 1] = [responderEmailAddress, responderEmailBody];
                responsesSheet.getRange(3, 3, responsesRowDataIndex, 2).setValues(responderEmailData);
              }
            }
          }
      
          //check for completion of job; move to next day repetition of process if so
          if ((batchIndex * 500) - 500 + (i + 1) == totalThreads) {
            deleteFirstTrigger();
            ScriptApp.newTrigger("labelResponses")                                                   
              .timeBased()
              .at(new Date(currentTime + (1000 * 60 * 60 * 24)))                                 
              .inTimezone(ss.getSpreadsheetTimeZone())
              .create();
            countersSheet.getRange("A2").setValue(1);
            responsesSheet.getRange("C2").setValue(responderEmailData.length);
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
    if (currentRetry == (maxRetries - 1)) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      countersSheet = ss.getSheetByName("Counters (Ignore)");
      countersSheet.getRange("A2").setValue(1);
    }
  } 
}
