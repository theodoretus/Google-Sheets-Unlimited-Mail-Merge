function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);  
  }
}

function deleteFirstTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers[0]) {                                                                      
    ScriptApp.deleteTrigger(triggers[0]);
  }
}

function getRowDataIndex(sheet) {
  var data = sheet.getDataRange().getValues();
  for (var i = data.length-1; i > 0; i--) {           
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j]) return i; 
    }
  }
  return 0;
}

function threadCounter(label, cell) {
  var label = label;
  var counter = 0;
  var threadStart = 0; 
  var threads = label.getThreads(threadStart,100);
  
  while (threads[0]) {
    counter += threads.length;
    threadStart += 100;
    threads = label.getThreads(threadStart,100);
  }
  
  if(cell) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var responsesSheet = ss.getSheetByName('Responses');
    responsesSheet.getRange(cell).setValue(counter);
  }
  return counter;
}

//**********************************************************************//

function getWebsiteData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var customInfoSheet = ss.getSheetByName('Custom Info');
  var rowDataIndex = getRowDataIndex(customInfoSheet);
  var customInfoDataRange = customInfoSheet.getRange(2, 1, rowDataIndex, 3);
  var headersRange = customInfoSheet.getRange(1, 1, 1, 3);
  var headers = headersRange.getValues()[0];
  
  return getObjects(customInfoDataRange.getValues(), normalizeHeaders(headers));
}

function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;                
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

//**********************************************************************//

function customizeMessage (template, data) {
  var email = template;
  
  var templateVars = template.match(/\$\{[^\}]+\}/g);

  for (var i = 0; i < templateVars.length; ++i) {
    var variableData = data[normalizeHeader(templateVars[i])];
    if (templateVars[i] == "${Custom Message}" && variableData) {
      variableData = "<br><br>" + variableData;
    }
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
}

//**********************************************************************//

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize

function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"

function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string

function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.

function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.

function isDigit(char) {
  return char >= '0' && char <= '9';
}
