/* Send Confirmation Email with Google Forms */
 
function Initialize() {
 
  var triggers = ScriptApp.getProjectTriggers();
 
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  ScriptApp.newTrigger("testAndRemoveDuplicateEntires")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}

function testAndRemoveDuplicateEntires(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var sheetData = sheet.getDataRange().getValues();
    
    var indexOfUniqueField = 3; //change this accordingly
    
    //now get the most recently updated row
    var lastRow = sheetData[sheet.getLastRow() - 1];
    var duplicate = false;
    var message = "placeHolderMessage";
    
    for(i in sheetData) {
      if(lastRow[indexOfUniqueField - 1] == sheetData[i][indexOfUniqueField - 1]) {
        duplicate = true;
        var email = sheetData[i][indexOfUniqueField - 1];
        sheet.deleteRow(sheet.getLastRow());
        SendConfirmationMail(message, sheetData[i][indexOfUniqueField - 1]);
        break;
      }
    }
    if(!duplicate) {
      SendConfirmationMail();
    }
  } catch (exception) {
    var e = exception.toString();
    Logger.log(exception.toString());
  }
}
 
function SendConfirmationMail(message, email) {
 
  try {
    
    // This is the body of the auto-reply
    var defaultMessage = "We have received your details.<br>Thanks!<br><br>";
    
    var message = (message === "placeHolderMessage") ? defaultMessage:message;
 
    // This will show up as the sender's name
    var sendername = "Your Name Goes Here";
 
    // Optional but change the following variable
    // to have a custom subject for Google Docs emails
    var subject = "Google Form Successfully Submitted";
 
    var spreadSheetRef = SpreadsheetApp.getActiveSheet();
 
    // This is the submitter's email address
    // a field called "Email Address" is necessary
    var recepientEmailAddress = email;
 
    textbody = message.replace("<br>", "\n");
 
    GmailApp.sendEmail(recepientEmailAddress, subject, textbody, {
      cc: recepientEmailAddress,
      name: sendername,
      htmlBody: message
    });
 
  } catch (exception) {
    var e = exception.toString();
    Logger.log(exception.toString());
  }
 
}