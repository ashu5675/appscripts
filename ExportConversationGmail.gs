var ui = SpreadsheetApp.getUi();
function onOpen(e){
  
  ui.createMenu("Gmail Manager").addItem("Get Emails by Label", "getGmailEmails").addToUi();
  
}

function getGmailEmails(){
  var input = ui.prompt('Label Name', 'Enter the label name that is assigned to your emails:', Browser.Buttons.OK_CANCEL);
  
  if (input.getSelectedButton() == ui.Button.CANCEL){
    return;
  }
  
  var label = GmailApp.getUserLabelByName(input.getResponseText());
  var threads = label.getThreads();
  
  for(var i = threads.length - 1; i >=0; i--){
    extractDetails(threads[i]);
    }
}  

  

function extractDetails(thread){
  var messages = thread.getMessages()
  var dateTimeStart = messages[0].getDate();
  var dateTimeEnd = messages[messages.length -1].getDate();
  var count = thread.getMessageCount();
  var subjectText = thread.getFirstMessageSubject();
  var bodyContents = getText(messages);

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  activeSheet.appendRow([dateTimeStart, dateTimeEnd, count, subjectText, bodyContents]);
}

function getText(text){
  body = ""
  for(var i = text.length - 1; i >=0; i--){
    body = body + "\n" + text[i].getPlainBody();
    }
  return body
}
