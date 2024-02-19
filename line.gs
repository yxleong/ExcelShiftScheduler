/** Setup your own value here! ESPECIALLY THE TOKEN */
let ACCESS_TOKEN = REPLACE_WITH_YOUR_OWN_TOKEN;

let CLEANING_REMINDER_MSG = "\nA gentle reminder that you have cleaning tasks to tackle today. Please pay attention to the following:\n1. Both of the floor drainers\n2. Sink & drain hole\n3. Toilet bowl\n4. Trash\nA clean space fosters a clear mind. Thank you for your efforts! \n\nAdditionally, please remember to check the box next to your name in the Google Sheet. Here is the link for your convenience: REPLACE_WITH_YOUR_OWN_LINK\n\nThank you!"

let COMPLETION_MSG = "SUCCESSFUL COMPLETION OF CLEANING\n\nDear all,\n\nWe're delighted to inform you that the cleaning has been successfully completed. You can now enjoy a refreshed and tidy environment! Thank you all for your cooperation.\n\nIf you happen to notice anything that needs further attention or if you find anything less than perfect, please don't hesitate to leave a message. Your feedback is highly appreciated. Here is the link for your convenience: REPLACE_WITH_YOUR_OWN_LINK\n\nThank you!"

let UPDATE_MSG_20231114 = "EXCITING UPDATES AND NEW FEATURES!\n\nDear all,\n\nWe're excited to share the latest updates with you:\n\n1. Personalized Reminders:\nLine Notify now allows for specific reminders by mentioning someone's name, ensuring more direct and effective communication.\n\n2. Feedback Platform:\nA feedback platform has been established to enable everyone to evaluate and express the thoughts regarding the cleaning results. Here is the link for your convenience: REPLACE_WITH_YOUR_OWN_LINK\n\nThank you!"

let CHECKBOX_START_ROW = 3;
let CHECKBOX_COL = 9;
let SATURDAY_COL = 7;
let NAME_COL = 8;
let RED_RANGE = "L2";

function sendToLine(message) {
  var accessToken = ACCESS_TOKEN;
  var notifyUrl = "https://notify-api.line.me/api/notify";
  var payload = {
    method: "post",
    headers: {
      Authorization: "Bearer " + accessToken
    },
    payload: {
      message: message
    },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(notifyUrl, payload);

  // Logger.log(response.getContentText());
  return response;
}

function receiveFeedbackMessage() {
  var feedbackSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("feedback");
  var lastFeedbackRow = feedbackSheet.getLastRow();
  var feedbackValues = feedbackSheet.getRange(lastFeedbackRow, 1, 1, feedbackSheet.getLastColumn()).getValues()[0];
  var feedbackComment = feedbackValues[1]; 

  var savingCell = feedbackSheet.getRange(2, 11);
  if(feedbackComment != savingCell.getValue()) {
    message = "IMPORTANT FEEDBACK REGARDING CLEANING\n\nThe comment from the feedback platform reads:\n\n\"" + feedbackComment + "\"\n\nNot to worry, we see it as a friendly reminder. We appreciate your input, and your diligence in the future is valued. Your efforts contribute to maintaining the high standards we aim for.\n\nIf there are any additional details or specific concerns you'd like to share, please feel free to let us know. Your satisfaction is paramount to us.\n\nThank you for your understanding and commitment to excellence.";

    savingCell.setValue(feedbackComment);
    savingCell.setFontColor('white');
    
    var sendResult = sendToLine(message);
    Logger.log("Message: " + message + " " + sendResult);
  }
  else{
    message = "no new msg";
    Logger.log("Message: " + message);
  }
}

function checkCheckboxStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("calendar");

  var lastRow = sheet.getLastRow();

  for (var row = CHECKBOX_START_ROW; row <= lastRow; row++) {

    var checkboxCell = sheet.getRange(row, CHECKBOX_COL);

    // to avoid empty message
    var message = "default";

    // check if data in cell is a checkbox
    if (checkboxCell.getValue() !== "") {
      var personName = sheet.getRange(row, NAME_COL).getValue().split(" ")[1];
      var fontColor = sheet.getRange(row, SATURDAY_COL).getFontColor();
      var red = sheet.getRange(RED_RANGE).getFontColor();

      // check if the data in Saturday column is red color
      // if (fontColor === '#FF0000' || fontColor === 'red') {
      if (fontColor === red) { 
        // check if this week's task is done
        if (checkboxCell.isChecked()) {
          message = COMPLETION_MSG;
        } else {
          message = " FRIENDLY REMINDER - YOUR CLEANING TASKS AWAIT!\n\nHi " + personName + ",\n"+ CLEANING_REMINDER_MSG;
        }

        break; // if found red

      } else { // if Saturday date not red
        message = "Not deadline";
      }

    } else { // if data in cell is not check box
      message = "Not checkbox";
    }
  }

  var sendResult = sendToLine(message);
  Logger.log("Message for row " + row + ": " + message + " " + sendResult);
  //Logger.log("Message for row " + row + ": " + message);

}
