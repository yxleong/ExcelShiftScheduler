/** Setup your own value here! ESPECIALLY THE TOKEN */
let ACCESS_TOKEN = REPLACE_WITH_YOUR_OWN_TOKEN;

let CLEANING_REMINDER_MSG = "\nA gentle reminder that you have cleaning tasks to tackle today. Please pay attention to the following:\n1. Both of the floor drainers\n2. Sink & drain hole\n3. Toilet bowl\n4. Trash\nA clean space fosters a clear mind. Thank you for your efforts! \n\nAdditionally, please remember to check the box next to your name in the Google Sheet. Here is the link for your convenience: REPLACE_WITH_YOUR_OWN_DOC_LINK\n\nThank you!"

let COMPLETION_MSG = "SUCCESSFUL COMPLETION OF CLEANING\n\nDear all,\n\nWe're delighted to inform you that the cleaning has been successfully completed. You can now enjoy a refreshed and tidy environment! Thank you all for your cooperation.\n\nIf you happen to notice anything that needs further attention or if you find anything less than perfect, please don't hesitate to send us a message. Your feedback is highly appreciated. Here is the link for your convenience: REPLACE_WITH_YOUR_OWN_DOC_LINK\n\nThank you!"

let UPDATE_MSG_20231114 = "EXCITING UPDATES AND NEW FEATURES!\n\nDear all,\n\nWe're excited to share the latest updates with you:\n\n1. Personalized Reminders:\nLine Notify now allows for specific reminders by mentioning someone's name, ensuring more direct and effective communication.\n\n2. Feedback Platform:\nA feedback platform has been established to enable everyone to evaluate and express your thoughts regarding the cleaning results.\n\nThank you!"

let CHECKBOX_COLUMN = 9;
let SATURDAY_COLUMN = 7;
let NAME_COLUMN = 8;
let RED_FONT_RANGE = "L2";

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
  sendToLine("IMPORTANT FEEDBACK REGARDING CLEANING\n\nThe comment from the feedback platform reads:\n\n\"" + feedbackComment + "\"\n\nNot to worry, we see it as a friendly reminder. We appreciate your input, and your diligence in the future is valued. Your efforts contribute to maintaining the high standards we aim for.\n\nIf there are any additional details or specific concerns you'd like to share, please feel free to let us know. Your satisfaction is paramount to us.\n\nThank you for your understanding and commitment to excellence.")
}

function checkCurrentCalendar() {
  var cleaningSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  const todayDate = new Date();
  var todayMonth = monthNames[todayDate.getMonth()];
  var todayYear = todayDate.getFullYear();

  var currentCalendar = todayMonth + " " + todayYear;

  var startRow = 1;
  var lastRow = cleaningSheet.getLastRow();

  for (var row = startRow; row <= lastRow; row++) {
    var taskDate = cleaningSheet.getRange(row, 1).getValue();
    if (taskDate instanceof Date) {
      var taskMonth = monthNames[taskDate.getMonth()];
      var taskYear = taskDate.getFullYear();
      var taskCalendar = taskMonth + " " + taskYear;
      var nextMonthCalendar = monthNames[todayDate.getMonth() + 1] + " " + todayDate.getFullYear();

      if (taskCalendar === currentCalendar) {
        checkRowStart = row;
      }
      if (taskCalendar === nextMonthCalendar) {
        checkRowEnd = row;
        break;
      }
    }
  }

  checkCheckboxStatus(checkRowStart, checkRowEnd);
}

function checkCheckboxStatus(startRow, endRow) {
  var cleaningSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  for (var row = startRow; row <= endRow; row++) {
    var checkboxCell = cleaningSheet.getRange(row, CHECKBOX_COLUMN);
    var message = "default";

    if (checkboxCell.getValue() !== "") {
      var fontColor = cleaningSheet.getRange(row, SATURDAY_COLUMN).getFontColor();
      var redColor = cleaningSheet.getRange(RED_FONT_RANGE).getFontColor();
      var personName = cleaningSheet.getRange(row, NAME_COLUMN).getValue().split(" ")[1];

      if (fontColor === redColor) {
        if (checkboxCell.isChecked()) {
          message = COMPLETION_MSG;
        } else {
          message = " FRIENDLY REMINDER - YOUR CLEANING TASKS AWAIT!\n\nHi " + personName + ",\n"+ CLEANING_REMINDER_MSG;
        }
        
        break;
      } else {
        message = "Not deadline";
      }

    } else {
      message = "Not checkbox";
    }
  }

  var sendResult = sendToLine(message);
  Logger.log("Message for row " + row + ": " + message + " " + sendResult);
}
