/** Setup your own value here! ESPECIALLY THE TOKEN */
let ACCESS_TOKEN = REPLACE_WITH_YOUR_OWN_TOKEN;
let CHECKBOX_COL = 9;
let CHECKBOX_START_ROW = 3;
let SATURDAY_COL = 7;
let RED_RANGE = "L2";

function sendToLine(message) {
  var accessToken = ACCESS_TOKEN;
  
  var url = "https://notify-api.line.me/api/notify";
  var payload = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + accessToken
    },
    "payload": {
      "message": message
    },
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, payload);

  // Logger.log(response.getContentText());
  return response;
}

function checkCheckboxStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var lastRow = sheet.getLastRow();

  for (var row = CHECKBOX_START_ROW; row <= lastRow; row++) {

    var checkboxCell = sheet.getRange(row, CHECKBOX_COL);

    // to avoid empty message
    var message = "default";

    // check if data in cell is a checkbox
    if (checkboxCell.getValue() !== "") {

      var fontColor = sheet.getRange(row, SATURDAY_COL).getFontColor();
      var red = sheet.getRange(RED_RANGE).getFontColor();

      // check if the data in Saturday column is red color
      // if (fontColor === '#FF0000' || fontColor === 'red') {
      if (fontColor === red) { 
        // check if this week's task is done
        if (checkboxCell.isChecked()) {
          message = "Cleaning has been successfully completed. Enjoy the refreshed environment! Thank you for your cooperation.";
        } else { // if not done, send reminder
          message = "Gentle reminder to tackle the cleaning tasks today. A clean space fosters a clear mind. Thank you for your efforts!";
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

}
