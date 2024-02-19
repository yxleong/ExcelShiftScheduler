/** Setup your own value here! */
let START_MONTH = 9;
let START_YEAR = 2023;
let HOW_MANY_MONTH = 12;
let DATA_RANGE = 'L3:L8';

// special case: 
let SPECIAL_DATA_RANGE = 'L3:L7';

// Calendar display format
let startRow = 1; // dayname row
let numRows = 6; // include dayname
let numCols = 7; // days
let count = 0; // loop the data range

function changeEverydayColor() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const todayDate = new Date();
  var todayMonth = getMonthName(todayDate.getMonth() + 1);

  const yesterdayDate = new Date(todayDate);
  yesterdayDate.setDate(todayDate.getDate() - 1);
  var yesterdayMonth = getMonthName(yesterdayDate.getMonth() + 1);
  var yesterdayYear = yesterdayDate.getFullYear();
  var yesterdayCalendar = yesterdayMonth + " " + yesterdayYear;

  var startRow = 1;
  var lastRow = sheet.getLastRow();

  var changeRowStart = -1;
  var changeRowEnd = -1;

  for (var row = startRow; row <= lastRow; row++) {
    var searchDate = sheet.getRange(row, 1).getValue();
    if (searchDate instanceof Date) {
      var searchMonth = getMonthName(searchDate.getMonth()+1);
      var searchYear = searchDate.getFullYear();
      var searchCalendar = searchMonth + " " + searchYear;
      
      if(((yesterdayDate.getMonth()+1+1)) >= 12){
        var nextMonth = getMonthName(1);
        var nextMonthYear = yesterdayDate.getFullYear() + 1;
      } else {
        var nextMonth = getMonthName(yesterdayDate.getMonth()+1+2);
        var nextMonthYear = yesterdayDate.getFullYear();
        
      }
      var nextMonthCalendar = nextMonth + " " + nextMonthYear;

      if (searchCalendar === yesterdayCalendar) {
        changeRowStart = row + 2;
        currentMonth = yesterdayMonth;
        currentYear = yesterdayYear;
      }

      if (searchCalendar === nextMonthCalendar) {
        changeRowEnd = row - 2;
        break;
      }
    }
  }
  setColor(sheet, todayDate, yesterdayDate, changeRowStart, changeRowEnd, numCols, todayMonth, yesterdayMonth);
}

function setColor(sheet, todayDate, yesterdayDate, changeRowStart, changeRowEnd, numCols, todayMonth, yesterdayMonth) {

  var todayRow = changeRowStart;

  for (var row = changeRowStart; row <= changeRowEnd; row++) {

    for (var col = 1; col <= numCols; col++) {

      var cell = sheet.getRange(row, col);
      var cellDate = cell.getValue();

      if (cell.getFontColor() === '#ff0000' && cellDate === yesterdayDate.getDate()){
        todayRow = row;
        cell.setFontColor('black');
        break;
      } 
    }
  }

  var valid = 0;
  var changed = 0;
  for (var row = todayRow; row <= changeRowEnd; row++) {

    for (var col = 1; col <= numCols; col++) {

      var cell = sheet.getRange(row, col);
      var cellDate = cell.getValue();

      if(todayMonth === yesterdayMonth){
        if (cellDate === todayDate.getDate()) {
          // Change font color to red for today's date
          cell.setFontColor('red');
          changed = 1;
        }
      } else {
        var searchDate = sheet.getRange(row, 1).getValue();
        if(searchDate instanceof Date) {
          var searchMonth = getMonthName(searchDate.getMonth()+1);
          if(searchMonth === todayMonth){
            valid = 1;
          }
        }
        if (valid === 1 && cellDate === todayDate.getDate() && cell.getFontColor() != '#aaaaaa') {
          // Change font color to red for today's date
          cell.setFontColor('red'); 
          changed = 1         
        }
      }
    }

    if(changed) break;
  }
}

function generateCalendar() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // var currentDate = new Date();
  // var currentYear = currentDate.getFullYear();
  // var currentMonth = currentDate.getMonth() + 1;
  
  var currentYear = START_YEAR;
  var currentMonth = START_MONTH;

  for (var i = 0; i < HOW_MANY_MONTH; i++) {
    // generate calendar month by month
    getCalendar(sheet, startRow, numRows, numCols, currentMonth + i, currentYear);

    // set next calendar's start row
    startRow += numRows + 2; // blank + mm/yy
  }
}

function getCalendar(sheet, startRow, numRows, numCols, month, year) {
  
  // to handle next year
  if (month > 12) {
    month -= 12;
    year++;
  }

  var startDate = new Date(year, month - 1, 1); // since month count from 0
  var endDate = new Date(year, month, 0); // next month 0th date = this month last date

  // count from 1st day of month
  var currentDate = new Date(startDate);

  // to set dayname start from Sunday
  var currentDay = (currentDate.getDay() + 6) % 7;
  var firstDay = (currentDay + 1) % 7;
  currentDate.setDate(currentDate.getDate() - firstDay);

  // to show mm/yy
  var titleCell = sheet.getRange(startRow, 1, 1, numCols); // startRow, startCol, numRows, numCols
  titleCell.setBackground('#dddddd');
  titleCell.merge();
  titleCell.setValue(`${getMonthName(month)} ${year}`);
  titleCell.setFontWeight('bold');
  titleCell.setHorizontalAlignment('center');

  // to show day
  for (var col = 1; col <= numCols; col++) {
    var dayCell = sheet.getRange(startRow + 1, col);
    dayCell.setValue(getWeekDay(col));
    dayCell.setHorizontalAlignment('center');
    dayCell.setFontColor('blue');
  }

  // to show date
  for (var row = startRow + 2; row <= startRow + numRows; row++) {

    for (var col = 1; col <= numCols; col++) {

      var dateValue = currentDate.getDate();

      var cell = sheet.getRange(row, col);
        cell.setValue(dateValue);
        cell.setFontWeight('bold');
        cell.setHorizontalAlignment('center');

      // to mark date by color displayed
      if (currentDate >= startDate && currentDate <= endDate) {

        var today = new Date();
        if (
          dateValue === today.getDate() &&
          month === today.getMonth() + 1 &&
          year === today.getFullYear()
        ) {
          cell.setFontColor('red');
        } else { // set yesterday date from red to original color
          cell.setFontColor('black');
        }
      } else {
        cell.setFontColor('#aaaaaa');
      }

      // set to next date before loop
      currentDate.setDate(currentDate.getDate() + 1);
    }

    // fill in task after last column of calendar
    setTasks(sheet, row, numCols + 1); // for each week(row)
  }
}

function setTasks(sheet, row, col) {

  var nameCell = sheet.getRange(row, col);

  if(row === 20) {
    var referenceRange = sheet.getRange(SPECIAL_DATA_RANGE);
  } else {
    var referenceRange = sheet.getRange(DATA_RANGE);
  }

  // get font color of Saturday column
  var satColor = sheet.getRange(row, col - 1).getFontColor();
  var black = sheet.getRange('L3').getFontColor();
  var red = sheet.getRange('L2').getFontColor();

  // check if the font color is not grey or blue
  if (satColor === black || satColor === red) {
    
    // offset within the specific data range
    var nameValue = referenceRange.offset(count++, 0).getValue();
    nameCell.setValue(nameValue);

    nameCell.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(referenceRange).build()
    );

    // show the checkbox after column of task for the check in
    sheet.getRange(row, col + 1).insertCheckboxes();

    // to loop again the data given
    if(count >= referenceRange.getNumRows()){
      count = 0;
    }
  }
}

/** Chinese Version

function getMonthName(monthIndex) {
  var monthNames = [
    '一月', '二月', '三月', '四月', '五月', '六月',
    '七月', '八月', '九月', '十月', '十一月', '十二月'
  ];
  return monthNames[monthIndex - 1];
}

function getWeekDay(dayIndex) {
  var weekNames = ['日', '一', '二', '三', '四', '五', '六'];
  return `周${weekNames[dayIndex - 1]}`;
}

*/

function getMonthName(monthIndex) {
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return monthNames[monthIndex - 1];
}

function getWeekDay(dayIndex) {
  var weekNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return weekNames[dayIndex - 1];
}