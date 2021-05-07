// CONFIG
const DATE_COLNUM = 0;
const TIME_COLNUM = 1;
const USECASE_COLNUM = 2;
const STAFF_EMAIL_COLNUM = 3;
const BUSINESS_COLNUM = 4;
const EMAIL_COLNUM = 5;
const CANCPASS_COLNUM = 6;
const STATUS_COLNUM = 7;

const STATUS_ACTIVE = 1;
const STATUS_CANCLE = 0;

// -----------------------------------------------------------------------------------------------------
// Support
const xcelGetDateCol = ()=>{ return xcelGetCol(DATE_COLNUM); }

function doGet() {
  //return HtmlService.createHtmlOutputFromFile('index').setTitle('RPA Clinic Booking');;
  return HtmlService.createTemplateFromFile("index").evaluate().setTitle('RPA Clinic Booking 1.1');
}

function include (fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function getColName(colNum) {
  var temp, letter = '';
  while (colNum > 0) {
    temp = (colNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colNum = (colNum - temp - 1) / 26;
  }
  return letter;
}

function validCell(cellVal) {
  return (typeof(cellVal) == "string" && cellVal.length > 0)
}

function getWorkbook() {
  let app = SpreadsheetApp;
  let appSheet = app.getActiveSpreadsheet().getActiveSheet();
  return appSheet;
}


// -----------------------------------------------------------------------------------------------------
// My Function
function xcelGetCol(whichColNum) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();
  let resultArr = [];
  
  for (let i = 0; i < dataRow.length; i++) {
    let cellVal = dataRow[i][whichColNum];
    if (validCell(cellVal)) {
      resultArr.push(cellVal);
    }
    else break;
  }
  return resultArr;
}

function xcelHasData(whichColNum, searchData) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();

  for (let i = 0; i < dataRow.length; i++) {
    let cellVal = dataRow[i][whichColNum];
    if (validCell(cellVal)) {
      if (cellVal == searchData) return true;
    }
    else break;
  }
  return false;
}

function xcelGetValidTime(dataArr) {
  // Get Arguments
  let pickedDate = dataArr['pickedDate'];
  let bookTimeArr = dataArr['bookTimeArr'];

  // Execute
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();

  for (let i = 0; i < dataRow.length; i++) {
    let dateVal = dataRow[i][DATE_COLNUM];
    if (validCell(dateVal)) {
      let statusVal = dataRow[i][STATUS_COLNUM];
      if (dateVal == pickedDate && statusVal == STATUS_ACTIVE) {
        let timeVal = dataRow[i][TIME_COLNUM];
        for (let j = 0; j < bookTimeArr.length; j++) {
          if (timeVal == bookTimeArr[j]) {
            bookTimeArr.splice(j, 1);
            break;
          }
        }
        if (bookTimeArr.length < 1) { i = dataRow.length; } // If bookTimeArr is empty, stop looping by i.
      }
    }
    else break;
  }
  return bookTimeArr;
}

function xxx() {
  let validStaff = xcelGetValidStaff({
    'bookDate': bookDate,
    'staffArr': staffArr,
    'staffMaxStack': staffMaxStack
  });
  let foundStaff = false
  for (i = 0; i < validStaff.length; i++) {
    if (validStaff[i].email == bookStaffEmail) { foundStaff = true; break; }
  }
  if (foundStaff == false) return "Staff is invalid !";
}

function xcelGetBookedData(pickedDate) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();
  let bookedData = [];

  for (let i = 0; i < dataRow.length; i++) {
    let dateVal = dataRow[i][DATE_COLNUM];
    if (validCell(dateVal)) {
      let statusVal = dataRow[i][STATUS_COLNUM];
      if (dateVal == pickedDate && statusVal == STATUS_ACTIVE) {
        //let currUseGmail = Session.getActiveUser().getEmail();
        bookedData.push([
          dateVal,
          dataRow[i][TIME_COLNUM],
          dataRow[i][USECASE_COLNUM],
          dataRow[i][STAFF_EMAIL_COLNUM],
          dataRow[i][BUSINESS_COLNUM],
          dataRow[i][EMAIL_COLNUM],
          "{{img}}https://sv1.picz.in.th/images/2021/03/26/D6JOUa.png"
          /* ((dataRow[i][CANCPASS_COLNUM] == currUseGmail) ? "{{img}}https://cdn0.iconfinder.com/data/icons/social-messaging-ui-color-shapes-3/3/13-512.png" : "-") */
        ]);
      }
    }
    else break;
  }
  return bookedData;
}

function xcelBookingData(dataArr) {
  // Get Arguments
  let bookDate = dataArr['bookDate'];
  let bookTime = dataArr['bookTime'];
  let bookCase = dataArr['bookCase'];
  let bookStaffEmail = dataArr['bookStaffEmail'];
  let bookTeam = dataArr['bookTeam'];
  let bookEmail = dataArr['bookEmail'];
  let bookCancelPass = dataArr['bookCancelPass'];

  let staffArr = dataArr['staffArr'];
  let staffMaxStack = dataArr['staffMaxStack']; 

  
  let validStaff = xcelGetValidStaff({
    'bookDate': bookDate,
    'staffArr': staffArr,
    'staffMaxStack': staffMaxStack
  });
  if (validStaff.length < 1) return "Have no valid staff !";

  // If email is empty (Random)
  if (bookStaffEmail.length < 1) {
    bookStaffEmail = validStaff[Math.floor(Math.random() * validStaff.length)].email;
  }
  else { // Checking Staff Duplicate
    let foundStaff = false
    for (i = 0; i < validStaff.length; i++) {
      if (validStaff[i].email == bookStaffEmail) { foundStaff = true; break; }
    }
    if (foundStaff == false) return "Staff is invalid !";
  }


  //let sheet = (SpreadsheetApp.getActiveSpreadsheet()).getSheets()[0];
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();
  for (let i = 0; i < dataRow.length; i++) {
    let dateVal = dataRow[i][DATE_COLNUM];
    if (validCell(dateVal)) {
      let timeVal = dataRow[i][TIME_COLNUM];
      let statusVal = dataRow[i][STATUS_COLNUM];
      if (dateVal == bookDate && timeVal == bookTime && statusVal == STATUS_ACTIVE) {
        return "Your booking date has been booked!"; // Found duplicate data.
      }
    }
    else break; 
  }
  //let bookerGmail = Session.getActiveUser().getEmail();
  // 0: Active Status, 1: Cancel Status.
  sheet.appendRow([bookDate, bookTime, bookCase, bookStaffEmail, bookTeam, bookEmail, bookCancelPass, STATUS_ACTIVE]);
  xcelSetFormat();
  return "";
}

function xcelSetFormat(whichColArr = ['A', 'B'], stringFormat = "@") {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();
  let cell, currRow, lastRow = dataRow.length;

  // Set cell format
  for (let i = 0; i < lastRow; i++) {
    currRow = i + 1;
    for (let j = 0; j < whichColArr.length; j++) {
      cell = sheet.getRange(whichColArr[j] + currRow);
      cell.setNumberFormat(stringFormat);
    }
  }

  // Sort Data
  let range = sheet.getRange("A2:K" + (parseInt(lastRow).toString()));
  range.sort(1);
}

function xcelCancelBooking(dataArr) {
  // Get Arguments
  let cancelDate = dataArr['bookDate'];
  let cancelTime = dataArr['bookTime'];
  let cancelPass = dataArr['bookPass'];

  let isSucc = false;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();
  for (let i = 0; i < dataRow.length; i++) {
    let dateVal = dataRow[i][DATE_COLNUM];
    if (validCell(dateVal)) {
      let timeVal = dataRow[i][TIME_COLNUM];
      let statusVal = dataRow[i][STATUS_COLNUM];
      //let bookerGmail = Session.getActiveUser().getEmail();
      if (dateVal == cancelDate && 
          timeVal == cancelTime && 
          statusVal == STATUS_ACTIVE && 
          cancelPass == dataRow[i][CANCPASS_COLNUM]
          ){
        sheet.getRange(getColName(STATUS_COLNUM + 1) + (i + 1)).setValue(STATUS_CANCLE);
        isSucc = true
        break;
      }
    }
    else break;
  }
  return isSucc;
}

function xcelFullDate(bookTimeArr) {
  let tempDateArr = [];
  let fullDateArr = [];
  let bookTimeSize = bookTimeArr.length;

  // Loop MainSheet
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();
  for (let i = 0; i < dataRow.length; i++) {
    let dateVal = dataRow[i][DATE_COLNUM];
    if (validCell(dateVal)) {
      if (dataRow[i][STATUS_COLNUM] == STATUS_ACTIVE) {
        if (dateVal in tempDateArr == false) tempDateArr[dateVal] = 1;
        else tempDateArr[dateVal]++;
        
        if (tempDateArr[dateVal] >= bookTimeSize) fullDateArr.push(dateVal);
      }
    }
    else break;
  }

  // Loop Holiday
  let holidayArr = [];
  sheet = SpreadsheetApp.getActive().getSheetByName("Holiday");
  sheet.getRange("A:A").setNumberFormat("@");  // Change cell format to plain text
  dataRow = sheet.getDataRange().getValues();
  for (let i = 0; i < dataRow.length; i++) {
    let dateVal = dataRow[i][DATE_COLNUM];
    if (validCell(dateVal)) {
      let dateSplit = dateVal.split('-')
      let day = (dateSplit[0] < 10 && dateSplit[0][0] != 0 ? '0' : '') + dateSplit[0]
      let month = (dateSplit[1] < 10 && dateSplit[1][0] == '0' ? dateSplit[1].substring(1) : dateSplit[1])
      let year = dateSplit[2]
      //holidayArr.push(day + "-" + month + "-" + year)
      holidayArr.push(month + "-" + day + "-" + year)
    }
    else break;
  }
  return  {'fullDateArr':fullDateArr, 'holidayArr':holidayArr};
}

function xcelGetValidStaff(dataArr) {
  // Get Arguments
  let bookDate = dataArr['bookDate'];
  let staffArr = dataArr['staffArr'];
  let staffMaxStack = dataArr['staffMaxStack'];
  let staffIdDup = [];

  // Init staffIdDup
  for (let i = 0; i < staffArr.length; i++) { staffIdDup[i] = 0; }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRow = sheet.getDataRange().getValues();
  for (let i = 0; i < dataRow.length; i++) {
    let dateVal = dataRow[i][DATE_COLNUM];
    if (validCell(dateVal) && staffArr.length > 0) {
      if (dataRow[i][DATE_COLNUM] == bookDate & dataRow[i][STATUS_COLNUM] == STATUS_ACTIVE) {
        for (let j = 0; j < staffArr.length; j++) {
          if (dataRow[i][STAFF_EMAIL_COLNUM] == staffArr[j].email) {
            staffIdDup[staffArr[j].id]++;
            if (staffIdDup[staffArr[j].id] >= staffMaxStack) {
              staffArr.splice(j, 1);
              break;
            } 
          }
        }
      }
    }
    else break;
  }
  return staffArr;
}
