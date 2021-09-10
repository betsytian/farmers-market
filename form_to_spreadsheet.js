// function onFromSbmit(e) will be executated each time a form response is submitted.
// For each farm, its own spreadsheet will be updated with the list of orders. (The sheets are overwritten
// each time rather than being incrementally updated)

// TODO 
// Change farmNames
// Set correct farmspreadsheet IDs.

var refFriday = 'Fri Apr 10 2020 00:00:01 GMT-0700 (Pacific Daylight Time)'
var milliSecInOneDay = 24 * 60 * 60 * 1000;
var milliSecInOneWeek = 24 * 60 * 60 * 1000 * 7;
var MAX_ORDERS = 200;
var email_address = 'foo@gmail.com'

var responseSheetName = "2/12 Ordering Form";
var farmNames = ["Ken's Top Notch", /*"Medina Berry Farms",  "Bay Fresh Strawberries",*/ "California Specialty Ranch","Swank Farms","Great Valley Poultry"];
var userInfoStartColIndex = 62; // zero-based 
var userInfoEndColIndex   = 64; // zero-based
var farmStartColIndex     = [1, 2, 23, 56]; // zero-based
var farmEndColIndex       = [1, 22, 55, 60]; // zero-based
var farmSheetIDs = ["1aS6AXnk9I_hpcEu4LINsAkRWnrkVSc-q90RAYkAxx0c",
                    /*"1YFeVP74O6FfHTvKLE-UTqKp1dfOMBz6_adjJGbuIUF4", 
                    "1FNe7nq2B76tNoK1Wuuu76N2iDHS37_0kwX8NZzp16IE",*/
                    "1X3Vq-yxsb1CxLQbQEtpzZI0KUkNEZ3fZ0dF274S0gZ0",
                    "1vVtsQUmkatiJEcD7HfSG61gGEEWP6rBsxKLtPVoT0LQ",
                   "1Dw9myC106PdpVH6cJevEJscpSP9o9QayDupI_Mz2f4s"];

function getFormId() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var formURL = sheet.getFormUrl();
  var form = FormApp.openByUrl(formURL);
  var formId = form.getId();
  return formId;
}

// Figure out the date of next Friday, and return it as a string
function getNextFriday() {
  let dateOfRefFriday = Date.parse(refFriday);  // In milliseconds
  let dateOfToday = new Date();                 // Date object 
  let intervalInWeeks = (dateOfToday.getTime() - dateOfRefFriday) / milliSecInOneWeek;
  let numWeeksRoundedUp = Math.ceil(intervalInWeeks);
  let targetDate = new Date(dateOfRefFriday + numWeeksRoundedUp * milliSecInOneWeek);
  return targetDate;
}

function sendDebugMsgEmail(msg) {
  var htmlBody = '<ol>';
  htmlBody += '<li>' + msg + '</li>';
  htmlBody += '</ol>';
  GmailApp.sendEmail(email_address, msg, '', {htmlBody:htmlBody});
}

function stopAcceptingFormResponses(){
  var msg = 'Maximum number of orders has been reached for this week. Thank you and please try it next week.';
  var formId = getFormId();
  var form = FormApp.openById(formId);
  //sendDebugMsgEmail(msg);
  form.setAcceptingResponses(false).setCustomClosedFormMessage(msg);
}

// This function is called whenever someone submits to the form (one row is added to the sheet)
function onFormSubmit(e){
  // Get spreadsheet data in variable 'values'
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(responseSheetName);
  var name = sheet.getName();
  Logger.log(name);
  var range = sheet.getDataRange();
  var values = range.getValues();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  Logger.log('Number of rows : ' + numRows);
  Logger.log('Number of cols : ' + numCols);
  //sendDebugMsgEmail('Number of rows : ' + numRows);

  // Use date of next Friday as the name of the new sheet in Farmer's spreadsheet.
  let nextFriday = getNextFriday();
  let sheetName = nextFriday.toDateString();
  let lastFriday = new Date(nextFriday.getTime() - milliSecInOneWeek);
  
  // "sheets" is an array of sheets to be worked on, each from a spreadsheet of a farm
  // If the name of sheet sheetName doesn't exist, this is the first order for 
  // the week. Insert a new sheet to the left in the spreadsheet with sheetName.  
  // Otherwise use the existing sheet found
  var sheets = new Array(farmNames.length);
  for (let n = 0; n < farmNames.length; n++) {
    let spreadSheet = SpreadsheetApp.openById(farmSheetIDs[n]);
    let newsheet = spreadSheet.getSheetByName(sheetName);
    if (newsheet == null) {
      sheets[n] = spreadSheet.insertSheet(sheetName, 0);
    } else {
      sheets[n] = newsheet;
    }
    sheets[n].clear();
  }

  // In the form response sheet, the groups of columns are lined from left to right
  // for each farmer. When these columns are taken to each farmer's 
  // sheet, the columns need to shift left, except for the first farmer's.
  // Compute the left shift offsets and store them in offsets array
  var offsets = new Array(farmNames.length);
  offsets[0] = 0;
  for (let n = 1; n < farmNames.length; n++) {
    offsets[n] = farmEndColIndex[n-1] - farmStartColIndex[n-1] + 1 + offsets[n-1];
  }  
  
  // Write farmers' sheets.  
  // Source data is from 2D array "values", which starts from [0][0]
  // Destination uses sheet[n].getRange(x,y), which starts from (1, 1)
  
  // Write headers for each farm's spreadsheet
  // Array response is for debug purpose, it's printed out by logger in View->logs
  for (let n = 0; n < farmNames.length; n++) {
    let tmpRange = sheets[n].getRange(1, 1); // Sets timestamp header
    tmpRange.setValue(values[0][0]);
    for (let j = userInfoStartColIndex; j <= userInfoEndColIndex; j++) { // Sets UserInfo header
      let tmpRange = sheets[n].getRange(0+1, j-userInfoStartColIndex+2);  
      tmpRange.setValue(values[0][j]);
    }
    for (let j = farmStartColIndex[n]; j <= farmEndColIndex[n]; j++) { // Sets Produce header
      let tmpRange = sheets[n].getRange(0+1, j-offsets[n]+1+userInfoEndColIndex-userInfoStartColIndex+1); 
      tmpRange.setValue(values[0][j]);
    }
  }
  
  // Go through all rows in the sheet and update the sheet for each farm
  // Ignore rows from earlier than lastFriday as that data must have already
  // been recorded in previous sheets in the farmer's spreadsheets.
  // rowIndices[n] stores which row is to be written to the farmer's sheet 
  // as loop variable i can't be used because of skipped rows. It starts
  // from 3, since the first row is the header and the second row is the product total.
  var rowIndices = new Array(farmNames.length).fill(3);
  for (let i = 0; i < values.length; i++) {
    // If the leftmost cell of a row isn't a date, it's not a valid order, ignore it.
    let cell = values[i][0];
    let dateInMilliSec = Date.parse(cell);
    let rowIdx = i + 1;
    if (isNaN(dateInMilliSec)) {
      continue;
    }
    
    // If the date of the order is earlier than lastFriday, then it has been recorded
    // in previous sheet in the farmer's spreadsheets.  Ignore it.
    if (dateInMilliSec < lastFriday.getTime()) {
      continue;
    }
    
    // For each farm, if any produce of it is ordered, 
    // copy the user info and the ordered pruduces to the farm's sheet
    for (let n = 0; n < farmNames.length; n++) {
      // Set hasOrder to true if there is any cell that isn't empty
      let hasOrder = false;
      for (let j = farmStartColIndex[n]; j <= farmEndColIndex[n]; j++) {
        cell = values[i][j].length;
        if (values[i][j] !== "" && values[i][j] > 0.9) {
          hasOrder = true;  
        }
      }
      if (hasOrder) {
        let tmpRange = sheets[n].getRange(rowIndices[n], 1); // Sets timestamp info
        tmpRange.setValue(values[i][0]);
        for (let j = userInfoStartColIndex; j <= userInfoEndColIndex; j++) {
          let tmpRange = sheets[n].getRange(rowIndices[n], j-userInfoStartColIndex+2);  // Sets user info
          tmpRange.setValue(values[i][j]);
        }
        for (let j = farmStartColIndex[n]; j <= farmEndColIndex[n]; j++) { // Sets produce info
          let tmpRange = sheets[n].getRange(rowIndices[n], j-offsets[n]+1+userInfoEndColIndex-userInfoStartColIndex+1); 
          tmpRange.setValue(values[i][j]);
        }
        
        rowIndices[n]++;
      }
    }
  }
  
  // Calculates the total sum of each product and writes it into row 2
  for (let n = 0; n < farmNames.length; n++) {
    let tmpRange = sheets[n].getRange(2,1);
    tmpRange.setValue("Total");
    for (let j = farmStartColIndex[n]; j <= farmEndColIndex[n]; j++) {
      let tmpRange = sheets[n].getRange(2, j-offsets[n]+1+userInfoEndColIndex-userInfoStartColIndex+1); 
      let temp = rowIndices[n]-3;
      let formula = "=SUM(R[1]C[0]:R[" + temp + "]C[0])";
      tmpRange.setFormulaR1C1(formula);
    }
  }
  
  var numOrderThisWeek = 0;
  // Stop accepting form response if the number of orders within the week is more than a threshold
  if (numOrderThisWeek > MAX_ORDERS) {
    stopAcceptingFormResponses();
  }
}

function onFormClose(e) {
  var sheets = new Array(farmNames.length);
  for (let n = 0; n < farmNames.length; n++) {
    let nextFriday = getNextFriday();
    let sheetName = nextFriday.toDateString();
    let spreadSheet = SpreadsheetApp.openById(farmSheetIDs[n]);
    let newsheet = spreadSheet.getSheetByName(sheetName);
    if (newsheet == null) {
      return;
    } else {
      sheets[n] = newsheet;
    }
  }
 
  
}
