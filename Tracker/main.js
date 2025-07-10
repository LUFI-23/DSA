/**
 * Global variable for the Spreadsheet ID.
 */
var SPREADSHEET_ID = "13cvcaKGIe0ooaovF_5e-e93WQ8HEX5k8F0mIhP5lZRE";  // employee database
const REPORT_SHEET_ID = '1PN1x-5O7F2Xxj7yTF3o4IL58smjJ5QhtUGsY-htcF10'; // report of those who haven't filled
const REPORT_SHEET_ID2 = '1okTIkn-7BPT6Gjs5hSL_BNX2xZg5bQh_wIq-RI2mNYI'; // report on basis of hours worked




/**
 * This script checks ONLY column A, from row 15 to row 27.
 * If it finds the date "28/06/2025", it changes it to "27/06/2025".
 * This version correctly compares date objects by formatting them as strings.
 */

// utility function
function updateDate28(url) {
  // --- 1. SETUP ---
  const spreadsheet = SpreadsheetApp.openByUrl(url);
  const sheet = spreadsheet.getSheetByName("Day Shift Work Tracking");
  const spreadsheetTimezone = spreadsheet.getSpreadsheetTimeZone();
  const range = sheet.getRange("A15:A27");
  const values = range.getValues();
  const dateToFind = "07/07/2025"; 
  const replacementDate = "05/07/2025";

  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0]; 
    if (cellValue instanceof Date) {
      const formattedCellValue = Utilities.formatDate(cellValue, spreadsheetTimezone, "dd/MM/yyyy");
      if (formattedCellValue === dateToFind) {
        Logger.log()
        values[i][0] = replacementDate;
      }
    }
  }
}


/**
 * This script checks ONLY column A, from row 15 to row 27.
 * If it finds the date "28/06/2025", it changes it to "27/06/2025".
 * This version correctly compares date objects by formatting them as strings.
 */
function updateDateInSpecificRange_Fixed(url) {
  // --- 1. SETUP ---
  const spreadsheet = SpreadsheetApp.openByUrl(url);
  const sheet = spreadsheet.getSheetByName("Day Shift Work Tracking");
  // const ui = SpreadsheetApp.getUi();
  
  // Get the spreadsheet's timezone for accurate date formatting.
  const spreadsheetTimezone = spreadsheet.getSpreadsheetTimeZone();
  
  // Define the EXACT range to check: Column A, rows 15 through 27.
  const range = sheet.getRange("A15:A27");
  
  // Get the values ONLY from this specific range.
  const values = range.getValues();

  // --- 2. DEFINE DATES ---
  // The date we need to find, as a string in "dd/MM/yyyy" format.
  const dateToFind = "07/07/2025";
  
  // The date we will change it to.
  // Using a string is fine, Google Sheets will interpret it correctly.
  const replacementDate = "05/07/2025";
  
  let changesMade = 0; // To track if we found and changed anything.

  // --- 3. SEARCH AND REPLACE IN THE SPECIFIED RANGE ---
  // Loop through each row of the retrieved data (from row 15 to 27).
  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0]; // We only check the first (and only) column in our range.
    
    // Check if the cell contains a valid Date object.
    if (cellValue instanceof Date) {
      // *** FIX: Format the date from the cell into a "dd/MM/yyyy" string to compare. ***
      const formattedCellValue = Utilities.formatDate(cellValue, spreadsheetTimezone, "dd/MM/yyyy");
      
      // Now, we compare two strings, which works correctly.
      if (formattedCellValue === dateToFind) {
        // If it's a match, update the date in our 'values' array.
        values[i][0] = replacementDate;
        changesMade++;
      }
    }
  }

  // --- 4. APPLY CHANGES ---
  // If we made changes, write the updated 'values' array back to the sheet.
  if (changesMade > 0) {
    range.setValues(values);
    Logger.log("DONE");
    // ui.alert(`Process complete. Changed ${changesMade} cell(s) in range A15:A27.`);
  } else {
    Logger.log("NO");
    // If no matching dates were found in the range, do nothing and inform the user.
    // ui.alert(`No dates matching "${dateToFind}" were found in the specified range (A15:A27).`);
  }
}




function f() {


  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Successfully opened spreadsheet: " + ss.getName() + " for info extraction.");
  } catch (e) {
    Logger.log("Error opening spreadsheet with ID '" + SPREADSHEET_ID + "': " + e.toString());
    // SpreadsheetApp.getUi().alert("Error", "Could not open spreadsheet with ID: " + SPREADSHEET_ID + ". Details: " + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var sheet = ss.getSheets()[0]; // Assuming processing the first sheet
  if (!sheet) {
    Logger.log("Error: No sheets found in the spreadsheet.");
    // SpreadsheetApp.getUi().alert("Error", "No sheets found in the spreadsheet.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  Logger.log("Processing sheet: " + sheet.getName() + " for info extraction.");

  var headerRow = 1; // Assuming headers are in the first row
  var absentColumnName = "Absent";
  var nameColumnName = "Name";
  var emailColumnName = "Email";
  var sheetLinkColumnName = "Sheet Link"; // Or whatever the exact column name is

  // Get column indices
  var absentCol = getColumnIndexByName_(sheet, absentColumnName, headerRow);
  var nameCol = getColumnIndexByName_(sheet, nameColumnName, headerRow);
  var emailCol = getColumnIndexByName_(sheet, emailColumnName, headerRow);
  var sheetLinkCol = getColumnIndexByName_(sheet, sheetLinkColumnName, headerRow);

  // Check if all required columns are found
  if (absentCol === -1 || nameCol === -1 || emailCol === -1 || sheetLinkCol === -1) {
    var missingColumns = [];
    if (absentCol === -1) missingColumns.push(absentColumnName);
    if (nameCol === -1) missingColumns.push(nameColumnName);
    if (emailCol === -1) missingColumns.push(emailColumnName);
    if (sheetLinkCol === -1) missingColumns.push(sheetLinkColumnName);
    Logger.log("Error: One or more required columns not found: " + missingColumns.join(", ") + ". Halting info extraction.");
    // SpreadsheetApp.getUi().alert("Column(s) Not Found", "The following column(s) were not found: " + missingColumns.join(", ") + ". Please check sheet headers.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  Logger.log("Column indices found: Absent=" + absentCol + ", Name=" + nameCol + ", Email=" + emailCol + ", Sheet Link=" + sheetLinkCol);

  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn());
  var data = dataRange.getValues();
  var extractedDataCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    var sheetLink = row[sheetLinkCol - 1];

    updateDateInSpecificRange_Fixed(sheetLink);

    extractedDataCount++;
  }

  Logger.log("Finished extracting information. Processed " + data.length + " data rows. Extracted info for " + extractedDataCount + " non-absent individuals.");
}





/**
 * Helper function to delete all triggers for this project.
 * Useful for cleanup during development.
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) { 
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log(triggers.length > 0 ? 'All triggers deleted.' : 'No triggers to delete.');
  Browser.msgBox(triggers.length > 0 ? 'All triggers deleted.' : 'No triggers to delete.');
}


function isTodaySunday() {
  const today = new Date();
  // The getDay() method returns the day of the week, where Sunday is 0.
  return today.getDay() === 0;
}


function writeLeaveToCellB2(sheetUrl) {


  spreadsheetId = extractSpreadsheetId(sheetUrl);
  const sheetName = "Day Shift Work Tracking"



  // Validate and open the spreadsheet by ID
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    Logger.log("Error: Spreadsheet ID must be provided as a string.");
    // SpreadsheetApp.getUi().alert("Error: Please provide a valid Spreadsheet ID to the function.");
    return false;
  }


  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    Logger.log("Error opening spreadsheet with ID '" + spreadsheetId + "': " + e.toString());
    // SpreadsheetApp.getUi().alert("Error: Could not open spreadsheet with ID '" + spreadsheetId + "'. Ensure the ID is correct and the script has permission.");
    return false;
  }

  // Validate sheetName
  if (!sheetName || typeof sheetName !== 'string') {
    Logger.log("Error: Sheet name must be provided as a string.");
    // SpreadsheetApp.getUi().alert("Error: Please provide a valid sheet name to the function.");
    return false;
  }

  const sheet = ss.getSheetByName(sheetName);

  
  // Get the specific cell B2.
  // You can also use getRange(2, 2) where 2 is the row and 2 is the column number.
  const cell = sheet.getRange("B2");
  
  // Set the value of the cell to "Leave".
  cell.setValue("Leave");
}



function copyAndClearRows(sheetUrl , shiftType ) {

  let hours = [0 , 0 , 0];
  let ss;
  spreadsheetId = extractSpreadsheetId(sheetUrl);
  const sheetName = "Day Shift Work Tracking"

  // Validate and open the spreadsheet by ID
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    Logger.log("Error: Spreadsheet ID must be provided as a string.");
    // SpreadsheetApp.getUi().alert("Error: Please provide a valid Spreadsheet ID to the function.");
    return false;
  }

  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    Logger.log("Error opening spreadsheet with ID '" + spreadsheetId + "': " + e.toString());
    // SpreadsheetApp.getUi().alert("Error: Could not open spreadsheet with ID '" + spreadsheetId + "'. Ensure the ID is correct and the script has permission.");
    return false;
  }

  // Validate sheetName
  if (!sheetName || typeof sheetName !== 'string') {
    Logger.log("Error: Sheet name must be provided as a string.");
    // SpreadsheetApp.getUi().alert("Error: Please provide a valid sheet name to the function.");
    return false;
  }

  const sheet = ss.getSheetByName(sheetName);


  var mainCols = 13; // Columns A to M
  var extraCols = 2; // Columns N and O
  var totalCols = mainCols + extraCols; // A to O = 15 columns
  var rowsToCopy = [];

  // Check rows 2 to 6 for data in column B
  for (var i = 2; i <= 11; i++) {
    var colBValue = sheet.getRange(i, 2).getValue(); // Column B
    if (colBValue !== "") {
      rowsToCopy.push(i);
    }
  }

  if (rowsToCopy.length === 0) return; // No rows to copy, exit function

  for (var i = rowsToCopy.length - 1; i >= 0; i--) { // Reverse to keep order
    var sourceRow = rowsToCopy[i];

    // Get data from columns A to M (1 to 13)
    var mainData = sheet.getRange(sourceRow, 1, 1, mainCols).getValues()[0];

    // Get data from columns N and O (14 and 15)
    var extraData = sheet.getRange(sourceRow, 14, 1, extraCols).getValues()[0];

    // Combine into one array: A to O
    var fullData = [mainData.concat(extraData)];

    const workMode = fullData[0][1];
    const hourSpent = fullData[0][7];

    if((workMode == 'In Office')){
        hours[0] = hours[0] + hourSpent;
    }
    else if( (workMode == 'WFH')){  
        hours[0] = hours[0] + hourSpent;
    }
    else if( (workMode == 'Night')){
        hours[2] = hours[2] + hourSpent;
    }
    else if( (workMode == 'On Duty')){
        hours[0] = hours[0] + hourSpent;
    }
    else if( (workMode == "Half Day")){
        hours[1] = hours[1] + hourSpent;
    }

    if(shiftType == "night"){
      const today = new Date();
      const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
      const formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      fullData[0][0] = formattedYesterday;
    }

    
    // Insert new row at 15 and paste data
    sheet.insertRowBefore(15);
    sheet.getRange(15, 1, 1, totalCols).setValues(fullData);
  }

  // Clear contents in original rows from B to M (2nd to 13th columns)
  rowsToCopy.forEach(row => {
    sheet.getRange(row, 2, 1, mainCols - 1).clearContent(); // Clear B to M
  });

  return hours;
}








/**
 * Creates a time-driven trigger that will execute 'myTargetFunction'
 * once today at 5:00 PM (17:00).
 *
 * To use this:
 * 1. Replace 'myTargetFunction' with the actual name of the function you want to trigger.
 * 2. Run this 'createOneTimeTriggerAt5PM' function once manually from the Apps Script editor
 * or by calling it from another function.
 */
function createOneTimeTrigger(functionToTrigger , triggerHour , triggerMinute) {


  // Check if the target function exists
  try {
    if (typeof this[functionToTrigger] !== 'function') {
      Logger.log(`Error: The function named "${functionToTrigger}" does not exist or is not a function. Please create it or check the name.`);
      return;
    }
  } catch (e) {
    Logger.log(`Error checking function existence: ${e}`);
    return;
  }

  // Delete any existing triggers for the same function to avoid duplicates
  // This is optional but good practice if you might run this setup function multiple times.
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === functionToTrigger) {
      ScriptApp.deleteTrigger(existingTriggers[i]);
      Logger.log(`Deleted existing trigger for "${functionToTrigger}".`);
    }
  }

  // Set the time for the trigger
  const now = new Date();
  const triggerTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), triggerHour, triggerMinute, 0);

  // If 5 PM has already passed for today, schedule it for tomorrow.
  // Or, you could choose to not schedule it, or log an error.
  // For this request ("only once on the same day"), we will proceed and it will fire immediately if past 5 PM.
  // If you want to prevent this, you can add a check:
  if (triggerTime.getTime() < now.getTime()) {
    Logger.log(`It's already past ${triggerHour}:${triggerMinute < 10 ? '0' : ''}${triggerMinute} today. The trigger, if created for today, would fire immediately or not at all depending on Apps Script behavior for past times.`);
    // Option 1: Log and exit
    // Browser.msgBox("Trigger Not Created", `It's already past ${triggerHour}:${triggerMinute} today. Trigger not created.`, Browser.Buttons.OK);
    // return;

    // Option 2: Schedule for tomorrow (uncomment if desired)
    /*
    triggerTime.setDate(triggerTime.getDate() + 1);
    Logger.log(`Scheduled for tomorrow at ${triggerHour}:${triggerMinute < 10 ? '0' : ''}${triggerMinute} instead.`);
    */
   // For the purpose of "only once on the same day", we let it be. If it's past 5 PM,
   // Apps Script might run it very soon or skip it if too far in the past.
   // The most reliable way to ensure it runs "today" is to set it before 5 PM.
  }


  // Create the trigger
  try {
    ScriptApp.newTrigger(functionToTrigger)
      .timeBased()
      .at(triggerTime)
      .create();
    Logger.log(`Trigger created for "${functionToTrigger}" to run today at approximately ${triggerHour}:${triggerMinute < 10 ? '0' : ''}${triggerMinute}.`);
  } catch (e) {
    Logger.log(`Error creating trigger: ${e}`);
  }
}





// -------- CONFIGURATION --------
// This constant can be used for the email signature
const BOT_NAME = "Reminder Assistant"; 

// -------- EMAIL CONTENT TEMPLATES --------

// For sending email to a team member
const MEMBER_EMAIL_SUBJECT_CASE1 = "Reminder: Daily Worklog Updation";
const MEMBER_EMAIL_BODY_CASE1_TEMPLATE = `
Hi ,

This is a courteous reminder to update your daily worklog of [shift_type] shift at your earliest convenience. Your prompt submission helps us stay aligned and organized.
Thank you for your continued cooperation.

You can update the form using the following link:: [sheetLogLink]

Warm regards,
${BOT_NAME}
`;

const MEMBER_EMAIL_SUBJECT_CASE2 = "Confirmation: Daily Worklog Submitted for [shift] shift";
const MEMBER_EMAIL_BODY_CASE2_TEMPLATE = `
Hi [memberName],

Thank you for filling out your daily form.
Your submission for the worklog ([sheetLogLink]) has been recorded.

Warm regards,
${BOT_NAME}
`;





// Updated column mappings for master sheet ( User for Making Report and reporting spoc's ).
const COLUMNS = {   
  TEAM_MEMBERS: 1,   // A - Team Members ( Name of the Member)
  EMAIL: 2,          // B - e-mail address  
  SHEET_LINK: 3,     // C - Sheet Link (Work-Log SpreadSheet link)
  ABSENT: 4,         // F - Absent
  NIGHT_SHIFT: 5,    // G - Night Shift
  TEAM_SPOC: 6,      // H - Team SPOC
  SPOC_EMAIL: 7,     // I - SPOC email
  DAY_WORK : 8,
  NIGHT_WORK : 9,
  HALF_DAY :10
};


function extractSpreadsheetId(url) {
  // const regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  const regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)(?:\/|$|\?)/;

  const match = url.match(regex);
  return match ? match[1] : null;
}



/**
 * Check if two dates are the same day
 */
function isSameDate(date1, date2) {
  // Create new Date objects to avoid timezone issues
  const d1 = new Date(date1.getFullYear(), date1.getMonth(), date1.getDate());
  const d2 = new Date(date2.getFullYear(), date2.getMonth(), date2.getDate());
  
  return d1.getTime() === d2.getTime();
}


/**
 * Function to check if members work log is pushed Down
 * -> True if found
 * -> false otherwise
 */
function check(worksheetLink) {
  try {
    if (!worksheetLink || worksheetLink.toString().trim() === '') {
      Logger.log('No worksheet link provided');
      return false;
    }
    
    // Extract sheet ID from the link
    const sheetId = extractSpreadsheetId(worksheetLink);
    if (!sheetId) {
      Logger.log('Could not extract sheet ID from link');
      return false;
    }
    
    const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    const cellA10Value = sheet.getRange('A10').getValue();
    
    Logger.log(`Checking worksheet, A10 value: ${cellA10Value}, Type: ${typeof cellA10Value}`);
    
    if (cellA10Value instanceof Date) {
      const today = new Date();
      const isSame = isSameDate(cellA10Value, today);
      Logger.log(`Date comparison result: ${isSame}`);
      return isSame;
    }
    
    // If it's a string, try to parse it as a date
    if (typeof cellA10Value === 'string' && cellA10Value.trim() !== '') {
      const parsedDate = new Date(cellA10Value);
      if (!isNaN(parsedDate.getTime())) {
        const today = new Date();
        const isSame = isSameDate(parsedDate, today);
        Logger.log(`Parsed date comparison result: ${isSame}`);
        return isSame;
      }
    }
    
    Logger.log('No valid date found in A10');
    return false;
  } catch (error) {
    Logger.log('Error in check function: ' + error.toString());
    console.error('Error in check function:', error);
    return false;
  }
}


/**
 * TRIGGER 3: Check for defaulters - Runs between 11:00-11:59 PM daily
 * Will Process all the members and will check if there worklog is pushed or not.
 * A member is considered defaulters if their worklog is not pushed.
 */
function check_defaulters_day() {
  try {
    const isSunday = isTodaySunday();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    const lastRow = sheet.getLastRow();
    const defaulters = [];
    
    Logger.log(`Starting defaulter check for ${lastRow - 1} employees`);
    
    for (let row = 2; row <= lastRow; row++) {
      const name = sheet.getRange(row, COLUMNS.TEAM_MEMBERS).getValue();
      
      // Only process rows with team member data
      if (name && name.toString().trim() !== '') {
        const email = sheet.getRange(row, COLUMNS.EMAIL).getValue();
        const worksheetLink = sheet.getRange(row, COLUMNS.SHEET_LINK).getValue();
        const teamSpoc = sheet.getRange(row, COLUMNS.TEAM_SPOC).getValue();
        const spocEmail = sheet.getRange(row, COLUMNS.SPOC_EMAIL).getValue();
        const isAbsent = sheet.getRange(row, COLUMNS.ABSENT).getValue();
        // copyAndClearRows(worksheetLink);

        if(!isAbsent){
          const spreadsheetId = extractSpreadsheetId(worksheetLink)
          const isValidAttendance = checkSheetDataConditions(spreadsheetId , "Day Shift Work Tracking");

          
          if (!isValidAttendance) {
            defaulters.push({
              name: name,
              email: email,
              worksheetLink: worksheetLink,
              teamSpoc: teamSpoc,
              spocEmail: spocEmail
            });
            Logger.log(`Found defaulter: ${name}`);
            if(!isSunday) writeLeaveToCellB2(worksheetLink);
            copyAndClearRows(worksheetLink , 'day');

          }
          else{
              const hours = copyAndClearRows(worksheetLink , 'day');
              hours[0] = hours[0] + sheet.getRange(row , COLUMNS.DAY_WORK).getValue();
              hours[1] = hours[1] + sheet.getRange(row , COLUMNS.NIGHT_WORK).getValue();
              hours[2] = hours[2] + sheet.getRange(row , COLUMNS.HALF_DAY).getValue();
              sheet.getRange(row, COLUMNS.DAY_WORK).setValue(hours[0]);
              sheet.getRange(row, COLUMNS.NIGHT_WORK).setValue(hours[2]);
              sheet.getRange(row, COLUMNS.HALF_DAY).setValue(hours[1]);
          }
        }
        else{
            if(!isSunday) writeLeaveToCellB2(worksheetLink);
            copyAndClearRows(worksheetLink , 'day');
        }

      }
    }
    
    Logger.log(`Found ${defaulters.length} defaulters`);
    
    if (defaulters.length > 0) {
      const defaulterSheetUrl = createDefaulterSheet(defaulters);
      updateReportSheet(defaulterSheetUrl , 'day' , 1);
    } else {
      Logger.log('No defaulters found today');
    }
    
  } catch (error) {
    Logger.log('Error in check_defaulters: ' + error.toString());
    console.error('Error in check_defaulters:', error);
  }
}





function check_defaulters_night() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    const lastRow = sheet.getLastRow();
    const defaulters = [];
    
    Logger.log(`Starting defaulter check for ${lastRow - 1} employees`);
    
    for (let row = 2; row <= lastRow; row++) {
      const name = sheet.getRange(row, COLUMNS.TEAM_MEMBERS).getValue();
      
      // Only process rows with team member data
      if (name && name.toString().trim() !== '') {
        const email = sheet.getRange(row, COLUMNS.EMAIL).getValue();
        const worksheetLink = sheet.getRange(row, COLUMNS.SHEET_LINK).getValue();
        const teamSpoc = sheet.getRange(row, COLUMNS.TEAM_SPOC).getValue();
        const spocEmail = sheet.getRange(row, COLUMNS.SPOC_EMAIL).getValue();
        const nightShift = sheet.getRange(row, COLUMNS.NIGHT_SHIFT).getValue();

        // copyAndClearRows(worksheetLink);

        if(nightShift){
          const spreadsheetId = extractSpreadsheetId(worksheetLink)
          const isValidAttendance = checkSheetDataConditions(spreadsheetId , "Day Shift Work Tracking");
          
          if (!isValidAttendance) {
            defaulters.push({
              name: name,
              email: email,
              worksheetLink: worksheetLink,
              teamSpoc: teamSpoc,
              spocEmail: spocEmail
            });
            Logger.log(`Found defaulter: ${name}`);
          }
          else{
              const hours = copyAndClearRows(worksheetLink , 'night');
              hours[0] = hours[0] + sheet.getRange(row , COLUMNS.DAY_WORK).getValue();
              hours[1] = hours[1] + sheet.getRange(row , COLUMNS.NIGHT_WORK).getValue();
              hours[2] = hours[2] + sheet.getRange(row , COLUMNS.HALF_DAY).getValue();
              sheet.getRange(row, COLUMNS.DAY_WORK).setValue(hours[0]);
              sheet.getRange(row, COLUMNS.HALF_DAY).setValue(hours[1]);
              sheet.getRange(row, COLUMNS.NIGHT_WORK).setValue(hours[2]);
          }
        }

      }
    }
    
    Logger.log(`Found ${defaulters.length} defaulters`);
    
    if (defaulters.length > 0) {
      const defaulterSheetUrl = createDefaulterSheet(defaulters);
      updateReportSheet(defaulterSheetUrl , 'night' , 1);
    } else {
      Logger.log('No defaulters found today');
    }
    
  } catch (error) {
    Logger.log('Error in check_defaulters: ' + error.toString());
    console.error('Error in check_defaulters:', error);
  }
}



function sendEmailToSPOC() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    const lastRow = sheet.getLastRow();
    const defaulters = [];
    
    Logger.log(`Starting defaulter check for ${lastRow - 1} employees`);
    
    for (let row = 2; row <= lastRow; row++) {
      const name = sheet.getRange(row, COLUMNS.TEAM_MEMBERS).getValue();
      
      // Only process rows with team member data
      if (name && name.toString().trim() !== '') {
        const email = sheet.getRange(row, COLUMNS.EMAIL).getValue();
        const worksheetLink = sheet.getRange(row, COLUMNS.SHEET_LINK).getValue();
        const teamSpoc = sheet.getRange(row, COLUMNS.TEAM_SPOC).getValue();
        const spocEmail = sheet.getRange(row, COLUMNS.SPOC_EMAIL).getValue();
        const isAbsent = sheet.getRange(row, COLUMNS.ABSENT).getValue();


        if( !(email && worksheetLink && teamSpoc && spocEmail)){
          continue;
        } 

        if(!isAbsent){  
          // const isValidAttendance = check(worksheetLink);
          spreadsheetId = extractSpreadsheetId(worksheetLink);
          const response = checkSheetDataConditions(spreadsheetId , "Day Shift Work Tracking");
          
          Logger.log("Member : " +name + " ; Response = " + response);

          //checkSheetDataConditions(spreadsheetId, sheetName)
          if (!response) {
            defaulters.push({
              name: name,
              email: email,
              worksheetLink: worksheetLink,
              teamSpoc: teamSpoc,
              spocEmail: spocEmail
            });
            Logger.log(`Found defaulter: ${name}`);
          }
        }

      }
    }
    
    Logger.log(`Found ${defaulters.length} defaulters`);
    
    if (defaulters.length > 0) {
      const defaulterSheetUrl = createDefaulterSheet(defaulters);
      sendReportToManagers(defaulters);
    } else {
      Logger.log('No defaulters found today');
    }
    
  } catch (error) {
    Logger.log('Error in check_defaulters: ' + error.toString());
    console.error('Error in check_defaulters:', error);
  }
}















































/**
 * Create defaulter sheet for the current date
 */
function createDefaulterSheet(defaulters) {
  try {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const sheetName = `Worklog_Feedback_${today}`;
    
    // Create new spreadsheet for defaulters
    const defaulterSpreadsheet = SpreadsheetApp.create(sheetName);


    const sheet = defaulterSpreadsheet.getActiveSheet();


    // Set headers
    sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Email', 'Worksheet Link', 'Team SPOC', 'SPOC Email']]);
    
    // Add defaulter data
    const defaulterData = defaulters.map(d => [d.name, d.email, d.worksheetLink, d.teamSpoc, d.spocEmail]);
    if (defaulterData.length > 0) {
      sheet.getRange(2, 1, defaulterData.length, 5).setValues(defaulterData);
    }
    
    // Format the sheet
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.autoResizeColumns(1, 5);
    
    Logger.log(`Created defaulter sheet: ${sheetName}`);
    return defaulterSpreadsheet.getUrl();
  } catch (error) {
    Logger.log('Error creating defaulter sheet: ' + error.toString());
    console.error('Error creating defaulter sheet:', error);
    return null;
  }
}



/**
 * Update the main report sheet with today's defaulter link
 */
function updateReportSheet(defaulterSheetUrl , shiftType ,type) {
  try {
    let sheetID;
    if(type == 1) sheetID = REPORT_SHEET_ID;
    else sheetID = REPORT_SHEET_ID2;
    const reportSheet = SpreadsheetApp.openById(sheetID).getActiveSheet();
    let today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    if(shiftType != "day"){
      const temp = new Date();
      const yesterday = new Date(temp.getTime() - 24 * 60 * 60 * 1000);
      const formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      today = formattedYesterday;
    }
    
    // Check if headers exist, if not add them
    const lastRow = reportSheet.getLastRow();
    if (lastRow === 0) {
      reportSheet.getRange(1, 1, 1, 3).setValues([['Date', 'Sheet Link' , 'Shift Type']]);
      reportSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    
    // Find the next empty row
    const nextRow = reportSheet.getLastRow() + 1;
    
    // Add date and defaulter sheet link
    reportSheet.getRange(nextRow, 1).setValue(today);
    reportSheet.getRange(nextRow, 2).setValue(defaulterSheetUrl);
    reportSheet.getRange(nextRow, 3).setValue(shiftType);
    
    Logger.log(`Updated report sheet with defaulter link`);
  } catch (error) {
    Logger.log('Error updating report sheet: ' + error.toString());
    console.error('Error updating report sheet:', error);
  }
}






/**
 * Send report to team managers with defaulter information
 */
function sendReportToManagers(defaulters) {
  try {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

    // Group defaulters by SPOC
    const spocGroups = {};
    defaulters.forEach(defaulter => {
      const spocKey = defaulter.teamSpoc;
      const spocEmail = defaulter.spocEmail;

      if (!spocGroups[spocKey]) {
        spocGroups[spocKey] = {
          email: defaulter.spocEmail,
          defaulters: []
        };
      }
      spocGroups[spocKey].defaulters.push(defaulter);
    });
    
    // Send emails to each SPOC
    for (const spocName in spocGroups) {
      const spocData = spocGroups[spocName];
      const subject = `Worklog Feedback : ${today}`;
      
      let emailBody = `Dear ${spocName},\n\n`;
      emailBody += `Please find below the list of team members who have not updated their worklog: \n\n`;
      
      spocData.defaulters.forEach(defaulter => {
        emailBody += `- ${defaulter.name} (${defaulter.email})\n`;
        emailBody += `  Worksheet: ${defaulter.worksheetLink}\n\n`;
      });
      
      emailBody += `Please follow up with the concerned team members.\n\n`;
      emailBody += `Best regards,\n${BOT_NAME}`;
      
      try {
        MailApp.sendEmail(spocData.email, subject, emailBody);
        Logger.log(`Report sent to SPOC: ${spocName} (${spocData.email})`);
      } catch (emailError) {
        Logger.log(`Error sending email to SPOC ${spocName}: ${emailError.toString()}`);
      }
    }
    
  } catch (error) {
    Logger.log('Error in sendReportToManagers: ' + error.toString());
    console.error('Error in sendReportToManagers:', error);
  }
}
















/**
 * Runs Daily at 9
 * update the mastersheet 
 * 1 -> make everyone present
 * 2-> unmark night shift for everyone.
 */



function masterSheetReset() {
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE" || SPREADSHEET_ID === "") {
    Logger.log("ERROR: SPREADSHEET_ID is not set. Please update the global variable SPREADSHEET_ID at the top of the script.");
    // SpreadsheetApp.getUi().alert("Configuration Error", "Please open the script editor (Extensions > Apps Script) and set the SPREADSHEET_ID global variable before running.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Successfully opened spreadsheet: " + ss.getName());
  } catch (e) {
    Logger.log("Error opening spreadsheet with ID '" + SPREADSHEET_ID + "': " + e.toString());
    // SpreadsheetApp.getUi().alert("Error", "Could not open spreadsheet with ID: " + SPREADSHEET_ID + ". Check the ID and permissions. Details: " + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Assuming you are working with the first sheet.
  // If you want to work with a specific sheet by name, use:
  // var sheet = ss.getSheetByName("Your Sheet Name");
  var sheet = ss.getSheets()[0];
  if (!sheet) {
    Logger.log("Error: No sheets found in the spreadsheet.");
    // SpreadsheetApp.getUi().alert("Error", "No sheets found in the spreadsheet.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  Logger.log("Processing sheet: " + sheet.getName());

  unmarkOrMarkAbsences(sheet);
  setNightShiftsToNo(sheet);

  Logger.log("Master reset process completed for sheet: " + sheet.getName());
}


/**
 * Finds the column named "Absent" and clears all data in it, except for the header.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to process.
 */
// function unmarkAllAbsences(sheet) {
//   Logger.log("Starting unmarkAllAbsences function...");
//   var headerRow = 1; // Assuming headers are in the first row
//   var absentColumnName = "Absent";
//   var absentColumnIndex = getColumnIndexByName_(sheet, absentColumnName, headerRow);

//   if (absentColumnIndex === -1) {
//     Logger.log("Error: Column '" + absentColumnName + "' not found in sheet '" + sheet.getName() + "'.");
//     // SpreadsheetApp.getUi().alert("Column Not Found", "Column '" + absentColumnName + "' not found.", SpreadsheetApp.getUi().ButtonSet.OK);
//     return;
//   }
//   Logger.log("Found '" + absentColumnName + "' column at index: " + absentColumnIndex);

//   var lastRow = sheet.getLastRow();
//   if (lastRow > headerRow) { // Only proceed if there's data below the header
//     // Get the range from the row after the header to the last row in the "Absent" column
//     var rangeToClear = sheet.getRange(headerRow + 1, absentColumnIndex, lastRow - headerRow);
//     rangeToClear.clearContent();
//     Logger.log("Cleared content in '" + absentColumnName + "' column from row " + (headerRow + 1) + " to " + lastRow);
//   } else {
//     Logger.log("No data to clear in '" + absentColumnName + "' column below the header.");
//   }
// }


function unmarkOrMarkAbsences(sheet) {
  Logger.log("Starting unmarkOrMarkAbsences function...");
  var headerRow = 1; // Assuming headers are in the first row
  var absentColumnName = "Absent";
  var absentColumnIndex = getColumnIndexByName_(sheet, absentColumnName, headerRow);

  if (absentColumnIndex === -1) {
    Logger.log("Error: Column '" + absentColumnName + "' not found in sheet '" + sheet.getName() + "'.");
    // SpreadsheetApp.getUi().alert("Column Not Found", "Column '" + absentColumnName + "' not found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  Logger.log("Found '" + absentColumnName + "' column at index: " + absentColumnIndex);

  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) {
    Logger.log("No data to process in '" + absentColumnName + "' column below the header.");
    return;
  }

  // Get the current day (0 for Sunday, 1 for Monday, ..., 6 for Saturday)
  var today = new Date();
  var dayOfWeek = today.getDay();
  Logger.log("Current day of the week (0=Sunday): " + dayOfWeek);

  var rangeToModify = sheet.getRange(headerRow + 1, absentColumnIndex, lastRow - headerRow);

  if (dayOfWeek === 0) { // It's Sunday
    Logger.log("It's Sunday. Marking all checkboxes in '" + absentColumnName + "' column.");
    // Set all checkboxes in the range to TRUE (checked)
    var valuesToSet = [];
    for (var i = 0; i < lastRow - headerRow; i++) {
      valuesToSet.push([true]);
    }
    rangeToModify.setValues(valuesToSet);
    Logger.log("Marked checkboxes in '" + absentColumnName + "' column from row " + (headerRow + 1) + " to " + lastRow);
  } else { // It's not Sunday
    Logger.log("It's not Sunday. Unmarking all checkboxes in '" + absentColumnName + "' column.");
    // Clear content (uncheck checkboxes)
    rangeToModify.clearContent(); // Or rangeToModify.setValue(false) if they are just boolean values rather than actual checkboxes
                                 // For actual checkbox data validation cells, clearContent() or setValue(false) works.
    Logger.log("Cleared content in '" + absentColumnName + "' column from row " + (headerRow + 1) + " to " + lastRow);
  }
}

/**
 * Finds the column named "Night Shift" and make all values unmarked ", except for the header.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to process.
 */

function setNightShiftsToNo(sheet) {
  Logger.log("Starting setNightShiftsToNo function...");
  var headerRow = 1; // Assuming headers are in the first row
  var nightShiftColumnName = "Night Shift";
  var nightShiftColumnIndex = getColumnIndexByName_(sheet, nightShiftColumnName, headerRow);

  if (nightShiftColumnIndex === -1) {
    Logger.log("Error: Column '" + nightShiftColumnName + "' not found in sheet '" + sheet.getName() + "'.");
    // SpreadsheetApp.getUi().alert("Column Not Found", "Column '" + nightShiftColumnName + "' not found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  Logger.log("Found '" + nightShiftColumnName + "' column at index: " + nightShiftColumnIndex);

  var lastRow = sheet.getLastRow();
  if (lastRow > headerRow) { // Only proceed if there's data below the header
     // Get the range from the row after the header to the last row in the "Absent" column
    var rangeToClear = sheet.getRange(headerRow + 1, nightShiftColumnIndex, lastRow - headerRow);
    rangeToClear.clearContent();
    Logger.log("Cleared content in '" + nightShiftColumnName + "' column from row " + (headerRow + 1) + " to " + lastRow);
  } else {
    Logger.log("No data to update in '" + nightShiftColumnName + "' column below the header.");
  }
}




function dayShiftReminder() {
  Logger.log("Starting extractNonAbsentEmployeeInfo function...");
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE" || SPREADSHEET_ID === "") {
    Logger.log("ERROR: SPREADSHEET_ID is not set. Please update the global variable SPREADSHEET_ID at the top of the script.");
    // SpreadsheetApp.getUi().alert("Configuration Error", "Please set the SPREADSHEET_ID global variable before running.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Successfully opened spreadsheet: " + ss.getName() + " for info extraction.");
  } catch (e) {
    Logger.log("Error opening spreadsheet with ID '" + SPREADSHEET_ID + "': " + e.toString());
    // SpreadsheetApp.getUi().alert("Error", "Could not open spreadsheet with ID: " + SPREADSHEET_ID + ". Details: " + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var sheet = ss.getSheets()[0]; // Assuming processing the first sheet
  if (!sheet) {
    Logger.log("Error: No sheets found in the spreadsheet.");
    // SpreadsheetApp.getUi().alert("Error", "No sheets found in the spreadsheet.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  Logger.log("Processing sheet: " + sheet.getName() + " for info extraction.");

  var headerRow = 1; // Assuming headers are in the first row
  var absentColumnName = "Absent";
  var nameColumnName = "Name";
  var emailColumnName = "Email";
  var sheetLinkColumnName = "Sheet Link"; // Or whatever the exact column name is

  // Get column indices
  var absentCol = getColumnIndexByName_(sheet, absentColumnName, headerRow);
  var nameCol = getColumnIndexByName_(sheet, nameColumnName, headerRow);
  var emailCol = getColumnIndexByName_(sheet, emailColumnName, headerRow);
  var sheetLinkCol = getColumnIndexByName_(sheet, sheetLinkColumnName, headerRow);

  // Check if all required columns are found
  if (absentCol === -1 || nameCol === -1 || emailCol === -1 || sheetLinkCol === -1) {
    var missingColumns = [];
    if (absentCol === -1) missingColumns.push(absentColumnName);
    if (nameCol === -1) missingColumns.push(nameColumnName);
    if (emailCol === -1) missingColumns.push(emailColumnName);
    if (sheetLinkCol === -1) missingColumns.push(sheetLinkColumnName);
    Logger.log("Error: One or more required columns not found: " + missingColumns.join(", ") + ". Halting info extraction.");
    // SpreadsheetApp.getUi().alert("Column(s) Not Found", "The following column(s) were not found: " + missingColumns.join(", ") + ". Please check sheet headers.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  Logger.log("Column indices found: Absent=" + absentCol + ", Name=" + nameCol + ", Email=" + emailCol + ", Sheet Link=" + sheetLinkCol);

  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn());
  var data = dataRange.getValues();
  var extractedDataCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var isAbsent = row[absentCol - 1]; // -1 because array is 0-indexed, col is 1-indexed


    if (isAbsent) {
      Logger.log("Row " + (i + headerRow + 1) + ": Marked as absent ('" + isAbsent + "'). Skipping.");
      continue;
    }

    // Extract information if not absent
    var name = row[nameCol - 1];
    var email = row[emailCol - 1];
    var sheetLink = row[sheetLinkCol - 1];

    if ( (name) && (email)  && (sheetLink ) ) {
      Logger.log("Send Forward....");
      Logger.log("Name = " + (name) + " : email = " + (email) + " ; sheetlink = " + (sheetLink));
      const spreadsheetId = extractSpreadsheetId(sheetLink);
      const response = checkSheetDataConditions(spreadsheetId , "Day Shift Work Tracking");
      Logger.log("Response = " + response);
      if(response) Logger.log("Sending mail skipped because response is true");
      else{
          sendEmailToTeamMember(name , email , sheetLink , 1 , "day");
          Logger.log("Mail sent to " + (name));
      }
      
    }

    Logger.log("Row " + (i + headerRow + 1) + ": Not Absent. Extracted: Name='" + name + "', Email='" + email + "', Sheet Link='" + sheetLink + "'");
    extractedDataCount++;
  }

  Logger.log("Finished extracting information. Processed " + data.length + " data rows. Extracted info for " + extractedDataCount + " non-absent individuals.");
}






function nightShiftReminder() {
  Logger.log("Starting Reminder function for night shift...");
  if (SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE" || SPREADSHEET_ID === "") {
    Logger.log("ERROR: SPREADSHEET_ID is not set. Please update the global variable SPREADSHEET_ID at the top of the script.");
    // SpreadsheetApp.getUi().alert("Configuration Error", "Please set the SPREADSHEET_ID global variable before running.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("Successfully opened spreadsheet: " + ss.getName() + " for info extraction.");
  } catch (e) {
    Logger.log("Error opening spreadsheet with ID '" + SPREADSHEET_ID + "': " + e.toString());
    // SpreadsheetApp.getUi().alert("Error", "Could not open spreadsheet with ID: " + SPREADSHEET_ID + ". Details: " + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var sheet = ss.getSheets()[0]; // Assuming processing the first sheet
  if (!sheet) {
    Logger.log("Error: No sheets found in the spreadsheet.");
    // SpreadsheetApp.getUi().alert("Error", "No sheets found in the spreadsheet.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  Logger.log("Processing sheet: " + sheet.getName() + " for info extraction.");

  var headerRow = 1; // Assuming headers are in the first row
  var nightShiftColumnName = "Night Shift";
  var nameColumnName = "Name";
  var emailColumnName = "Email";
  var sheetLinkColumnName = "Sheet Link"; // Or whatever the exact column name is

  // Get column indices
  var nightShiftCol = getColumnIndexByName_(sheet, nightShiftColumnName, headerRow);
  var nameCol = getColumnIndexByName_(sheet, nameColumnName, headerRow);
  var emailCol = getColumnIndexByName_(sheet, emailColumnName, headerRow);
  var sheetLinkCol = getColumnIndexByName_(sheet, sheetLinkColumnName, headerRow);

  // Check if all required columns are found
  if (nightShiftCol === -1 || nameCol === -1 || emailCol === -1 || sheetLinkCol === -1) {
    var missingColumns = [];
    if (nightShiftCol === -1) missingColumns.push(nightShiftColumnName);
    if (nameCol === -1) missingColumns.push(nameColumnName);
    if (emailCol === -1) missingColumns.push(emailColumnName);
    if (sheetLinkCol === -1) missingColumns.push(sheetLinkColumnName);
    Logger.log("Error: One or more required columns not found: " + missingColumns.join(", ") + ". Halting info extraction.");
    // SpreadsheetApp.getUi().alert("Column(s) Not Found", "The following column(s) were not found: " + missingColumns.join(", ") + ". Please check sheet headers.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  Logger.log("Column indices found: Night Shift=" + nightShiftCol + ", Name=" + nameCol + ", Email=" + emailCol + ", Sheet Link=" + sheetLinkCol);

  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn());
  var data = dataRange.getValues();
  var extractedDataCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var isNightShift = row[nightShiftCol - 1]; // -1 because array is 0-indexed, col is 1-indexed


    if (!isNightShift) {
      Logger.log("Row " + (i + headerRow + 1) + ": Unmarked for Night Shift. Skipping.");
      continue;
    }

    // Extract information 
    var name = row[nameCol - 1];
    var email = row[emailCol - 1];
    var sheetLink = row[sheetLinkCol - 1];

    if ( (name) && (email)  && (sheetLink ) ) {
      Logger.log("Send Forward....");
      Logger.log("Name = " + (name) + " : email = " + (email) + " ; sheetlink = " + (sheetLink));
      const spreadsheetId = extractSpreadsheetId(sheetLink);
      const response = checkSheetDataConditions(spreadsheetId , "Day Shift Work Tracking");
      Logger.log("Response = " + response);
      if(response) Logger.log("Sending mail skipped because response is true");
      else sendEmailToTeamMember(name , email , sheetLink , 1 , "night");
      Logger.log("Mail sent to " + (name));
    }

    Logger.log("Row " + (i + headerRow + 1) + ": Marked for Night Shift. Extracted: Name='" + name + "', Email='" + email + "', Sheet Link='" + sheetLink + "'");
    
    extractedDataCount++;
  }

  Logger.log("Finished extracting information. Processed " + data.length + " data rows. Extracted info for " + extractedDataCount + " individuals.");
}





function checkSheetDataConditions(spreadsheetId, sheetName) {
  let ss;

  // Validate and open the spreadsheet by ID
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    Logger.log("Error: Spreadsheet ID must be provided as a string.");
    // SpreadsheetApp.getUi().alert("Error: Please provide a valid Spreadsheet ID to the function.");
    return false;
  }

  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    Logger.log("Error opening spreadsheet with ID '" + spreadsheetId + "': " + e.toString());
    // SpreadsheetApp.getUi().alert("Error: Could not open spreadsheet with ID '" + spreadsheetId + "'. Ensure the ID is correct and the script has permission.");
    return false;
  }

  // Validate sheetName
  if (!sheetName || typeof sheetName !== 'string') {
    Logger.log("Error: Sheet name must be provided as a string.");
    // SpreadsheetApp.getUi().alert("Error: Please provide a valid sheet name to the function.");
    return false;
  }

  const sheet = ss.getSheetByName(sheetName);

  // Define the rows to check (2, 3, 4, 5, 6)
  const rowsToInspect = [2, 3, 4, 5, 6 , 7 , 8 , 9 , 10 , 11];

  // Define the column indices that must be filled.
  // Column B is 2, C is 3, D is 4, E is 5, F is 6, H is 8, I is 9.
  const requiredColumnIndices = [2, 3, 4, 5, 6, 8, 9];

  // Loop through each row number specified in rowsToInspect
  for (let i = 0; i < rowsToInspect.length; i++) {
    const currentRowNumber = rowsToInspect[i];
    let allRequiredCellsInRowAreFilled = true; // Assumption for the current row

    // Check each required column for the current row
    for (let j = 0; j < requiredColumnIndices.length; j++) {
      const currentColumnIndex = requiredColumnIndices[j];
      // Get the cell's value
      const cellValue = sheet.getRange(currentRowNumber, currentColumnIndex).getValue();

      // Check if the cell is empty.
      // A cell is considered empty if it's null, an empty string, or contains only whitespace.
      if (cellValue === null || String(cellValue).trim() === "") {
        allRequiredCellsInRowAreFilled = false; // Mark this row as not meeting the criteria
        break; // Exit the inner loop (column check) for this row, as one required cell is empty
      }
    }

    // If all required cells were filled for this particular row
    if (allRequiredCellsInRowAreFilled) {
      Logger.log("Condition met in row: " + currentRowNumber + " on sheet: " + sheetName + " (Spreadsheet ID: " + spreadsheetId + ")");
      return true; // At least one row satisfies the condition, so return true immediately
    }
  }

  // If the script has looped through all specified rows and hasn't returned true,
  // it means no row met the conditions.
  Logger.log("No row met the specified conditions on sheet: " + sheetName + " (Spreadsheet ID: " + spreadsheetId + ")");
  return false;
}





/**
 * Sends an email to a team member based on their worklog submission status.
 *
 * @param {string} memberName The name of the team member.
 * @param {string} memberEmail The email address of the team member.
 * @param {string} sheetLogLink The link to the Google Form or sheet for worklog submission.
 * @param {number} caseNumber Determines the email content:
 * 1: Member has not filled the log yet (reminder).
 * 2: Member has filled the log (confirmation).
 * @param {string} [companyOrTeamNameOptional=COMPANY_TEAM_NAME] Optional: The name of the company or team for the email signature. Defaults to COMPANY_TEAM_NAME.
 */
function sendEmailToTeamMember(memberName, memberEmail, sheetLogLink, caseNumber , shift) {
  if (!memberName || !memberEmail || !sheetLogLink || (caseNumber !== 1 && caseNumber !== 2)) {
    Logger.log("Error: Missing or invalid arguments for sendEmailToTeamMember. Required: memberName, memberEmail, sheetLogLink, caseNumber (1 or 2).");
    Logger.log(`Provided: Name='${memberName}', Email='${memberEmail}', Link='${sheetLogLink}', Case='${caseNumber}'`);
    return;
  }

  let subject = "";
  let body = "";

  try {
    if (caseNumber === 1) {
      subject = MEMBER_EMAIL_SUBJECT_CASE1
        .replace("[shift]" , shift);
      body = MEMBER_EMAIL_BODY_CASE1_TEMPLATE
        .replace("[memberName]", memberName.trim())
        .replace("[sheetLogLink]", sheetLogLink)
        .replace("[shift_type]" , shift)
       ; // Ensure global constant replacement if template uses it
    } else if (caseNumber === 2) {
      subject = MEMBER_EMAIL_SUBJECT_CASE2
        .replace("[shift]" , shift);
      body = MEMBER_EMAIL_BODY_CASE2_TEMPLATE
        .replace("[memberName]", memberName.trim())
        .replace("[sheetLogLink]", sheetLogLink)
       
        ; // Ensure global constant replacement
    }

    MailApp.sendEmail(memberEmail, subject, body);
    Logger.log(`Email sent to ${memberName} (${memberEmail}) for case ${caseNumber}. Subject: ${subject}`);

  } catch (e) {
    Logger.log(`Error in sendEmailToTeamMember for ${memberEmail}: ${e.toString()} \nStack: ${e.stack}`);
    // Optionally, notify an admin of this error
    // MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "Script Error: sendEmailToTeamMember", `Error: ${e.toString()} \nFor: ${memberEmail}\nStack: ${e.stack}`);
  }
}




/**
 * Helper function to get the column index (1-based) by its name from the header row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {string} columnName The name of the column to find.
 * @param {number} headerRow The row number where headers are located (e.g., 1 for the first row).
 * @return {number} The 1-based column index, or -1 if not found.
 * @private
 */
function getColumnIndexByName_(sheet, columnName, headerRow) {
  if (!sheet || !columnName) {
    Logger.log("getColumnIndexByName_ called with invalid parameters.");
    return -1;
  }
  headerRow = headerRow || 1; // Default to row 1 if not specified
  var lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) { // Sheet is empty or has no columns with data
      Logger.log("Sheet " + sheet.getName() + " has no columns with data.");
      return -1;
  }
  var headersRange = sheet.getRange(headerRow, 1, 1, lastColumn);
  var headers = headersRange.getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === columnName.trim()) {
      return i + 1; // Column index is 1-based
    }
  }
  Logger.log("Column '" + columnName + "' not found in header row " + headerRow + " of sheet '" + sheet.getName() + "'. Headers found: " + headers.join(", "));
  return -1; // Not found
}


function defaulter_check(){
  try {
    const isSunday = isTodaySunday();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
    const lastRow = sheet.getLastRow();
    const defaulters = [];
    
    Logger.log(`Starting defaulter check for ${lastRow - 1} employees`);
    
    for (let row = 2; row <= lastRow; row++) {
      const name = sheet.getRange(row, COLUMNS.TEAM_MEMBERS).getValue();
      
      // Only process rows with team member data
      if (name && name.toString().trim() !== '') {
        const email = sheet.getRange(row, COLUMNS.EMAIL).getValue();
        const worksheetLink = sheet.getRange(row, COLUMNS.SHEET_LINK).getValue();
        const teamSpoc = sheet.getRange(row, COLUMNS.TEAM_SPOC).getValue();
        const spocEmail = sheet.getRange(row, COLUMNS.SPOC_EMAIL).getValue();
        const dayWork = sheet.getRange(row , COLUMNS.DAY_WORK).getValue();
        const halfDay = sheet.getRange(row , COLUMNS.HALF_DAY).getValue();
        const nightWork = sheet.getRange(row , COLUMNS.NIGHT_WORK).getValue();
        sheet.getRange(row , COLUMNS.DAY_WORK).setValue(0);
        sheet.getRange(row , COLUMNS.HALF_DAY).setValue(0);
        sheet.getRange(row , COLUMNS.NIGHT_WORK).setValue(0);
        let str = '';

        if(dayWork > 7.5){
            if(str.length == 0) str = 'Day-Shift';
            else str = str + " , Day-Shift";
        }

        if(halfDay > 4.5){
            if(str.length == 0) str = 'Half-Day';
            else str = str + " , Half-Day";
        }

        if(nightWork > 7.5){
            if(str.length == 0) str = 'Night';
            else str = str + " , Night";
        }
        

        if(str.length > 0){
            defaulters.push({
              name: name,
              email: email,
              worksheetLink: worksheetLink,
              teamSpoc: teamSpoc,
              spocEmail: spocEmail
            });
        }

      }
    }
    
    Logger.log(`Found ${defaulters.length} defaulters`);
    
    
    if (defaulters.length > 0) {
      const defaulterSheetUrl = createDefaulterSheet(defaulters);
      updateReportSheet(defaulterSheetUrl , '--' , 2);
    } else {
      Logger.log('No defaulters found today');
    }
    
  } catch (error) {
    Logger.log('Error in check_defaulters: ' + error.toString());
    console.error('Error in check_defaulters:', error);
  }
}


function checkSpecificDates() {
  // Get the current date and time
  const today = new Date();

  // Get the day of the month (a number from 1 to 31)
  const day = today.getDate();

  // Get the month (a number from 0 to 11, where 0 is January, 1 is February, etc.)
  // This is a common source of errors, so we must be careful.
  // May is the 5th month, so its index is 4.
  // August is the 8th month, so its index is 7.
  const month = today.getMonth();

  
  const isJan01 = (month === 1 && day === 1);
  const isJan26 = (month === 1 && day === 26);
  const isMar14 = (month === 3 && day === 14);
  const isMar15 = (month === 3 && day === 15);
  const isApr10 = (month === 4 && day === 10);
  const isAug09 = (month === 8 && day === 09);
  const isAug15 = (month === 8 && day === 15);
  const isAug16 = (month === 8 && day === 16);
  const isOct01 = (month === 10 && day === 01);
  const isOct02 = (month === 10 && day === 02);
  const isOct21 = (month === 10 && day === 21);
  const isOct22 = (month === 10 && day === 22);
  const isOct23 = (month === 10 && day === 23);


  // Check if either of the conditions are true
  if (isJan01 || isJan26 || isMar14 || isMar15 || isApr10 || isAug09 || isAug15 || isAug16 || isOct01 || isOct02 || isOct21 || isOct22 || isOct23) {
    // If today is one of the target dates, log a message and return true.
    // You can replace Logger.log with any action you want to perform,
    // like sending an email or updating a spreadsheet.
    Logger.log("Today is a Holiday");
    return true;
  } else {
    // If it's not one of the target dates, log a different message and return false.
    Logger.log("Today is not a Holiday.");
    return false;
  }
}



// will be called everyday bw 00-01
function MasterReset() {
  createOneTimeTrigger('check_defaulters_night' , 02 , 00); // for night report and night worklog push
  createOneTimeTrigger('defaulter_check' , 03 , 00);

  const flag = checkSpecificDates();
  if(flag) return;

  createOneTimeTrigger('masterSheetReset' , 8 , 00); // for daily default updation of employe dataset
  createOneTimeTrigger('dayShiftReminder' , 17 , 20); // for day shift reminder
  createOneTimeTrigger('sendEmailToSPOC' , 19 , 00); // for day shift email to spoc's
  createOneTimeTrigger('check_defaulters_day' , 22 , 30); // for day shift report and day push
  createOneTimeTrigger('nightShiftReminder' , 23 , 30); // for night shift reminder
  
}




