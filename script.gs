var EMAIL_ADDRESS = 'E-Mail address';
var EMAIL_REPLY_TO_ADDRESS = 'care@iha.help';
var EMAIL_DISPLAY_NAME = 'Inter-­European Aid Association Care Team';

var SUBJECT_DE = 'Feedback zum IHA Einsatz';
var SUBJECT_EN = 'Feedback on the IHA mission';

var BODY_DE = '1q52hJBRIrYVkcxdd3bs8yNEP_v8NpHbXpNLCyOgK6vM';
var BODY_EN = '1n4GmdSDGP_8H0-dSAK11lnoNZvn4MrwoFy0XV_sJsFc';

var SHEET_TEAMS = 'Fragebogen - Teams';
var SHEET_INDIVIDUALS = 'Fragebogen - Individuell';

var EMAIL_ADDRESS_COLUMN_TEAM = 'E-mail address';
var EMAIL_ADDRESS_COLUMN_INDIVIDUALS = 'E-Mail address';
var END_COLUMN = 'End';

var LANGUAGE_COLUMN = 'Preferred communication language';

var LANGUAGE_ENGLISH = 'English';
var LANGUAGE_GERMAN = 'German';

var DAYS_AFTER_BACK = 5;

function sendFeedbackEmails() {
  var tSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEAMS);  
  var iSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INDIVIDUALS);
  
  var tEndColumn = findColumn(tSheet, END_COLUMN);  
  var iEndColumn = findColumn(iSheet, END_COLUMN);
  
  var tEmailColumn = findColumn(tSheet, EMAIL_ADDRESS_COLUMN_TEAM);
  var iEmailColumn = findColumn(iSheet, EMAIL_ADDRESS_COLUMN_INDIVIDUALS);
  
  var tLanguageColumn = findColumn(tSheet, LANGUAGE_COLUMN);
  var iLanguageColumn = findColumn(iSheet, LANGUAGE_COLUMN);
  
  var tItems = findMatchingItems(tSheet, tEndColumn);
  var iItems = findMatchingItems(iSheet, iEndColumn);
  
  for(var i=0; i<tItems.length; i++) {
    var email = tSheet.getRange(tItems[i], tEmailColumn).getValue();
    var language = tSheet.getRange(tItems[i], tLanguageColumn).getValue();
    if(language === LANGUAGE_GERMAN) {
      sendEmail(email, SUBJECT_DE, BODY_DE);
    } else {
      sendEmail(email, SUBJECT_EN, BODY_EN);
    }
  }
  
  for(var i=0; i<iItems.length; i++) {
    var email = iSheet.getRange(iItems[i], iEmailColumn).getValue();
    var language = iSheet.getRange(iItems[i], iLanguageColumn).getValue();
    if(language === LANGUAGE_GERMAN) {
      sendEmail(email, SUBJECT_DE, BODY_DE);
    } else {
      sendEmail(email, SUBJECT_EN, BODY_EN);
    }
  }
}

function sendEmail(email, subject, contentID, attachments) {
  var content = DocumentApp.openById(contentID)
                            .getBody()
                            .editAsText()
                            .getText();
  MailApp.sendEmail(email, subject, content, {
    attachments: attachments,
    name: EMAIL_DISPLAY_NAME,
    replyTo: EMAIL_REPLY_TO_ADDRESS
  });
}

function findMatchingItems(sheet, column) {
  var matchingItems = [];
  
  var range = sheet.getRange(2, column, sheet.getLastRow()-1).getValues();
  var dateOffset = (24*60*60*1000) * DAYS_AFTER_BACK; //'DAYS_AFTER_BACK' days
  var today = new Date();
  today.setTime(today.getTime() - dateOffset);
  
  for(var j=0; j<range.length; j++) {
    if(range[j][0] instanceof Date) {
      var sameDate = range[j][0].getDate() === today.getDate() && range[j][0].getMonth() === today.getMonth() && range[j][0].getFullYear() === today.getFullYear();
      if(sameDate) {
        matchingItems.push(j+2);
      }
    }
  }
  return matchingItems;
}

function findColumn(sheet, type) {
  var range = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  for(i=0; i<range[0].length; i++) {
    if (range[0][i] === type) {
      return i+1;
    }
  }
}

function findRows(sheet, column, value) {
  var range = sheet.getRange(2, column, sheet.getLastRow()-1).getValues();
  var rows = [];
  for(var n=range.length-1; n>=0; n--) {
    if(range[n][0] === value) {
      rows.push({
        line: n+2,
        value: range[n][0]
      });
    }
  }
  return rows;
}

/**
 * Copys a range to another range with this cell at the top left.
 *
 * @param {array} range The range to copy.
 * @return The copied range
 * @customfunction
 */
function COPY(range) {
  return range;
}

/**
 * Copys a range to another range with this cell at the top left.
 *
 * @param {array} base The range to of all data.
 * @param {number} columnIndex The index of the column which should be compare with name.
 * @param {string} range The desired name.
 * @return The copied range
 * @customfunction
 */
function COPY_IF_EQUALS(base, columnIndex, name) {
  var range = [];
  for(var i=0; i<base.length; i++) {
    if(base[i][columnIndex] === name) {
      range[range.length] = base[i];    
    }
  }
  return range;
}
