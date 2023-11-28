function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return {
    numRows: sheet.getLastRow(),
    numCols: sheet.getLastColumn(),
    sheet: sheet
  };
}

function findCurrentDateRow(sheet, numRows) {
  const startRow = 6;
  const startCol = 4;
  const currentDate = new Date();
  const currentDateString = Utilities.formatDate(currentDate, 'JST', 'M/d');

  for (let row = startRow; row <= numRows; row++) {
    const dateCell = sheet.getRange(row, startCol);
    const dateString = Utilities.formatDate(dateCell.getValue(), 'JST', 'M/d');

    if (dateString = currentDateString) {
      return row;
    }
  }
  return null;
}

function sendEmailForDate(sheet, currentDateRow, numCols) {
  const startRow = 5;
  const startCol = 5;
  const dateString = Utilities.formatDate(new Date(), 'JST', 'M/d');

  for (let col = startCol; col <= numCols; col++) {
    const newMemberName = sheet.getRange(startRow, col).getValue();
    const currentMemberName = sheet.getRange(currentDateRow, col).getValue();
    const email = getEmailFromName_(currentMemberName);

    if (newMemberName && currentMemberName) {
      const subject = createSubjectLine(newMemberName, dateString);
      const finalHtmlBody = createEmailBody(newMemberName, currentMemberName);

      try {
        // GmailApp.sendEmail(email, subject, "placeholder", {htmlBody: finalHtmlBody})
      } catch(e) {
        console.error("Gmail error:", e)
      }
    }
  }
}

function createSubjectLine(newMemberName, dateString) {
  return `\u23F0リマインド\u23F0【ご連絡】${dateString}分 ${newMemberName}さんの日報返信について(翌営業日まで)`;
}

function createEmailBody(newMemberName, currentMemberName) {
  const htmlBody = HtmlService.createTemplateFromFile('Email');
  htmlBody.newMemberName = newMemberName;
  htmlBody.currentMemberName = currentMemberName;
  return htmlBody.evaluate().getContent();
}

function sendReminderEmail() {
  const { numRows, numCols, sheet} = getSheetData();
  const currentDateRow = findCurrentDateRow(sheet, numRows);

  if (currentDateRow !== null) {
    sendEmailForDate(sheet, currentDateRow, numCols);
  }
}
