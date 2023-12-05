function sendReminderEmails() {
  const sheetData = getSheetData_(SHEET_CONFIG.assignSheetname);
  const currentDate = Utilities.formatDate(new Date(), 'JST', 'M/d');
  sendEmailsIfDateMatches_(sheetData, currentDate);
}

function getSheetData_(sheetname) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  return {
    numRows: sheet.getLastRow(),
    numCols: sheet.getLastColumn(),
    sheet: sheet
  };
}

function sendEmailsIfDateMatches_(sheetData, currentDate) {
  const startRow = ROW_COL_CONFIG.assignStartRow;
  const dateCol = ROW_COL_CONFIG.assignDateCol;
  for (let row = startRow; row <= sheetData.numRows; row++) {
    const dateCell = sheetData.sheet.getRange(row, dateCol);
    console.log('Date cell:', dateCell.getValue());
    const dateString = Utilities.formatDate(dateCell.getValue(), 'JST', 'M/d');
    console.log("date:", dateString);

    if (dateString === currentDate) {
      sendEmailsForRow_(sheetData, row);
      break; // Stop the loop as the date is found
    }
  }
}

function sendEmailsForRow_(sheetData, row) {
  const startCol = ROW_COL_CONFIG.nameStartCol;
  const newMemberRow = ROW_COL_CONFIG.nameNewMemberRow;
  for (let col = startCol; col <= sheetData.numCols; col++) {
    const newMemberName = sheetData.sheet.getRange(newMemberRow, col).getValue();
    const currentMemberName = sheetData.sheet.getRange(row, col).getValue();
    const [email, name] = getEmailFromName_(currentMemberName);
    if (!email || !name) {
      continue
    }
    console.log("new:", newMemberName, "current:", currentMemberName, "email:", email);

    if (newMemberName && currentMemberName) {
      sendEmail_(newMemberName, email, name);
    }
  }
}

function sendEmail_(newMemberName, email, name) {
  const senderName = SHEET_CONFIG.senderName;
  const sheetName = SHEET_CONFIG.emailBodySheetname;
  const subjectCell = ROW_COL_CONFIG.emailSubjectCell;
  const bodyCell = ROW_COL_CONFIG.emailBodyCell;
  const templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const subjectTemplate = templateSheet.getRange(subjectCell).getValue();
  let bodyTemplate = templateSheet.getRange(bodyCell).getValue();
  bodyTemplate = bodyTemplate.replace(/\n/g, "<br>");

  const subject = subjectTemplate.replace("{{NEW MEMBER}}", newMemberName);
  const htmlBodyTemplate = HtmlService.createTemplate(bodyTemplate.replace("{{NEW MEMBER}}", newMemberName).replace("{{NAME}}", name));

  const finalHtmlBody = htmlBodyTemplate.evaluate().getContent();

  try {
    GmailApp.sendEmail(email, subject, "_", {htmlBody: finalHtmlBody, name: senderName});
    console.log('Email sent!');
  } catch(e) {
    console.error("Gmail error:", e);
  }
}
