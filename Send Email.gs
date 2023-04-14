function sendReminderEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1")
  const numRows = sheet.getLastRow();
  const numCols = sheet.getLastColumn();
  const currentDate = new Date();
  const currentDateString = Utilities.formatDate(currentDate, 'JST', 'M/d');
  let dateFound = false

  for (let row = 6; row <= numRows && !dateFound; row++) {
    const dateCell = sheet.getRange(row, 4);
    const dateString = Utilities.formatDate(dateCell.getValue(), 'JST', 'M/d');
    console.log("date:", dateString);

    if (dateString === currentDateString) {
      dateFound = true
      for (let col = 5; col <= numCols; col++) {
        const newMemberName = sheet.getRange(5, col).getValue();
        const oldMemberName = sheet.getRange(row, col).getValue();
        const email = getEmailFromName_(oldMemberName);
        console.log("new:", newMemberName, "old:", oldMemberName, "email:", email);

        if (newMemberName && oldMemberName) {
          const subject = `\u23F0リマインド\u23F0【ご連絡】${dateString}分 ${newMemberName}の日報返信について(翌営業日まで)`;
          const htmlBody = HtmlService.createTemplateFromFile('Email');
          htmlBody.newMemberName = newMemberName;
          htmlBody.oldMemberName = oldMemberName;
          const finalHtmlBody = htmlBody.evaluate().getContent();
          console.log(subject)
          // console.log(finalHtmlBody)

          try {
            GmailApp.sendEmail(email, subject, "placeholder", {htmlBody: finalHtmlBody, name: "渡辺絵美子"})
          } catch(e) {
            console.error("Gmail error:", e)

          }
        }
      }
    }
  }
}