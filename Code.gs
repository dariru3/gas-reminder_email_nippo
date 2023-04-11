function sendReminderEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
        // const oldMemberEmail = sheet.getRange(row, col).getValue();
        const newMemberName = sheet.getRange(5, col).getValue();
        const oldMemberName = sheet.getRange(row, col).getValue();
        console.log("new:", newMemberName, "old:", oldMemberName);

        if (newMemberName && oldMemberName) {
          const subject = `Reminder: Welcome New Member ${newMemberName}`;
          const body = `Hi ${oldMemberName},
This is a reminder for you to send a welcome email to our new member, ${newMemberName}. 
Please take a moment to introduce yourself and make them feel welcomed to the group.

Best regards,
Your Group`;
          console.log(subject)
          console.log(body)

          // MailApp.sendEmail(oldMemberEmail, subject, body);
        }
      }
    }
  }
}

/*
function createTrigger() {
  ScriptApp.newTrigger('sendReminderEmails')
    .timeBased()
    .everyDays(1)
    .create();
}
*/