function addEventToCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
  const numRows = sheet.getLastRow();
  const numCols = sheet.getLastColumn();

  // Replace with your name and calendar ID
  let yourName = 'Villalobos Daryl';
  yourName = yourName.toUpperCase();
  const calendarId = 'daryl.villalobos@link-cc.co.jp';
  const calendar = CalendarApp.getCalendarById(calendarId);

  for (let row = 6; row <= numRows; row++) {
    const dateCell = sheet.getRange(row, 4);
    if(!dateCell){
      continue
    }

    for (let col = 5; col <= numCols; col++) {
      const newMemberName = sheet.getRange(5, col).getValue();
      const oldMemberName = sheet.getRange(row, col).getValue();
      // console.log("Name:", oldMemberName)

      if (oldMemberName === yourName) {
        const eventDate = new Date(dateCell.getValue());
        const eventTitle = `Reply to ${newMemberName}`;
        console.log("Date:", eventDate, "Title:", eventTitle);

        const event = calendar.createAllDayEvent(eventTitle, eventDate);
        console.log('Event created: ' + event.getTitle());
      }
    }
  }
}
