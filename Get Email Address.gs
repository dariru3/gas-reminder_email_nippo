function getEmailFromName_(name) {
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート2");
  const emailNameList = referenceSheet.getRange("A:B").getValues();
  name = name.trim();

  const emailLookup = {}; // dictionary for easier lookup
  emailNameList.forEach(row => {emailLookup[row[0]] = row[1]});
  // console.log(nameLookup)

  if (name != "" && name in emailLookup) {
    // console.log("Email:", emailLookup[name]);
    return emailLookup[name];
  } else {
    console.error("Email address not found.");
  }
}