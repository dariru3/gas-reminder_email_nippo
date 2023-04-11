function getEmailFromName_(name) {
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート2");
  const emailNameList = referenceSheet.getRange("A:B").getValues();
  name = name.trim();

  const nameLookup = {}; // dictionary for easier lookup
  emailNameList.forEach(row => {nameLookup[row[0]] = row[1]});
  console.log(nameLookup)

  if (name != "" && name in nameLookup) {
    console.log("Name:", nameLookup[name]);
    return nameLookup[name];
  } else {
    console.error("Email address not found.");
  }
}