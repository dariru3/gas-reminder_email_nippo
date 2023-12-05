function getEmailFromName_(name) {
  const sheetname = SHEET_CONFIG.emailAddressSheetname;
  const lookupRange = ROW_COL_CONFIG.emailAddressLookupRange;
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  const emailNameList = referenceSheet.getRange(lookupRange).getValues();
  name = name.trim();

  const emailLookup = {}; // dictionary for easier lookup
  emailNameList.forEach(row => {
    emailLookup[row[0]] = {'email': row[1], 'name': row[2]}
  });

  if (name != "" && name in emailLookup) {
    return [emailLookup[name]['email'], emailLookup[name]['name']];
  } else if (name == '') {
    console.warn('Blank cell');
    return [null, null]
  } else {
    console.error("Email address not found.");
    return [null, null]
  }
}