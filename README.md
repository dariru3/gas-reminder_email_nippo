# Email Reminder Program
gas-reminder_email_nippo

## Overview
This Google Apps Script program is designed to send reminder emails based on schedule and assignment data stored in a Google Spreadsheet. It works by matching dates and names from the spreadsheet, constructing personalized email content, and sending it to the respective individuals.

## Features
- **Date Matching**: Automatically checks for dates that match the current date and triggers email reminders accordingly.
- **Dynamic Email Construction**: Generates personalized email content using templates and recipient-specific information.
- **Error Handling**: Includes robust error handling and logging for easy debugging and maintenance.

## Configuration
Before using the script, configure the following in your Google Spreadsheet:
- `SHEET_CONFIG`: Object containing sheet names used in the program.
- `ROW_COL_CONFIG`: Object specifying key rows and columns, like start row for assignments, date column, etc.

## Functions
- `sendReminderEmails()`: Main function to orchestrate the email sending process.
- `getSheetData_(sheetname)`: Fetches sheet data such as number of rows and columns.
- `sendEmailsIfDateMatches_(sheetData, currentDate)`: Checks rows for the matching date and processes email sending.
- `sendEmailsForRow_(sheetData, row)`: Sends emails for a given row, iterating through columns.
- `getEmailFromName_(name)`: Retrieves email and additional name information from the configured sheet.
- `sendEmail_(newMemberName, email, name)`: Constructs and sends an email based on the given templates.

## Usage
1. Set up your Google Spreadsheet with the necessary data and configuration.
2. Open the Google Apps Script editor and paste the program.
3. Update the `SHEET_CONFIG` and `ROW_COL_CONFIG` with your specific sheet names and cell references.
4. Run the `sendReminderEmails` function either manually or set a trigger for automatic execution.

## Dependencies
- Google Apps Script environment.
- Google Spreadsheet with the required data structure.

## Limitations
- The program is tailored to specific spreadsheet structures and email formats.
- Error handling is basic and may need enhancement for broader cases.
