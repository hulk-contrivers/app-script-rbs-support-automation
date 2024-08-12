/* 
 * This script performs the following tasks:
 * 1. Defines a pattern to match CID (Content-ID) in email bodies.
 * 2. Logs support emails sent to a specific email address within the past 7 days.
 * 3. Accesses the active Google Spreadsheet to log the email details.
 * 4. Formats the date for Gmail query to filter emails from the past 7 days.
 * 5. Searches for emails matching the query.
 * 6. Iterates through each email thread and processes the messages.
 * 
 * Author: [Your Name]
 * Date: [Current Date]
 */
const cidPattern = /cid:[^\s]+/g;

function logSupportEmails() {
  // Define the email address to track
  const supportEmail = "support+rbs@contrivers.com";

  // Access the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Define the date filter (e.g., emails from the past 7 days)
  const dateFilter = new Date();
  dateFilter.setDate(dateFilter.getDate() - 7); // Adjust this to the desired time range

  // Format the date for Gmail query
  const formattedDate = Utilities.formatDate(
    dateFilter,
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );

  // Define the search query using the date filter
  const query = "to:" + supportEmail + " after:" + formattedDate;

  // Search for emails matching the query
  const threads = GmailApp.search(query);

  // Iterate through each email thread
  threads.forEach((thread) => {
    const messages = thread.getMessages();

    messages.forEach((message) => {
      // Get email details
      const dateReceived = message.getDate();
      const subject = message.getSubject();
      const fullBody = message.getPlainBody();

      // Extract the relevant portion of the email body from start to first "Wrote:"
      const bodySnippet = extractRelevantEmailPortion(fullBody);

      // Determine the month and year for the sheet name
      const month = Utilities.formatDate(
        dateReceived,
        Session.getScriptTimeZone(),
        "MMMM yyyy"
      );
      console.log(month);
      const sheetName = `${month}`;

      // Get or create the sheet for the month and year
      let sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        // Set up headers if this is a new sheet
        sheet.appendRow(["Date", "Module", "Description", "Time"]);
      }

      // Assuming 'sheet' is already defined and refers to the active sheet
      const lastRow = sheet.getLastRow(); // Get the total number of rows
      const moduleColumnIndex = 2; // Assuming the "Module" column is the second column
      let modules = []; // Array to store module values

      // Loop through each row and get the value from the "Module" column
      for (let row = 1; row <= lastRow; row++) {
        let moduleValue = sheet.getRange(row, moduleColumnIndex).getValue();
        modules.push(moduleValue);
      }

      // Check if the subject exists in the modules array
      if (!modules.includes(subject)) {
        const startTime = new Date(); // Start timer for reading email

        // Simulate reading the email
        Utilities.sleep(5000); // Simulates 5 seconds of reading time

        const endTime = new Date(); // End timer
        const timeSpent = (endTime - startTime) / 1000; // Time spent in seconds

        // Log details in the sheet
        sheet.appendRow([dateReceived, subject, bodySnippet, timeSpent]);
      }
    });
  });
}
function extractRelevantEmailPortion(fullBody) {
  const endIndex = fullBody.indexOf("wrote:");

  if (endIndex !== -1) {
    return fullBody.substring(0, endIndex).replace(cidPattern, "").trim();
  }
  // If "From:" is not found, return the full body
  return fullBody.replace(cidPattern, "").trim();
}
