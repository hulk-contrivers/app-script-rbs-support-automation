const cidPattern = /cid:[^\s]+/g;
function logSupportEmails() {
    // Define the email address to track
    const supportEmail = 'support+rbs@contrivers.com';
    
    // Access the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Define the date filter (e.g., emails from the past 7 days)
    const dateFilter = new Date();
    dateFilter.setDate(dateFilter.getDate() - 7); // Adjust this to the desired time range
    
    // Format the date for Gmail query
    const formattedDate = Utilities.formatDate(dateFilter, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    // Define the search query using the date filter
    const query = 'to:' + supportEmail + ' after:' + formattedDate;
    
    // Search for emails matching the query
    const threads = GmailApp.search(query);
    
    // Iterate through each email thread
    threads.forEach(thread => {
        const messages = thread.getMessages();
        
        messages.forEach(message => {
            // Get email details
            const dateReceived = message.getDate();
            const subject = message.getSubject();
            const fullBody = message.getPlainBody();
            
            // Extract the relevant portion of the email body from start to first "Wrote:"
            const bodySnippet = extractRelevantEmailPortion(fullBody);
            
            // Determine the month and year for the sheet name
            const month = Utilities.formatDate(dateReceived, Session.getScriptTimeZone(), 'MMMM yyyy');
            console.log(month)
            const sheetName = `${month}`;
            
            // Get or create the sheet for the month and year
            let sheet = spreadsheet.getSheetByName(sheetName);
            if (!sheet) {
                sheet = spreadsheet.insertSheet(sheetName);
                // Set up headers if this is a new sheet
                sheet.appendRow(['Date', 'Module', 'Description', 'Time']);
            }
            
            // Check if this subject (module) already exists in the sheet
            const lastRow = sheet.getLastRow();
 
            const existingModule = lastRow > 1 ? sheet.getRange(lastRow, 2).getValue() : '';
            
            // Only add a new row if the subject (module) is different
            if (subject !== existingModule) {
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
        return fullBody.substring(0, endIndex).replace(cidPattern, '').trim();
    }    
    // If "From:" is not found, return the full body
    return fullBody.replace(cidPattern, '').trim();
}