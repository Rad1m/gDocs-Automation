function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Adjust if sheet name differs
  const lastRow = sheet.getLastRow();

  console.log(`Form submitted. Processing row: ${lastRow}`);

  // Update status in column B and column C
  updateStatus(sheet, lastRow);

  // Generate and save the gDOC, then update the sheet and send the email
  saveDocAndSendEmail(lastRow);
}

function updateStatus(sheet, row) {
  console.log(`Updating status for row: ${row}`);
  // Set status in column B to 'New' and column C to 'NOT ASSIGNED'
  sheet.getRange(row, 2).setValue('New'); // Update status in column B
  sheet.getRange(row, 3).setValue('NOT ASSIGNED'); // Update status in column C
  console.log(`Status updated for row: ${row}`);
}

function generateAndSaveDoc(row) {
  try {
    console.log(`Generating Google Doc for row: ${row}`);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Adjust if sheet name differs
    const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get the row data
    console.log(`Row data: ${data}`);

    const headers = sheet.getRange(1, 8, 1, sheet.getLastColumn() - 7).getValues()[0]; // Dynamically get headers from column H to the last column
    console.log(`Headers: ${headers}`);

    const title = `${data[10]}`; //  title based on column K (index 10, as column indices are 0-based)
    const docName = `${new Date().toISOString().split('T')[0]}_${title}`; // Today's date (YYYY-MM-DD) + title; // Today's date + title
    console.log(`Creating document with name: ${docName}`);

    const doc = DocumentApp.create(docName); // Name gDOC based on column K
    const body = doc.getBody();

    // Add description from the first row as a Heading 1
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    console.log(`Added heading: ${headers[1]}`);

    // Include dynamic fields from columns B and C
    body.replaceText('{{Status}}', data[1]); // Dynamic link to Column B
    body.replaceText('{{Assigned to}}', data[2]); // Dynamic link to Column C
    console.log(`Replaced placeholders for Status and Assigned to.`);

    // Include all fields from columns H to AC
    headers.forEach((header, index) => {
      body.appendParagraph(String(header || 'Unknown Header')).setHeading(DocumentApp.ParagraphHeading.HEADING2); // Use header as a sub-heading
      body.appendParagraph(String(data[index + 7] || 'No data')); // Use the actual field as normal text
    });
    console.log(`Added all fields from columns H to AC.`);

    // Save the gDOC in the folder
    const folder = DriveApp.getFolderById('1mvJ04rzU7G4Lfg0r5S6L9Xc5dds9sISP'); // Use the provided folder ID
    const file = DriveApp.getFileById(doc.getId());
    // folder.add(file); // Deprecated line removed
    file.moveTo(folder); // Move the file to the specified folder // Remove the file from the root folder
    console.log(`Document saved in folder: ${folder.getName()}`);

    return file.getUrl(); // Return the link to the saved gDOC
  } catch (error) {
    console.error(`Error in generateAndSaveDoc: ${error.message}`);
    throw error; // Re-throw error to stop execution
  }
}

function saveDocAndSendEmail(row) {
  console.log(`Saving document and sending email for row: ${row}`);
  const docLink = generateAndSaveDoc(row); // Generate and save the gDOC, then get its link
  const docName = `Proposal_${SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1').getRange(row, 11).getValue()}`;
  console.log(`Document link: ${docLink}`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // Adjust if sheet name differs
  
  // Update column G with the gDOC link
  sheet.getRange(row, 7).setValue(docLink); // Column G is the 7th column
  console.log(`Document link saved to column G for row: ${row}`);

  // Send email with the gDOC link
  const tablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tables'); // Get the sheet named 'Tables'
  const emails = tablesSheet.getRange(2, 5, tablesSheet.getLastRow() - 1, 1).getValues().flat(); // Get all emails from column E starting from row 2
  console.log(`Emails to send: ${emails}`); // Get all emails from column E starting from row 2

  emails.forEach(email => {
    if (!email || !email.includes('@')) {
      console.warn(`Skipping invalid email: ${email}`);
      return;
    }
    console.log(`Sending email to: ${email}`);
    const subject = `New proposal: ${docName}`;
    const body = `Please find your proposal at the following link: ${docLink}`;

    GmailApp.sendEmail(email, subject, body); // Send the email with the link
    console.log(`Email sent to: ${email}`);
  });
  console.log(`All emails sent for row: ${row}`);
}
