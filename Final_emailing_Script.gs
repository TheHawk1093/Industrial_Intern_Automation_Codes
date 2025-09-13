/*("Company Name", "Company Email", "PoC Name", "Contact Number")*/
function sendCustomEmails() {
  var sentStatusColIdx = 4; // 5th column zero-based
  var replyStatusColIdx = 5; // 6th column zero-based for initial "No reply" marking

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // Map columns by header row (case insensitive, spaces removed)
  var headerRow = data[0].map(h => (h || '').toString().trim().toLowerCase().replace(/\s+/g,''));
  var colIndex = {
    companyName: headerRow.findIndex(h => h === "companyname"),
    companyEmail: headerRow.findIndex(h => h === "emailid" || h === "companyemail" || h === "companymail"),
    pocName: headerRow.findIndex(h => h === "poc" || h === "pocname" || h === "hr"),
    contactNumber: headerRow.findIndex(h => h === "mobilenoofpoc" || h === "mobileno" || h === "contactnumber")
  };

  for (var field in colIndex) {
    if (colIndex[field] === -1) throw new Error("Missing column: " + field);
  }

  // Find Gmail draft by subject
  var templateSubject = "Invitation: IIT Madras Engineering Design Dept. – Internship Drive (Dec 2025–May 2026)";
  var draftMessages = GmailApp.getDraftMessages();
  var draft = draftMessages.find(d => d.getSubject() === templateSubject);
  if (!draft) throw new Error("Draft with specified subject not found.");

  var htmlTemplate = draft.getBody();

  const firstPdfFileId = '1Nb5CJsQKOrqcaL-vtozcSmkS4apgz8C1';
  const firstPdfFile = DriveApp.getFileById(firstPdfFileId);
  const firstPdfBlob = firstPdfFile.getAs(MimeType.PDF); 

  const secondPdfFileId = '1OnR_7enFClQbx9R9ZTT4mmLn4lQwqMfr';
  const secondPdfFile = DriveApp.getFileById(secondPdfFileId);
  const secondPdfBlob = secondPdfFile.getAs(MimeType.PDF);

  const pngFileId = '16QKzsJsN2zuaSkYFLMRlgyxYxsJZSHv3';
  const pngFile = DriveApp.getFileById(pngFileId);
  const pngBlob = pngFile.getAs(MimeType.PNG);

  for (var i = 1; i < data.length; i++) {
    var companyName = data[i][colIndex.companyName];
    var email = data[i][colIndex.companyEmail];
    var pocName = data[i][colIndex.pocName];
    var phone = data[i][colIndex.contactNumber];

    if (!email || !companyName || !pocName || !phone) {
      sheet.getRange(i+1, sentStatusColIdx+1).setValue("Skipped: Missing fields");
      continue;
    }

    // Replace placeholders in HTML template
    var htmlBody = htmlTemplate
      .replace(/\[Company_Name\]/gi, companyName)
      .replace(/\[Name\]/gi, pocName)
      .replace(/\[Coordinator Name\]/gi, pocName)
      .replace(/\[Phone Number\]/gi, phone)
      .replace(/\[Position\]/gi, "");

    var attachments = [firstPdfBlob, secondPdfBlob, pngBlob];

    try {
      MailApp.sendEmail({
        to: email,
        cc: "gsaravana@iitm.ac.in",
        subject: templateSubject,
        htmlBody: htmlBody,
        attachments: attachments,
        body: "Please view this email in HTML."
      });
      sheet.getRange(i+1, sentStatusColIdx+1).setValue("Sent");
    } catch (e) {
      sheet.getRange(i+1, sentStatusColIdx+1).setValue("Delivery Failed: " + e.message);
    }

    // Initialize reply status if empty
    var replyStatus = data[i][replyStatusColIdx];
    if (!replyStatus || String(replyStatus).trim() === '') {
      sheet.getRange(i+1, replyStatusColIdx+1).setValue("No reply");
    }
  }
}
