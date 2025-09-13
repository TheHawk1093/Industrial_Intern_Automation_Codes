function sendReminders() {
  var replyStatusColIdx = 5; // 6th column zero-based for reply status
  var emailColIdx = 1;       // 2nd column zero-based for email address(es)

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // User sets the draft subject from which to get the reminder email body
  var reminderDraftSubject = "Reminder: IIT Madras Internship Invitation";

  const firstPdfFileId = '1Nb5CJsQKOrqcaL-vtozcSmkS4apgz8C1';
  const firstPdfFile = DriveApp.getFileById(firstPdfFileId);
  const firstPdfBlob = firstPdfFile.getAs(MimeType.PDF); 

  const secondPdfFileId = '1OnR_7enFClQbx9R9ZTT4mmLn4lQwqMfr';
  const secondPdfFile = DriveApp.getFileById(secondPdfFileId);
  const secondPdfBlob = secondPdfFile.getAs(MimeType.PDF);

  const pngFileId = '16QKzsJsN2zuaSkYFLMRlgyxYxsJZSHv3';
  const pngFile = DriveApp.getFileById(pngFileId);
  const pngBlob = pngFile.getAs(MimeType.PNG);

  // Find the draft with this subject to get the body
  var draftMessages = GmailApp.getDraftMessages();
  var reminderBody = "";
  for (var i = 0; i < draftMessages.length; i++) {
    if (draftMessages[i].getSubject() === reminderDraftSubject) {
      reminderBody = draftMessages[i].getBody();
      break;
    }
  }
  if (!reminderBody) {
    throw new Error("No draft found for reminder with subject: " + reminderDraftSubject);
  }

  // Plain text fallback if you want to strip HTML tags:
  // reminderBody = reminderBody.replace(/<[^>]+>/g, '');

  // We'll use the draft subject as reminder subject or define separately
  var reminderSubject = reminderDraftSubject;
  var attachments = [firstPdfBlob, secondPdfBlob, pngBlob];
  for (var i = 1; i < data.length; i++) {
    var replyStatus = (data[i][replyStatusColIdx] || '').toString().toLowerCase();
    var email = data[i][emailColIdx];

    if (email && replyStatus === "no reply") {
      try {
        MailApp.sendEmail({
          to: email,
          subject: reminderSubject,
          attachments: attachments,
          htmlBody: reminderBody, // Send as HTML email
          body: "This email requires HTML view." // fallback/plain text
        });
        sheet.getRange(i+1, replyStatusColIdx+1).setValue("Reminder Sent");
      } catch (e) {
        sheet.getRange(i+1, replyStatusColIdx+1).setValue("Reminder Failed: " + e.message);
      }
    }
  }
}
