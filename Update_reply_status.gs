function updateReplyStatus() {
  var replyStatusColIdx = 5; // 6th column zero-based for reply status
  var emailColIdx = 1;       // 2nd column zero-based for email address(es)

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // Build a map of email to row index for quick lookup
  var emailRowMap = {};
  for (var i = 1; i < data.length; i++) {
    var emails = (data[i][emailColIdx] || '').toString().toLowerCase().split(",");
    emails.forEach(function(email) {
      email = email.trim();
      if (email) emailRowMap[email] = i;
    });
  }

  // Search recent Inbox threads (adjust as needed)
  var threads = GmailApp.getInboxThreads(0, 100);

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();

    for (var m = 0; m < messages.length; m++) {
      var fromAddress = messages[m].getFrom().toLowerCase().match(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i);
      if (!fromAddress) continue;
      fromAddress = fromAddress[0].trim();

      if (emailRowMap.hasOwnProperty(fromAddress)) {
        var row = emailRowMap[fromAddress];
        var currentStatus = sheet.getRange(row+1, replyStatusColIdx+1).getValue();
        if (currentStatus.toString().toLowerCase() !== "reply received") {
          sheet.getRange(row+1, replyStatusColIdx+1).setValue("Reply Received");
        }
      }
    }
  }
}

/**
 * Uncomment and run this function once to create a time-driven trigger
 * that will run updateReplyStatus() daily (or adjust interval).
 */
/*
function createUpdateReplyTrigger() {
  ScriptApp.newTrigger("updateReplyStatus")
    .timeBased()
    .everyDays(1)      // Change frequency as needed: everyHours(1), everyMinutes(30), etc.
    .atHour(2)         // Run at 2 am server time - adjust as per preference
    .create();
}
*/
