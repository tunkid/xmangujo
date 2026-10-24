function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    var totalParticipants = lastRow > 0 ? lastRow - 1 : 0;
    var maxQuota = 50;

    if (totalParticipants >= maxQuota) {
      return ContentService.createTextOutput(JSON.stringify({
        "result": "full",
        "message": "Mohon maaf, kuota pendaftaran telah terpenuhi."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter;
    }

    sheet.appendRow([
      new Date(),
      data.instagram || '',
      data.whatsapp || '',
      data.alasan || ''
    ]);

    var remaining = maxQuota - (totalParticipants + 1);

    return ContentService.createTextOutput(JSON.stringify({
      "result": "success",
      "remaining": remaining
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var totalParticipants = lastRow > 0 ? lastRow - 1 : 0;
  var maxQuota = 50;
  var remaining = maxQuota - totalParticipants;
  if (remaining < 0) remaining = 0;

  return ContentService.createTextOutput(JSON.stringify({
    "count": totalParticipants,
    "remaining": remaining,
    "isFull": totalParticipants >= maxQuota
  })).setMimeType(ContentService.MimeType.JSON);
}
