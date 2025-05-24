function backupTransactionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");
  if (!sheet) return;

  const backupFolderId = "1lHVhBeGtiEDfRnFEWxLPI2rqYJoEeiXz"; // Your backup folder
  const folder = DriveApp.getFolderById(backupFolderId);

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const backupName = `Transactions Backup - ${dateStr}`;

  // Step 1: Create a new blank spreadsheet
  const newFile = SpreadsheetApp.create(backupName);
  const newFileId = newFile.getId();

  // Step 2: Move the file
  DriveApp.getFileById(newFileId).moveTo(folder);

  // Step 3: Retry opening the spreadsheet
  let backupSpreadsheet = null;
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      backupSpreadsheet = SpreadsheetApp.openById(newFileId);
      if (backupSpreadsheet) break;
    } catch (e) {
      Utilities.sleep(1000); // Wait 1 second and try again
    }
  }

  if (!backupSpreadsheet) {
    SpreadsheetApp.getUi().alert("❌ Failed to open backup spreadsheet after multiple attempts.");
    return;
  }

  const backupSheet = backupSpreadsheet.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  SpreadsheetApp.getUi().alert(`✅ Backup created: ${backupName}`);
}
