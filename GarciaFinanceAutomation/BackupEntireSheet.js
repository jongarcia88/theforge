function runBackupEntireSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceFileId = ss.getId();

    const backupFolderId = "1lHVhBeGtiEDfRnFEWxLPI2rqYJoEeiXz"; // Your backup folder ID
    const folder = DriveApp.getFolderById(backupFolderId);

    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    const backupName = `Spreadsheet Backup - ${dateStr}`;

    const copiedFile = DriveApp.getFileById(sourceFileId).makeCopy(backupName, folder);

    const successMessage = `✅ Spreadsheet backed up as "${backupName}" to folder "${folder.getName()}"`;
    console.log(successMessage);
    return successMessage;
  } catch (error) {
    const errorMessage = `❌ Backup failed: ${error.message}`;
    console.error(errorMessage);
    return errorMessage;
  }
}
