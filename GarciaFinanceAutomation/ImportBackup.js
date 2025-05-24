function replaceTransactionsFromExternalBackup() {
  const ui = SpreadsheetApp.getUi();

  // STEP 1: SETUP
  const backupFileId = "1o4gonetSwFeKzUk7nRW-mOAkYCUgeOb-FsXIgKqzPxk"; // <-- Update this!
  const sourceSheetName = "Transactions";           // Sheet name in the backup
  const targetSheetName = "Transactions";           // Your active sheet name

  const targetSS = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = targetSS.getSheetByName(targetSheetName);
  const backupSS = SpreadsheetApp.openById(backupFileId);
  const sourceSheet = backupSS.getSheetByName(sourceSheetName);

  if (!targetSheet || !sourceSheet) {
    ui.alert("❌ One of the sheets could not be found.");
    return;
  }

  // STEP 2: GET DATA FROM BACKUP
  const data = sourceSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert("⚠️ Backup has no transactions to import.");
    return;
  }

  // STEP 3: CLEAR TARGET SHEET (except header)
  targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();

  // STEP 4: COPY DATA INTO TARGET SHEET
  targetSheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));

  ui.alert(`✅ Imported ${data.length - 1} transactions from backup.`);
}
