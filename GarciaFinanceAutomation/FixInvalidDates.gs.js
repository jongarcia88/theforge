// === INVALID DATE FIXER ===

function fixInvalidDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol = headers.indexOf("Date");
  const tagCol = headers.indexOf("Tags");

  // Add columns if they don't exist
  let timestampCol = headers.indexOf("Fix Timestamp");
  let commentCol = headers.indexOf("Fix Comment");

  if (timestampCol === -1) {
    timestampCol = headers.length;
    sheet.getRange(1, timestampCol + 1).setValue("Fix Timestamp");
  }

  if (commentCol === -1) {
    commentCol = headers.length + (timestampCol === headers.length ? 1 : 0);
    sheet.getRange(1, commentCol + 1).setValue("Fix Comment");
  }

  const lastRow = data.length;
  let fixedCount = 0;
  let failedCount = 0;
  const ui = SpreadsheetApp.getUi();
  const clearedRows = [];

  const now = new Date();

  for (let i = 1; i < lastRow; i++) {
    const rowIdx = i + 1;
    const rowData = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dateCell = rowData[dateCol];
    const tagCell = (rowData[tagCol] || "").toLowerCase();

    if (tagCell.includes("invaliddate")) {
      if (typeof dateCell === "string" && dateCell.trim() !== "") {
        const parsed = new Date(dateCell);
        if (!isNaN(parsed.getTime())) {
          sheet.getRange(rowIdx, dateCol + 1).setValue(parsed);
          sheet.getRange(rowIdx, timestampCol + 1).setValue(now);
          sheet.getRange(rowIdx, commentCol + 1).setValue("Date corrected from text");
          sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).setBackground("#d9ead3"); // light green
          clearedRows.push(rowIdx);
          fixedCount++;
        } else {
          sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).setBackground("#f4cccc"); // light red
          failedCount++;
        }
      } else {
        sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).setBackground("#f4cccc"); // light red
        failedCount++;
      }
    }
  }

  if (clearedRows.length > 0) {
    ScriptApp.newTrigger("clearFixHighlights")
      .timeBased()
      .after(3 * 60 * 1000)
      .create();

    PropertiesService.getScriptProperties().setProperty("highlightRows", JSON.stringify(clearedRows));
  }

  ui.alert(`🛠️ Fix Summary:\n✅ Fixed: ${fixedCount} row(s)\n❌ Failed to fix: ${failedCount} row(s)`);
}

function clearFixHighlights() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  if (!sheet) return;

  const rows = JSON.parse(PropertiesService.getScriptProperties().getProperty("highlightRows") || "[]");

  for (const rowIdx of rows) {
    sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).setBackground(null);
  }

  PropertiesService.getScriptProperties().deleteProperty("highlightRows");
}
