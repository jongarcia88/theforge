// === APPLE CARD IMPORT SCRIPT WITH FULL LOGGING ===

function runManualImport() {
  processAppleCardCSV("Manual");
}

function runAutoImport() {
  processAppleCardCSV("Automated");
}

function processAppleCardCSV(triggerType) {
  const folderId = "1lHVhBeGtiEDfRnFEWxLPI2rqYJoEeiXz";
  const folder = DriveApp.getFolderById(folderId);
  const sheetName = "Transactions";
  const props = PropertiesService.getScriptProperties();
  const importIssuesSheetName = "Import Issues Log";
  const skippedRows = [];

  const logContent = [];
  let totalCsvFiles = 0, unprocessedCount = 0, failedCount = 0;

  // Count CSV files and statuses
  const allFilesForCount = folder.getFiles();
  while (allFilesForCount.hasNext()) {
    const f = allFilesForCount.next();
    if (f.getName().endsWith(".csv")) {
      totalCsvFiles++;
      const status = props.getProperty(f.getName());
      if (status !== "processed") {
        unprocessedCount++;
        if (status === "failed") failedCount++;
      }
    }
  }

  // Process one CSV file per execution
  const filesIter = folder.getFiles();
  while (filesIter.hasNext()) {
    const file = filesIter.next();
    const fileName = file.getName();
    if (!fileName.endsWith(".csv") || props.getProperty(fileName) === "processed") continue;

    let entryLog = `File: ${fileName}`;
    try {
      // Parse CSV
      const csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error("Transactions sheet not found.");
      const headers = sheet.getDataRange().getValues()[0];
      const numColumns = headers.length;
      const dateAdded = getFormattedTimestamp();

      // Build existing fingerprints
      const existingData = sheet.getDataRange().getValues();
      const existingFingerprints = {};
      for (let j = 1; j < existingData.length; j++) {
        const d = existingData[j][1];
        if (d instanceof Date && !isNaN(d)) {
          const dateKey = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
          const desc = existingData[j][2];
          const amt = existingData[j][5];
          const fp = `${dateKey}|${desc}|${amt}`;
          existingFingerprints[fp] = (existingFingerprints[fp] || 0) + 1;
        }
      }

      // Build CSV fingerprints
      const csvFingerprints = {};
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        const txDate = Utilities.parseDate(row[0], Session.getScriptTimeZone(), "MM/dd/yyyy");
        const desc = row[2];
        let amt = parseFloat(row[6]);
        amt = (row[5] === "Credit" || row[5] === "Payment") ? Math.abs(amt) : -Math.abs(amt);
        const iso = Utilities.formatDate(txDate, "UTC", "yyyy-MM-dd");
        const fp = `${iso}|${desc}|${amt}`;
        csvFingerprints[fp] = (csvFingerprints[fp] || 0) + 1;
      }

      // Append new rows and record skips
      const addedTracker = {};
      let addedCount = 0, skippedCount = 0;
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        const txDate = Utilities.parseDate(row[0], Session.getScriptTimeZone(), "MM/dd/yyyy");
        const desc = row[2];
        let amt = parseFloat(row[6]);
        amt = (row[5] === "Credit" || row[5] === "Payment") ? Math.abs(amt) : -Math.abs(amt);
        const iso = Utilities.formatDate(txDate, "UTC", "yyyy-MM-dd");
        const fp = `${iso}|${desc}|${amt}`;

        const exist = existingFingerprints[fp] || 0;
        const addedSoFar = addedTracker[fp] || 0;
        const allowed = csvFingerprints[fp] || 0;

        if (addedSoFar + exist >= allowed) {
          skippedRows.push({
            Date: Utilities.formatDate(txDate, Session.getScriptTimeZone(), "M/d/yyyy"),
            Description: desc,
            Amount: amt,
            Reason: "Skipped: Limit reached",
            File: fileName
          });
          skippedCount++;
          continue;
        }

        const weekStart = getTillerWeekStartDate(txDate);
        const monthStart = getFirstDayOfMonth(txDate);
        const newRow = new Array(numColumns).fill("");
        newRow[1] = txDate;
        newRow[2] = desc;
        newRow[5] = amt;
        newRow[6] = "Apple Card";
        newRow[7] = "Apple Card";
        newRow[8] = "Apple Card";
        newRow[9] = monthStart;
        newRow[10] = weekStart;
        newRow[14] = desc;
        newRow[15] = dateAdded;
        sheet.appendRow(newRow);

        addedTracker[fp] = addedSoFar + 1;
        addedCount++;
      }

      // Mark processed and archive
      props.setProperty(fileName, "processed");
      const archiveFolder = folder.getFoldersByName("Archived").hasNext()
        ? folder.getFoldersByName("Archived").next()
        : folder.createFolder("Archived");
      file.moveTo(archiveFolder);

      entryLog += `\nStatus: ✅ Processed successfully. Added=${addedCount}, Skipped=${skippedCount}`;
    } catch (e) {
      props.setProperty(fileName, "failed");
      entryLog += `\nStatus: ❌ Failed\nError: ${e.message}`;
    }

    entryLog += `\nTrigger: ${triggerType}\nTimestamp: ${getFormattedTimestamp()}`;
    logContent.push(entryLog);
    break; // only process one file per run
  }

  if (logContent.length > 0) {
    const details = formatSkippedRows(skippedRows);
    const fullLog = logContent.join("\n\n") + (details ? "\n\n" + details : "");
    writeToLogFile(fullLog);
    sendEmailSummary(fullLog);
    if (skippedRows.length) writeSkippedRowsToSheet(skippedRows, importIssuesSheetName);
  }
}

// === Helper Functions ===

function formatSkippedRows(rows) {
  if (!rows.length) return "";
  let output = "Details of Skipped Transactions:\n";
  output += "Date\tDescription\tAmount\tReason\tFile\n";
  rows.forEach(r => {
    output += `${r.Date}\t${r.Description}\t${r.Amount}\t${r.Reason}\t${r.File}\n`;
  });
  return output.trim();
}

function writeSkippedRowsToSheet(rows, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Description", "Amount", "Reason", "File"]);
  }
  rows.forEach(r => {
    sheet.appendRow([r.Date, r.Description, r.Amount, r.Reason, r.File]);
  });
}

function sendEmailSummary(body) {
  MailApp.sendEmail({
    to: "garcia.jonathan@gmail.com",
    subject: "Apple Card Import Summary",
    body: body
  });
}

function writeToLogFile(content) {
  const folder = DriveApp.getFolderById("1lHVhBeGtiEDfRnFEWxLPI2rqYJoEeiXz");
  const logFileName = "AppleCardImportLog.txt";
  const files = folder.getFilesByName(logFileName);
  if (files.hasNext()) {
    const logFile = files.next();
    const old = logFile.getBlob().getDataAsString();
    logFile.setContent(old + "\n\n" + content);
  } else {
    folder.createFile(logFileName, content);
  }
}

function getTillerWeekStartDate(date) {
  const d = new Date(date);
  d.setDate(d.getDate() - d.getDay());
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yy");
}

function getFirstDayOfMonth(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yy");
}

function getFormattedTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy H:mm:ss");
}
