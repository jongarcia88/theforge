function showResetProcessedDialog() {
  const html = HtmlService.createHtmlOutputFromFile("ResetProcessedDialog")
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Reset Processed CSV Files");
}

function getProcessedCSVFiles() {
  const folderId = "1lHVhBeGtiEDfRnFEWxLPI2rqYJoEeiXz"; // your import folder ID
  const folder = DriveApp.getFolderById(folderId);
  const props = PropertiesService.getScriptProperties();
  const files = folder.getFiles();
  const processed = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.toLowerCase().endsWith(".csv") && props.getProperty(name) === "processed") {
      processed.push(name);
    }
  }
  return processed;
}

function resetProcessedFlag(fileName) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(fileName);
  return `✅ Flag reset for: ${fileName}`;
}

function showAllCSVStatus() {
  const folderId = "1lHVhBeGtiEDfRnFEWxLPI2rqYJoEeiXz"; // your Apple Card CSV folder ID
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const props = PropertiesService.getScriptProperties();

  const statusLines = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.toLowerCase().endsWith(".csv")) {
      const status = props.getProperty(name) || "unprocessed";
      statusLines.push(`${name} — ${status}`);
    } else {
      statusLines.push(`${name} — ❗ Not a .csv file`);
    }
  }

  const message = statusLines.length
    ? "📋 CSV File Status:\n\n" + statusLines.join("\n")
    : "📂 No files found in your import folder.";

  SpreadsheetApp.getUi().alert(message);
}