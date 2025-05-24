function showPendingCSVFiles() {
  const folderId = "1lHVhBeGtiEDfRnFEWxLPI2rqYJoEeiXz"; // your import folder
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const props = PropertiesService.getScriptProperties();
  const pendingFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.endsWith(".csv")) {
      const status = props.getProperty(name);
      if (!status || status === "failed") {
        pendingFiles.push(`${name} (${status || "unprocessed"})`);
      }
    }
  }

  if (pendingFiles.length === 0) {
    SpreadsheetApp.getUi().alert("🎉 All CSV files have been processed!");
  } else {
    SpreadsheetApp.getUi().alert(`⏳ Pending CSV files:\n\n${pendingFiles.join("\n")}`);
  }
}
