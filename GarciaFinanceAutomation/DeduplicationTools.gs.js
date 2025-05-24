// === DEDUPLICATION TOOLS ===

function tagDuplicateTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tagCol = headers.indexOf("Tags");
  const dateCol = headers.indexOf("Date");
  const descCol = headers.indexOf("Description");
  const amountCol = headers.indexOf("Amount");
  const institutionCol = headers.indexOf("Institution");
  const dateAddedCol = headers.indexOf("Date Added");

  const seen = new Map();
  let invalidDateCount = 0;
  let duplicateCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][institutionCol] !== "Apple Card") continue;

    const rawDate = data[i][dateCol];
    const currentTags = data[i][tagCol] || "";

    if (!rawDate || Object.prototype.toString.call(rawDate) !== "[object Date]" || isNaN(rawDate.getTime())) {
      if (!currentTags.includes("InvalidDate")) {
        const updatedTags = currentTags ? currentTags + ", InvalidDate" : "InvalidDate";
        sheet.getRange(i + 1, tagCol + 1).setValue(updatedTags);
      }
      invalidDateCount++;
      continue;
    }

    const date = rawDate.toISOString().slice(0, 10);
    const desc = data[i][descCol];
    const amt = data[i][amountCol];
    const dateAdded = new Date(data[i][dateAddedCol]);
    const fingerprint = `${date}|${desc}|${amt}`;

    if (seen.has(fingerprint)) {
      const existing = seen.get(fingerprint);
      const timeDiff = Math.abs(dateAdded.getTime() - existing.dateAdded.getTime());

      if (timeDiff <= 5 * 60 * 1000) {
        if (!currentTags.includes("duplicate")) {
          const updatedTags = currentTags ? currentTags + ", duplicate" : "duplicate";
          sheet.getRange(i + 1, tagCol + 1).setValue(updatedTags);
          duplicateCount++;
        }
      }
    } else {
      seen.set(fingerprint, { index: i, dateAdded });
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Duplicate transactions tagged: ${duplicateCount}\n⏭️ Rows tagged as InvalidDate: ${invalidDateCount}`);
}

function confirmAndDeleteDuplicates() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Confirm Deletion", "This will permanently delete all rows tagged as 'duplicate' and with 'Apple Card' as Institution. Are you sure?", ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    deleteTaggedDuplicates();
    ui.alert("🗑️ Deleted all Apple Card rows tagged as 'duplicate'.");
  } else {
    ui.alert("🚫 Deletion canceled.");
  }
}

function deleteTaggedDuplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tagCol = headers.indexOf("Tags");
  const institutionCol = headers.indexOf("Institution");

  for (let i = data.length - 1; i > 0; i--) {
    const tags = (data[i][tagCol] || "").toString().toLowerCase();
    const institution = data[i][institutionCol];
    if (tags.includes("duplicate") && institution === "Apple Card") {
      sheet.deleteRow(i + 1);
    }
  }
}
