function cleanDescriptionColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  if (!sheet) {
    throw new Error("Transactions sheet not found.");
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  const descColIndex = headers.indexOf("Description");
  const tagsColIndex = headers.indexOf("Tags");

  if (descColIndex === -1) {
    throw new Error("No 'Description' column found.");
  }
  if (tagsColIndex === -1) {
    throw new Error("No 'Tags' column found.");
  }

  let cleanedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const tagsValue = data[i][tagsColIndex];
    // Skip rows that already have the 'ImportedRY' tag
    if (tagsValue && tagsValue.toString().includes("ImportedRY")) {
      continue;
    }

    const original = data[i][descColIndex];
    if (original) {
      const cleaned = original
        .toString()
        .replace(/[\r\n]+/g, ' ')                      // Remove carriage returns and line breaks
        .replace(/\b(SQ|TST)\b/gi, "")                 // Remove standalone "SQ" and "TST"
        .replace(/^check card purchase\s*/i, "")      // Remove "Check Card Purchase" at start
        .replace(/#[0-9]+/g, "")                      // Remove hash-numbers like "#1234"
        .replace(/[^\w\s]/g, "")                      // Remove special characters
        .replace(/\s+/g, " ")                         // Collapse multiple spaces
        .trim();

      if (cleaned !== original) {
        sheet.getRange(i + 1, descColIndex + 1).setValue(cleaned);
        cleanedCount++;
      }
    }
  }

  return `✅ Cleaned ${cleanedCount} descriptions.`;
}


function cleanDescriptionColumnAutoCat() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AutoCat");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Transactions sheet not found.");
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  const descColIndex = headers.indexOf("Description Contains");

  if (descColIndex === -1) {
    SpreadsheetApp.getUi().alert("No 'Description Contains' column found.");
    return;
  }

  let cleanedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const original = data[i][descColIndex];
    if (original) {
      const cleaned = original
        .toString()
        .replace(/[\r\n]+/g, ' ')                      // Remove carriage returns and line breaks
        .replace(/\b(SQ|TST)\b/gi, "")                 // Remove standalone "SQ" and "TST"
        .replace(/^check card purchase\s*/i, "")       // Remove "Check Card Purchase" at start
        .replace(/#[0-9]+/g, "")                       // Remove hash-numbers like "#1234"
        .replace(/\b\d+\b/g, "")                       // Remove standalone numbers
        .replace(/[^\w\s]/g, "")                       // Remove special characters
        .replace(/\s+/g, " ")                          // Collapse multiple spaces
        .trim();

      if (cleaned !== original) {
        sheet.getRange(i + 1, descColIndex + 1).setValue(cleaned);
        cleanedCount++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Cleaned ${cleanedCount} descriptions!`);
}

function cleanDescriptionColumnAutoTag() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AutoTag");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("AutoTag sheet not found.");
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  const descColIndex = headers.indexOf("Description Contains");
  const descColIndex2 = headers.indexOf("Exclude If Contains");

  if (descColIndex === -1) {
    SpreadsheetApp.getUi().alert("No 'Description Contains' column found.");
    return;
  }

  if (descColIndex2 === -1) {
    SpreadsheetApp.getUi().alert("No 'Exclude If Contains' column found.");
    return;
  }

  let cleanedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const original = data[i][descColIndex];
    if (original) {
      const cleaned = original
        .toString()
        .replace(/[\r\n]+/g, ' ')                      // Remove carriage returns and line breaks
        .replace(/\b(SQ|TST)\b/gi, "")                 // Remove standalone "SQ" and "TST"
        .replace(/^check card purchase\s*/i, "")       // Remove "Check Card Purchase" at start
        .replace(/#[0-9]+/g, "")                       // Remove hash-numbers like "#1234"
        .replace(/\b\d+\b/g, "")                       // Remove standalone numbers
        .replace(/[^\w\s]/g, "")                       // Remove special characters
        .replace(/\s+/g, " ");                         // Collapse multiple spaces (no trim)

      if (cleaned !== original) {
        sheet.getRange(i + 1, descColIndex + 1).setValue(cleaned);
        cleanedCount++;
      }
    }
  }

  for (let i = 1; i < data.length; i++) {
    const original = data[i][descColIndex2];
    if (original) {
      const cleaned = original
        .toString()
        .replace(/[\r\n]+/g, ' ')
        .replace(/^check card purchase\s*/i, "")
        .replace(/#[0-9]+/g, "")
        .replace(/\b\d+\b/g, "")
        .replace(/[^\w\s,]/g, "")                      // Preserve commas
        .replace(/\s+/g, " ");                         // Collapse spaces only

      if (cleaned !== original) {
        sheet.getRange(i + 1, descColIndex2 + 1).setValue(cleaned);
        cleanedCount++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Cleaned ${cleanedCount} descriptions!`);
}

