function findInvalidCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  const catSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  const txData = sheet.getDataRange().getValues();
  const headers = txData[0];
  const categoryCol = headers.indexOf("Category");
  const validCats = catSheet.getRange("A2:A").getValues().flat().filter(c => c);

  const invalids = [];

  for (let i = 1; i < txData.length; i++) {
    const category = txData[i][categoryCol];
    if (category && !validCats.includes(category)) {
      invalids.push(`Row ${i + 1}: "${category}"`);
    }
  }

  if (invalids.length > 0) {
    Logger.log("Invalid categories found:\n" + invalids.join("\n"));
  } else {
    Logger.log("✅ All categories are valid!");
  }
}