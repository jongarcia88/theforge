function suggestCategoryFromButton() {
  const html = HtmlService.createHtmlOutputFromFile('CategorySidebar')
    .setTitle('💡 Suggest Category')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getCategorySidebarData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return { description: "" };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const descColIndex = headers.indexOf('Description') + 1;
  if (!descColIndex) return { description: "" };

  const description = sheet.getRange(row, descColIndex).getValue();
  return { description };
}

function saveSuggestedCategory(category) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const catColIndex = headers.indexOf('Category') + 1;
  if (!catColIndex) return;

  sheet.getRange(row, catColIndex).setValue(category);
}
