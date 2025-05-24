function insertSampleEventTag() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EventTags")
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet("EventTags");

  sheet.clear(); // remove anything that's confusing it
  sheet.appendRow(["Event Name", "Start Date", "End Date", "Accounts Used", "Tag", "Description", "Categories (optional)"]);
  sheet.appendRow([
    "Weekend in Portland",
    "2025-05-02",
    "2025-05-05",
    "Apple Card, Joint Checking",
    "PortlandTrip2025",
    "Fun trip with spouse",
    "Restaurants, Entertainment, Parking"
  ]);
}
