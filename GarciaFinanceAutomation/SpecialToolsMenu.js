// === Custom Menus ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const eventTagMenu = ui.createMenu("👛 Event Tagging")
    .addItem("📊 Generate Summary Report", "generateEventSummary")
    .addItem("📝 Edit Events (Editor Sidebar)", "showEventEditorSidebar")
    .addItem("🪰 Open Tag Manager (Sidebar)", "showEventTagSidebar")
    .addItem("🔄 Generate Event Report", "buildEventReportSheet") // ✅ Added here
    .addItem("📦 Export Shared Expenses", "exportSharedExpenses")
    .addItem("🔄 Refresh Event Report", "manualRefreshEventReport");

  const autoTaggingMenu = ui.createMenu("🏷️ AutoTagging")
    .addItem("Apply AutoTag Rules", "applyTagRules")
    .addItem("Clean Transaction Descriptions", "cleanDescriptionColumn")
    .addItem("Clean AutoCat Description contains", "cleanDescriptionColumnAutoCat")
    .addItem("Clean AutoTag Description contains", "cleanDescriptionColumnAutoTag")
    .addItem("🧮 Open AutoCat Rule Helper", "showAutoCatSidebar")
    .addItem("Suggest Category for Selection", "suggestCategoryFromButton");

  const duplicatesMenu = ui.createMenu("🔁 Duplicates")
    .addItem("🔍 Tag Duplicates", "tagDuplicateTransactions")
    .addItem("🗑️ Delete Tagged Duplicates", "confirmAndDeleteDuplicates")
    .addItem("🛠️ Fix Invalid Dates", "fixInvalidDates");

  const appleCardMenu = ui.createMenu("📅 Apple Card Import")
    .addItem("🔄 Show Pending CSV Files", "showPendingCSVFiles")
    .addItem("📋 Show All CSV Status", "showAllCSVStatus")
    .addItem("🧹 Reset Processed Flags", "showResetProcessedDialog")
    .addItem("📅 Manually Import CSV", "runManualImport");

  const tabsMenu = ui.createMenu("📂 Tabs Tools")
    .addItem("🚀 Open Tabs Import/Export Sidebar", "showTabsSidebar");

  ui.createMenu("🛠 Special Tools")
    .addSubMenu(autoTaggingMenu)
    .addSubMenu(duplicatesMenu)
    .addSubMenu(appleCardMenu)
    .addSubMenu(eventTagMenu)
    .addSubMenu(tabsMenu)
    .addItem("🛠️ Open Daily Automation Scheduler", "showSchedulerSidebar")
    .addItem("💾 Backup Transactions Sheet", "backupTransactionsSheet")
    .addToUi();
}

// === 📝 Event Sidebar
function showEventEditorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("EventEditor")
    .setTitle("📝 Edit Event Tags");
  SpreadsheetApp.getUi().showSidebar(html);
}

// === 👛 Tag Manager Sidebar
function showEventTagSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("EventTagSidebar")
    .setTitle("👛 Manage Event Tags");
  SpreadsheetApp.getUi().showSidebar(html);
}

// === 🧮 AutoCat Sidebar
function showAutoCatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("🧮 AutoCat Rule Helper");
  SpreadsheetApp.getUi().showSidebar(html);
}

// === 🛠️ Daily Automation Scheduler Sidebar
function showSchedulerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("SchedulerSidebar")
    .setTitle("🛠️ Daily Automation Scheduler");
  SpreadsheetApp.getUi().showSidebar(html);
}

// === 📂 Tabs Tools Sidebar
function showTabsSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("EventsReportSideBar")
    .setTitle("📂 Tabs Import/Export");
  SpreadsheetApp.getUi().showSidebar(html);
}

// === 🏷️ Event Tag Action Dispatcher
function handleEventTagAction(tag, action) {
  let result;
  switch (action) {
    case 'apply':
      result = applyEventTagsForTag(tag);
      break;
    case 'onlyIfMissing':
      result = applyEventTagsOnlyIfMissingForTag(tag);
      break;
    case 'preview':
      result = previewEventTagMatchesForTag(tag);
      break;
    case 'clear':
      result = clearAllEventTagsForTag(tag);
      break;
    default:
      result = "❌ Unknown action: " + action;
  }
  logEventAction(tag, action, result);
  return result;
}
