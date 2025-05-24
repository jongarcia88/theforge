const AUTOMATION_FUNCTIONS = [
  { name: "Clean Description Column", id: "cleanDescriptionColumn" },
  { name: "Apply AutoTag Rules", id: "applyTagRules" },
  { name: "Apply Recent Event Tags (Past 60 Days)", id: "applyRecentEventTags" },
  { name: "Add AutoCat Suggestions", id: "applyAutoCatSuggestions" }
];

// 🖥️ Show Sidebar
function showSchedulerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("DailyAutomationScriptSideBar")
    .setTitle("🛠️ Daily Automation Scheduler");
  SpreadsheetApp.getUi().showSidebar(html);
}

// 📋 Get available function names from this script
function getAvailableFunctions() {
  return Object.keys(globalThis)
    .filter(k => typeof globalThis[k] === 'function')
    .filter(k => !k.startsWith('_')) // optional: skip private/internal
    .sort();
}

// 📦 Config save/load
function getSavedSchedulerConfig() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty("DAILY_AUTOMATION_CONFIG");
  return raw ? JSON.parse(raw) : [];
}

function saveSchedulerConfig(config) {
  PropertiesService.getDocumentProperties().setProperty("DAILY_AUTOMATION_CONFIG", JSON.stringify(config));
  return true;
}

// 🚀 Triggered entry point
function runDailyAutomation() {
  const results = runAutomationOnce(false);
  sendDailyAutomationEmail(results);
  writeAutomationLogToSheet(results);
}

// 🧪 Manual or test execution
function runAutomationOnce(isTest = false) {
  const config = getSavedSchedulerConfig();
  const results = [];

  for (const step of config.sort((a, b) => a.order - b.order)) {
    const func = globalThis[step.id];
    if (typeof func !== "function") {
      results.push({ step: step.name, success: false, detail: "Function not found" });
      continue;
    }

    let success = false, resultDetail = "", error = "";
    for (let i = 0; i <= step.retries; i++) {
      try {
        const result = isTest ? "✓ Test Success" : func();
        success = true;
        resultDetail = result ?? "Done";
        break;
      } catch (e) {
        error = e.message;
        Utilities.sleep(500);
      }
    }

    results.push({
      step: step.name,
      success,
      detail: resultDetail,
      error: success ? "" : error
    });
  }

  return results;
}

// ✉️ Email summary
function sendDailyAutomationEmail(stepResults) {
  const email = Session.getActiveUser().getEmail();
  const subject = "💼 Daily Financial Automation Report";

  const bodyHtml = `
    <h2>🧾 Daily Financial Automation Report</h2>
    <ul>
      ${stepResults.map(res => `
        <li>
          <strong>${res.step}</strong>: 
          ${res.success ? '✅ <span style="color:green;">Success</span>' : '❌ <span style="color:red;">Failed</span>'}
          <br><small>${res.detail || ''}</small>
          ${res.error ? `<br><em>Error: ${res.error}</em>` : ""}
        </li>
      `).join("")}
    </ul>
    <p>📅 Completed at: ${new Date().toLocaleString()}</p>
  `;

  MailApp.sendEmail({
    to: email,
    subject,
    htmlBody: bodyHtml,
  });
}

// 🧾 Write log to sheet
function writeAutomationLogToSheet(stepResults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Daily Automation Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Daily Automation Log");
    logSheet.appendRow(["Timestamp", "Step", "Success", "Detail", "Error"]);
  }

  const timestamp = new Date();
  for (const res of stepResults) {
    logSheet.appendRow([
      timestamp,
      res.step,
      res.success ? "YES" : "NO",
      res.detail,
      res.error || "",
    ]);
  }
}

// ⏰ Install time trigger
function installSchedulerTrigger() {
  uninstallSchedulerTrigger();
  ScriptApp.newTrigger("runDailyAutomation")
    .timeBased()
    .everyDays(1)
    .atHour(6) // Customize time here
    .create();
}

// ❌ Remove existing scheduler trigger
function uninstallSchedulerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "runDailyAutomation") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// 🔁 Tiller AutoCat wrapper (optional)
function runAutoCat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoCatSheet = ss.getSheetByName("AutoCat");
  if (!autoCatSheet) throw new Error("AutoCat sheet not found.");

  const autoCat = TillerAutoCat;
  if (typeof autoCat?.applyRules !== "function") {
    throw new Error("TillerAutoCat engine not available.");
  }

  autoCat.applyRules();
  return "✅ AutoCat rules applied.";
}
