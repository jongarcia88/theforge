// ✅ Event Tagging Full Updated Script (Final Polished Version)

function getEventTagData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EventTags");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (!Array.isArray(data) || data.length < 2) return [];

  const rows = data.slice(1).filter(row => row && row.length >= 5 && row[4]);
  return JSON.parse(JSON.stringify(rows));
}

function saveOrUpdateEvent(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EventTags");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const tagIndex = headers.indexOf("Tag");

  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][tagIndex] === data.tag) {
      rowIndex = i;
      break;
    }
  }

  const rowData = [
    data.eventName,
    data.startDate,
    data.endDate,
    data.accountsUsed,
    data.tag,
    data.description,
    data.categories,
    data.excludeTags
  ];

  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

function getAccountsAndCategories() {
  const txSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  const data = txSheet.getDataRange().getValues();
  const headers = data[0];
  const accountCol = headers.indexOf("Account");
  const categoryCol = headers.indexOf("Category");

  const accounts = new Set();
  const categories = new Set();

  for (let i = 1; i < data.length; i++) {
    if (data[i][accountCol]) accounts.add(data[i][accountCol]);
    if (data[i][categoryCol]) categories.add(data[i][categoryCol]);
  }

  return {
    accounts: [...accounts].sort(),
    categories: [...categories].sort()
  };
}

function getEventTagsList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EventTags");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tagIndex = 4; // Column E = Tag
  return data.slice(1)
    .map(row => row[tagIndex])
    .filter(tag => tag && tag.trim() !== "");
}

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

function logEventAction(tag, action, result) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EventTag Log")
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet("EventTag Log");

  logSheet.appendRow([
    new Date(),
    tag,
    action,
    result
  ]);
}

function applyEventTagsForTag(tag) {
  const result = processEventTag(tag, { previewOnly: false, onlyIfMissing: false });
  return result.tagged || 0;
}

function applyEventTagsOnlyIfMissingForTag(tag) {
  const result = processEventTag(tag, { previewOnly: false, onlyIfMissing: true });
  return result.tagged || 0;
}

function previewEventTagMatchesForTag(tag) {
  const result = processEventTag(tag, { previewOnly: true, onlyIfMissing: false });
  return `🔍 Preview: ${result.tagged} transaction(s) would be tagged with '${tag}'`;
}

function clearAllEventTagsForTag(tag) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
  if (!sheet) return "❌ Transactions sheet not found";

  const transactions = sheet.getDataRange().getValues();
  const headers = transactions[0];
  const tagCol = headers.indexOf("Tags");

  let cleared = 0;

  for (let i = 1; i < transactions.length; i++) {
    const rowTag = (transactions[i][tagCol] || "").toString();
    if (rowTag.includes(tag)) {
      const updatedTag = rowTag
        .split(",")
        .map(t => t.trim())
        .filter(t => t !== tag)
        .join(", ");
      sheet.getRange(i + 1, tagCol + 1).setValue(updatedTag);
      cleared++;
    }
  }

  return `🩹 Cleared ${cleared} tag(s) matching "${tag}"`;
}

function processEventTag(tag, options) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
    const eventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EventTags");
    if (!sheet || !eventSheet) return { tagged: 0, skipped: 0, preview: options.previewOnly };

    const eventData = eventSheet.getDataRange().getValues();
    const events = eventData.slice(1);
    const event = events.find(r => r[4] === tag);
    if (!event) return { tagged: 0, skipped: 0, error: `Event tag '${tag}' not found` };

    const [eventName, startDate, endDate, accountsUsed, , , categoryFilter, excludeTagsStr] = event;

    const accounts = (accountsUsed || "").split(",").map(s => s.trim());
    const categories = (categoryFilter || "").split(",").map(s => s.trim()).filter(Boolean);
    const excludeTags = (excludeTagsStr || "").split(",").map(s => s.trim().toLowerCase()).filter(Boolean);

    const transactions = sheet.getDataRange().getValues();
    const headers = transactions[0];

    const dateCol = headers.indexOf("Date");
    const accountCol = headers.indexOf("Account");
    const categoryCol = headers.indexOf("Category");
    const tagCol = headers.indexOf("Tags");

    let tagged = 0;
    let skipped = 0;

    for (let i = 1; i < transactions.length; i++) {
      const row = transactions[i];
      const rowDate = new Date(row[dateCol]);
      const inRange = rowDate >= new Date(startDate) && rowDate <= new Date(endDate);
      const accountMatch = accounts.includes(row[accountCol]);
      const categoryMatch = categories.length === 0 || categories.includes(row[categoryCol]);

      const currentTags = (row[tagCol] || "").split(",").map(t => t.trim());
      const alreadyTagged = currentTags.includes(tag);
      const hasExcludedTag = excludeTags.some(ex => currentTags.map(t => t.toLowerCase()).includes(ex));

      if (inRange && accountMatch && categoryMatch) {
        if (hasExcludedTag) {
          skipped++;
          continue;
        }

        if (options.previewOnly) {
          tagged++;
        } else if (!alreadyTagged && (!options.onlyIfMissing || currentTags.join("") === "")) {
          currentTags.push(tag);
          const updated = currentTags.filter(Boolean).join(", ");
          sheet.getRange(i + 1, tagCol + 1).setValue(updated);
          tagged++;
        } else {
          skipped++;
        }
      }
    }

    return { tagged, skipped, preview: options.previewOnly };
  } catch (err) {
    return { tagged: 0, skipped: 0, error: "❌ Error: " + err.message };
  }
}

function applyRecentEventTags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("EventTags");
  if (!sheet) throw new Error("EventTags sheet not found.");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const tagCol = headers.indexOf("Tag");
  const startDateCol = headers.indexOf("Start Date");

  if (tagCol === -1 || startDateCol === -1) {
    throw new Error("Missing 'Tag' or 'Start Date' columns in EventTags sheet.");
  }

  const today = new Date();
  const cutoffDate = new Date(today);
  cutoffDate.setDate(today.getDate() - 60);

  let totalTaggedTransactions = 0;
  const tagResults = [];

  for (let i = 1; i < data.length; i++) {
    const tag = data[i][tagCol];
    const start = new Date(data[i][startDateCol]);

    if (tag && start >= cutoffDate && start <= today) {
      try {
        const taggedCount = applyEventTagsForTag(tag);

        if (taggedCount > 0) {
          tagResults.push(`${tag} (${taggedCount} txns)`);
          totalTaggedTransactions += taggedCount;
        } else {
          console.log(`ℹ️ No transactions tagged for event "${tag}".`);
        }
      } catch (e) {
        console.warn(`⚠️ Could not apply tag "${tag}": ${e.message}`);
      }
    }
  }

  return tagResults.length > 0
    ? `✅ Applied event tags:\n${tagResults.join("\n")}\nTotal transactions tagged: ${totalTaggedTransactions}`
    : `ℹ️ No event tags applied (within last 60 days).`;
}

function showToast(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, "Event Tagger");
}
