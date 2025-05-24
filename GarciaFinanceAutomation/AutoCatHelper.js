function showAutoCatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("🧮 AutoCat Rule Helper");
  SpreadsheetApp.getUi().showSidebar(html);
}

function cleanDescription(text, removeLocation = true, preserveNumbers = true) {
  let cleaned = (text || "")
    .replace(/^check card purchase\s*/i, "")
    .replace(/\b(SQ|TST)\b/gi, "")
    .replace(/#[0-9]+/g, "")                      // Remove hash numbers like #1234
    .replace(/\d{5}(?:-\d{4})?/g, "")              // Remove zip codes
    .replace(/[^\w\s]/g, "")                       // Remove non-alphanumeric
    .replace(/\s+/g, " ")                          // Collapse spaces
    .trim();

  if (!preserveNumbers) {
    cleaned = cleaned.replace(/\b\d+\b/g, "");     // Optionally remove isolated numbers (if desired)
  }

  if (!removeLocation) return cleaned;

  const locationWords = [
    // Full State Names
    "ALABAMA","ALASKA","ARIZONA","ARKANSAS","CALIFORNIA","COLORADO","CONNECTICUT","DELAWARE",
    "FLORIDA","GEORGIA","HAWAII","IDAHO","ILLINOIS","INDIANA","IOWA","KANSAS","KENTUCKY",
    "LOUISIANA","MAINE","MARYLAND","MASSACHUSETTS","MICHIGAN","MINNESOTA","MISSISSIPPI",
    "MISSOURI","MONTANA","NEBRASKA","NEVADA","NEW HAMPSHIRE","NEW JERSEY","NEW MEXICO",
    "NEW YORK","NORTH CAROLINA","NORTH DAKOTA","OHIO","OKLAHOMA","OREGON","PENNSYLVANIA",
    "RHODE ISLAND","SOUTH CAROLINA","SOUTH DAKOTA","TENNESSEE","TEXAS","UTAH","VERMONT",
    "VIRGINIA","WASHINGTON","WEST VIRGINIA","WISCONSIN","WYOMING",
    // State Abbreviations
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY",
    "LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND",
    "OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY",
    // Major US Cities
    "SEATTLE","MIAMI","NEW YORK","LOS ANGELES","SAN FRANCISCO","CHICAGO","DALLAS",
    "AUSTIN","ORLANDO","BOSTON","DENVER","ATLANTA","PHOENIX","LAS VEGAS","PORTLAND",
    "SAN DIEGO","PHILADELPHIA","HOUSTON","DETROIT",
    // Countries
    "USA","CANADA","MEXICO","FRANCE","GERMANY","SPAIN","ITALY","UK","UNITED KINGDOM",
    "AUSTRALIA","NETHERLANDS","JAPAN","CHINA","SINGAPORE","BRAZIL","DUBAI","ARGENTINA",
    "SOUTH AFRICA","THAILAND","INDIA","NEW ZEALAND","SWITZERLAND","SWEDEN","NORWAY",
    "DENMARK","IRELAND","GREECE","TURKEY","CROATIA",
    // Continents
    "EUROPE","ASIA","AFRICA","SOUTH AMERICA","NORTH AMERICA"
  ];

  const regex = new RegExp(`\\b(?:${locationWords.join("|")})\\b`, "gi");
  return cleaned.replace(regex, "").replace(/\s+/g, " ").trim();
}

function normalizeRuleKey(category, description, clean = true, removeLoc = true, preserveNumbers = true) {
  const cat = (category || "").toLowerCase().trim();
  const desc = clean ? cleanDescription(description, removeLoc, preserveNumbers).toLowerCase().trim() : (description || "").toLowerCase().trim();
  return `${cat}|${desc}`;
}

function getSuggestedAutoCatRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName("Transactions");
  const ruleSheet = ss.getSheetByName("AutoCat");
  const skippedSheet = ss.getSheetByName("SkippedAutoCatRules");

  const txData = txSheet.getDataRange().getValues();
  const ruleData = ruleSheet.getDataRange().getValues();
  const headers = txData[0];
  const descCol = headers.indexOf("Description");
  const catCol = headers.indexOf("Category");

  const ruleHeaders = ruleData[0];
  const ruleDescCol = ruleHeaders.indexOf("Description Contains");

  const matchers = ruleData.slice(1)
    .map(r => (r[ruleDescCol] || "").toLowerCase().trim())
    .filter(Boolean);

  const skipped = skippedSheet && skippedSheet.getLastRow() > 1
    ? new Set(skippedSheet.getRange(2, 1, skippedSheet.getLastRow() - 1, 2)
        .getValues()
        .map(r => normalizeRuleKey(r[0], r[1])))
    : new Set();

  const existingRules = new Set(ruleData.slice(1)
    .filter(r => r[0] && r[ruleDescCol])
    .map(r => normalizeRuleKey(r[0], r[ruleDescCol])));

  const suggestions = {};

  for (let i = 1; i < txData.length; i++) {
    const desc = (txData[i][descCol] || "").trim();
    const cat = (txData[i][catCol] || "").trim();
    if (!desc || !cat) continue;

    const ruleKey = normalizeRuleKey(cat, desc);
    if (existingRules.has(ruleKey) || skipped.has(ruleKey) || suggestions[ruleKey]) continue;

    const lowerDesc = desc.toLowerCase();
    const matchesExisting = matchers.some(rule => lowerDesc.includes(rule));
    if (matchesExisting) continue;

    suggestions[ruleKey] = {
      category: cat,
      description: cleanDescription(desc),
      matchCount: 1,
      matchedRule: "None"
    };
  }

  return Object.values(suggestions);
}

function addRulesToAutoCat(rules, clean = true, removeLoc = true, preserveNumbers = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AutoCat_Proposed") || ss.getSheetByName("AutoCat").copyTo(ss);

  if (!sheet.getName().includes("AutoCat_Proposed")) {
    sheet.setName("AutoCat_Proposed");
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hasTimestamp = headers.includes("Timestamp Added");
  const timestampCol = hasTimestamp ? headers.indexOf("Timestamp Added") + 1 : null;

  const existingKeys = new Set(sheet.getDataRange().getValues().slice(1)
    .map(r => normalizeRuleKey(r[0], r[1])));

  const timestamp = new Date();

  const toAdd = rules.map(rule => {
    const cleanedDesc = clean ? cleanDescription(rule.description, removeLoc, preserveNumbers) : rule.description;
    return {
      category: rule.category,
      description: cleanedDesc,
      key: normalizeRuleKey(rule.category, rule.description, clean, removeLoc, preserveNumbers)
    };
  }).filter(r => !existingKeys.has(r.key));

  if (toAdd.length === 0) return "✅ No new rules to add.";

  const values = toAdd.map(r => {
    const base = [r.category, r.description, "", "", "", ""];
    return hasTimestamp ? [...base, timestamp] : base;
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);

  if (hasTimestamp && timestampCol) {
    sheet.getRange(startRow, timestampCol, values.length).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  }

  return `✅ ${values.length} new rule(s) added to '${sheet.getName()}'.`;
}

function applyAutoCatSuggestions() {
  const suggestions = getSuggestedAutoCatRules();
  if (!Array.isArray(suggestions) || suggestions.length === 0) {
    return "ℹ️ No AutoCat suggestions found.";
  }

  return addRulesToAutoCat(suggestions, true, true, true); // Now using preserveNumbers
}

function mergeAutoCatProposedIntoAutoCat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const proposedSheet = ss.getSheetByName("AutoCat_Proposed");
  const autoCatSheet = ss.getSheetByName("AutoCat");

  if (!proposedSheet || !autoCatSheet) {
    throw new Error("Missing 'AutoCat' or 'AutoCat_Proposed' sheet.");
  }

  const proposedData = proposedSheet.getDataRange().getValues();
  const autoCatData = autoCatSheet.getDataRange().getValues();
  const proposedHeaders = proposedData[0];
  const autoCatHeaders = autoCatData[0];

  const keyCols = ['Category', 'Description Contains', 'Account Contains', 'Institution Contains'];
  const keyIndexesProposed = keyCols.map(col => proposedHeaders.indexOf(col));
  const keyIndexesAutoCat = keyCols.map(col => autoCatHeaders.indexOf(col));

  // Future-proof: find "Date Automatically Added" dynamically by header name
  const dateAutoHeaderName = "Date Automatically Added";
  const dateAutoColIndex = autoCatHeaders.indexOf(dateAutoHeaderName) + 1; // +1 because Sheets API is 1-based

  const hasAutoDateCol = dateAutoColIndex > 0;

  const existingKeys = new Set(autoCatData.slice(1).map(row =>
    keyIndexesAutoCat.map(i => (row[i] || "").toString().toLowerCase()).join("|")
  ));

  const rowsToAdd = [];
  const timestamp = new Date();

  for (let i = 1; i < proposedData.length; i++) {
    const row = proposedData[i];
    const key = keyIndexesProposed.map(j => (row[j] || "").toString().toLowerCase()).join("|");

    if (!existingKeys.has(key)) {
      const baseRow = row.slice(0, autoCatHeaders.length); // Always copy all existing columns

      if (hasAutoDateCol) {
        // If baseRow isn't long enough to reach the auto date column, pad it
        while (baseRow.length < dateAutoColIndex) {
          baseRow.push("");
        }
        baseRow[dateAutoColIndex - 1] = timestamp; // Set the date at the correct column
      }

      rowsToAdd.push(baseRow);
    }
  }

  if (rowsToAdd.length > 0) {
    const startRow = autoCatSheet.getLastRow() + 1;
    autoCatSheet.getRange(startRow, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);

    if (hasAutoDateCol) {
      autoCatSheet.getRange(startRow, dateAutoColIndex, rowsToAdd.length).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    }
  }

  return `✅ Merged ${rowsToAdd.length} new rule(s) into AutoCat.`;
}

