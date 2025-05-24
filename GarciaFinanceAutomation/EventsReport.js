let lastSelectedEventTag = "";
let isSyncingSplit = false;

function buildEventReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = ss.getSheetByName("EventTags");
  const reportSheetName = "Event Report";

  let sheet = ss.getSheetByName(reportSheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(reportSheetName);

  // Clear any leftover helper data
  sheet.getRange("AZ1:AZ").clearContent();

  // Build dropdown source from EventTags!E
  const rawTags = eventsSheet.getRange("E2:E" + eventsSheet.getLastRow()).getValues().flat();
  const eventTags = rawTags.filter(t => t && t.toString().trim());
  if (!eventTags.length) {
    SpreadsheetApp.getUi().alert("⚠️ No event tags found in EventTags! Please check column E.");
    return;
  }
  const ddRange = sheet.getRange(1, 52, eventTags.length, 1);
  ddRange.setValues(eventTags.map(t => [t])).setFontColor("#ffffff");
  const ddRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ddRange, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("I1").setDataValidation(ddRule)
    .setNote("Select an event tag to view its report");

  // Default to last tag
  const defaultTag = eventTags[eventTags.length - 1];
  sheet.getRange("I1").setValue(defaultTag);
  lastSelectedEventTag = defaultTag;

  // Build report & add refresh button
  refreshEventReport();
  insertRefreshButton();
}

function updateEventDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Event Report");
  const eventsSheet = ss.getSheetByName("EventTags");
  const rawTags = eventsSheet.getRange("E2:E" + eventsSheet.getLastRow()).getValues().flat();
  const eventTags = rawTags.filter(t => t && t.toString().trim());
  const ddRange = sheet.getRange(1, 52, eventTags.length, 1);
  ddRange.setValues(eventTags.map(t => [t])).setFontColor("#ffffff");
  const ddRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ddRange, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("I1").setDataValidation(ddRule);
}

function insertRefreshButton() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event Report");
  if (!sheet) return;
  
  // If you re‐run buildEventReportSheet, clear out any old drawings:
  sheet.getDrawings().forEach(d => d.remove());
  
  // Put a little label so people know what to click:
  sheet.getRange("J1").setValue("🔄 Refresh Report");
  
  // Build the button as an embedded drawing:
  const builder = SpreadsheetApp.newDrawing()
    .setOnAction("manualRefreshEventReport")  // which function to call
    .setText("🔄 Refresh")                    // button text
    .setPosition(1, 10, 0, 0)                 // row 1, col 10 (J1), no offset
    .setWidth(100)
    .setHeight(30);
  
  const drawing = builder.build();
  sheet.insertDrawing(drawing);
}


function onSelectionChange(e) {
  if (e.range.getSheet().getName() !== "Event Report") return;
  if (e.range.getA1Notation() === "I1") {
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Tag changed—click Refresh");
  }
}




function batchSaveFields(txId, fields) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheetByName("Transactions");
    const hdr = txSheet.getRange(1,1,1,txSheet.getLastColumn()).getValues()[0].map(h=>h.toString().trim());
    const idCol = hdr.indexOf("TransactionID");
    if (idCol<0) return;

    // find row
    const lastRow = txSheet.getLastRow();
    const ids = txSheet.getRange(2, idCol+1, lastRow-1,1).getValues().flat();
    const row = ids.indexOf(txId);
    if (row<0) return;
    const targetRow = row+2;

    // write each field
    Object.entries(fields).forEach(([f,v])=>{
      let col = hdr.indexOf(f);
      if (col<0) {
        col = hdr.length;
        txSheet.getRange(1, col+1).setValue(f);
        hdr.push(f);
      }
      txSheet.getRange(targetRow, col+1).setValue(v);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function clearEventReport() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event Report");
  if (!s) return;
  s.getRange("A1:R1000").clearContent().clearFormat().clearDataValidations();
  s.getCharts().forEach(c=>s.removeChart(c));
}

function refreshEventReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Event Report");
  const txSheet = ss.getSheetByName("Transactions");
  const catSheet = ss.getSheetByName("Categories");
  if (!sheet || !txSheet || !catSheet) return;

  // 1) Get selected tag
  const eventTag = sheet.getRange("I1").getValue();
  if (!eventTag) return;

  // 2) Clear and rebuild dropdown
  clearEventReport();
  updateEventDropdown();
  sheet.getRange("I1").setValue(eventTag);

  // 3) Title
  sheet.getRange("A1:D1").merge()
    .setValue("📊 Event Report")
    .setFontSize(16).setFontWeight("bold").setFontColor("#1a237e")
    .setHorizontalAlignment("center");

  // 4) Fetch & filter transactions
  const data = txSheet.getDataRange().getValues();
  const hdr  = data.shift();
  const matches = data.map(r => {
    const tags = (r[hdr.indexOf("Tags")] || "").toString().split(",").map(t=>t.trim());
    if (!tags.includes(eventTag)) return null;
    // pull fields
    const id    = r[hdr.indexOf("TransactionID")];
    const date  = r[hdr.indexOf("Date")];
    const desc  = r[hdr.indexOf("Description")];
    const amt   = parseFloat(r[hdr.indexOf("Amount")]) || 0;
    const cat   = r[hdr.indexOf("Category")] || "Uncategorized";
    const acct  = r[hdr.indexOf("Account")] || "";
    // splits with defaults to 50/50
    const jgRaw = parseFloat(r[hdr.indexOf("JG Ownership %")]);
    const rjRaw = parseFloat(r[hdr.indexOf("RJ Ownership %")]);
    let   jgPct = isNaN(jgRaw) ? null : jgRaw;
    let   rjPct = isNaN(rjRaw) ? null : rjRaw;
    if (jgPct === null && rjPct === null) { jgPct = 50; rjPct = 50; }
    else if (jgPct === null)                { jgPct = 100 - rjPct; }
    else if (rjPct === null)                { rjPct = 100 - jgPct; }
    return { id, date, desc, amt, cat, tags, acct, jgPct, rjPct };
  }).filter(x=>x);

  // 5) Total cost
  const totalCost = matches.reduce((s,tx)=>s+tx.amt,0);
  sheet.getRange("A3")
    .setValue(`💸 Total cost for "${eventTag}": $${totalCost.toFixed(2)}`);

  // 6) Summary by Category
  const summary = {};
  matches.forEach(tx => summary[tx.cat] = (summary[tx.cat]||0) + tx.amt);
  const summaryArr = Object.entries(summary)
    .filter(([_,amt])=>amt!==0)
    .map(([cat,amt])=>[cat, Math.round(amt*100)/100])
    .sort((a,b)=>Math.abs(b[1])-Math.abs(a[1]));

  if (summaryArr.length) {
    const startRow = 5;
    // header
    sheet.getRange(startRow,1,1,2)
      .setValues([["Category","Total"]])
      .setFontWeight("bold").setBackground("#dbeeff");
    // data
    sheet.getRange(startRow+1,1,summaryArr.length,2)
      .setValues(summaryArr);

    // hidden abs data for pie
    const absArr = summaryArr.map(([c,v])=>[c,Math.abs(v)]);
    const absRange = sheet.getRange(startRow+1,27,absArr.length,2); // AA:AB
    absRange.setValues(absArr).setFontColor("#ffffff");

    // pie @ C5
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(absRange)
      .setOption("title","Spending by Category")
      .setPosition(startRow,3,0,0)  // row 5, col C
      .build();
    sheet.insertChart(chart);
  }

  // 7) Detailed Transactions @ J5
  const txCol = 10; // J
  let txRow = 5;
  // group by category
  const groups = {};
  matches.forEach(tx => {
    if (!groups[tx.cat]) groups[tx.cat] = { total:0, items:[] };
    groups[tx.cat].items.push(tx);
    groups[tx.cat].total += tx.amt;
  });
  const sorted = Object.entries(groups)
    .map(([cat,g])=>({cat,total:g.total,items:g.items}))
    .sort((a,b)=>Math.abs(b.total)-Math.abs(a.total));

  sorted.forEach(group => {
    // group header
    sheet.getRange(txRow, txCol)
      .setValue(`📂 ${group.cat} – $${Math.abs(group.total).toFixed(2)}`)
      .setFontWeight("bold").setFontColor("#1a237e");
    txRow++;
    // columns title
    sheet.getRange(txRow, txCol,1,9)
      .setValues([[ 
        "Date","Description","Amount","JG owns %","RJ owns %",
        "Category","Tags","Account","TransactionID"
      ]])
      .setFontWeight("bold").setBackground("#eeeeee");
    txRow++;
    // data rows
    const rows = group.items.map(tx=>[
      tx.date,
      tx.desc,
      tx.amt,
      tx.jgPct,
      tx.rjPct,
      tx.cat,
      tx.tags,
      tx.acct,
      tx.id
    ]);
    sheet.getRange(txRow,txCol,rows.length,9).setValues(rows);
    // category dropdown
    const catRange = catSheet.getRange("A2:A"+catSheet.getLastRow());
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(catRange,true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(txRow,txCol+5,rows.length,1).setDataValidation(rule);
    txRow += rows.length + 1;
  });

  // 8) Hide/show split cols so that JG+RJ always visible when appropriate
  const showSplit = eventTag.startsWith("RJ");
  sheet.setColumnWidth(txCol+2, showSplit ? 100 : 2); // L = JG owns %
  sheet.setColumnWidth(txCol+3, showSplit ? 100 : 2); // M = RJ owns %
}


function manualRefreshEventReport() {
  SpreadsheetApp.getActiveSpreadsheet().toast("⏳ Refreshing...");
  refreshEventReport();
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Refreshed!");
}
