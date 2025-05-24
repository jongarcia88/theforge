/**
 * Two-Way Sync Sidebar for JG ↔ RY with Console Debug Logging
 * Paste this entire content into your ImportExport.gs
 */

// === CONFIGURATION ===
const JG_SHEET_NAME = 'Transactions';
const RY_SHEET_ID   = '1AJGCgoWRsUKoxSV8BjoMA75gDViwz9fSekeJWH0yNG4';
const RY_SHEET_NAME = 'Sheet1';

// === COLUMN MAPS (1-based) ===
const JG_COL = {
  DATE:               2,
  DESCRIPTION:        3,
  CATEGORY:           4,
  TAGS:               5,
  AMOUNT:             6,
  ACCOUNT:            7,
  ACCOUNT_NUMBER:     8,
  INSTITUTION:        9,
  MONTH:              10,
  WEEK:               11,
  TRANSACTION_ID:     12,  // <- not used for sync
  ACCOUNT_ID:         13,
  CHECK_NUMBER:       14,
  FULL_DESCRIPTION:   15,
  DATE_ADDED:         16,
  CATEGORY_HINT:      17,
  CATEGORIZED_DATE:   18,
  TAG_LOG:            19,
  FIX_TIMESTAMP:      20,
  FIX_COMMENT:        21,  // <- Comments from RY now land here
  TRANSACTIONID:      22,  // <- this one we write to/from RY
  JG_OWNERSHIP:       23,
  RJ_OWNERSHIP:       24,
  KG_OWNERSHIP:       25,
  AMOUNT_REIMBURSED:  26,
  TOTAL_AMOUNT_PAID:  27,
  LAST_MODIFIED:      28,
  LAST_SYNCED_OUT:    29,
  LAST_SYNCED_IN:     30,
  NEEDS_SYNC_OUT:     31
};

const RY_COL = {
  DATE:               1,
  DESCRIPTION:        2,
  AMOUNT:             3,
  J_OWES_R:           4,
  R_OWES_J:           5,
  AMOUNT_REIMBURSED:  6,
  COMMENTS:           7,   // ← existing
  CATEGORY:           8,   // ← new
  TAGS:               9,   // ← new
  TRANSACTION_ID:    10,   // shifted
  LAST_MODIFIED:     11,
  LAST_SYNCED_OUT:   12,
  LAST_SYNCED_IN:    13,
  NEEDS_SYNC_IN:     14,
  ORIG_J_OWES_R:     15,
  ORIG_R_OWES_J:     16
};


// === CONSOLE LOGGER ===
function cLog(message) {
  console.log(message);
}

// === SIDEBAR UI ===
function showSyncSidebar() {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;padding:16px;">
      <h2>Sync Control</h2>
      <button onclick="markForSync()">📌 Mark for Sync</button><br/><br/>
      <button onclick="exportToRY()">➡️ Export to RY</button><br/><br/>
      <button onclick="importFromRY()">⬅️ Import from RY</button><br/><br/>
      <button onclick="showConflictSidebar()">⚠️ Review Conflicts</button>
      <h3>Logs:</h3>
      <pre id="log" style="height:200px;overflow:auto;background:#f9f9f9;padding:8px;border:1px solid #ddd;">No logs yet.</pre>
    </div>
    <script>
      function markForSync() {
        document.getElementById('log').innerText = 'Marking rows...';
        google.script.run.withSuccessHandler(logs => document.getElementById('log').innerText = logs.join('\\n')).markForSync();
      }
      function exportToRY() {
        document.getElementById('log').innerText = 'Starting export...';
        google.script.run.withSuccessHandler(logs => document.getElementById('log').innerText = logs.join('\\n')).exportToRY();
      }
      function importFromRY() {
        document.getElementById('log').innerText = 'Starting import...';
        google.script.run.withSuccessHandler(logs => document.getElementById('log').innerText = logs.join('\\n')).importFromRY();
      }
      function showConflictSidebar() {
        document.getElementById('log').innerText = 'Fetching conflicts...';
        google.script.run.withSuccessHandler(html => document.getElementById('log').innerHTML = html).showConflictSidebar();
      }
    </script>`  
  ).setTitle('Import/Export Sync');
  SpreadsheetApp.getUi().showSidebar(html);
}


// === MARK FOR SYNC ===
function markForSync() {
  const logs = [];
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const now = new Date();
  const isJG = sheet.getName() === JG_SHEET_NAME;
  const idCol = isJG
    ? JG_COL.TRANSACTIONID
    : RY_COL.TRANSACTION_ID;
  const modCol = isJG
    ? JG_COL.LAST_MODIFIED
    : RY_COL.LAST_MODIFIED;
  const flagCol = isJG
    ? JG_COL.NEEDS_SYNC_OUT
    : RY_COL.NEEDS_SYNC_IN;

  range.getValues().forEach((_, i) => {
    const row = range.getRow() + i;
    const idCell = sheet.getRange(row, idCol);
    if (!idCell.getValue()) {
      const uuid = Utilities.getUuid();
      idCell.setValue(uuid);
      logs.push(`Row ${row}: Generated UUID ${uuid}`);
      cLog(`markForSync: Row ${row} Generated ${uuid}`);
    }
    sheet.getRange(row, flagCol).setValue(true);
    sheet.getRange(row, modCol).setValue(now);
    logs.push(`Row ${row}: Marked for sync at ${now}`);
    cLog(`markForSync: Row ${row} Marked at ${now}`);
  });
  return logs;
}

// === EXPORT TO RY (UUID-on-demand), with Venmo special case ===
function exportToRY() {
  const logs   = [];
  const jgSh   = SpreadsheetApp.getActive().getSheetByName(JG_SHEET_NAME);
  const rySh   = SpreadsheetApp.openById(RY_SHEET_ID).getSheetByName(RY_SHEET_NAME);
  const jgData = jgSh.getDataRange().getValues();
  const ryData = rySh.getDataRange().getValues();
  const ryIds  = ryData.map(r => r[RY_COL.TRANSACTION_ID - 1]);
  const now    = new Date();

  // ── CONFIGURE YOUR EXPORT START DATE HERE ───────────────────────────────────
  const startDate = new Date('2025-03-01');  // ← adjust as needed

  let updatedCount = 0;
  let newCount     = 0;
  const newRows    = [];

  logs.push(`exportToRY start: JG rows=${jgData.length - 1}, RY rows=${ryData.length - 1}`);
  cLog(logs[logs.length - 1]);

  for (let i = 1; i < jgData.length; i++) {
    const rowIdx  = i + 1;
    const r       = jgData[i];

    // 0) skip if Account, Account # or Institution contains "RY Bank"
    const acct      = (r[JG_COL.ACCOUNT       - 1] || '').toString();
    const acctNum   = (r[JG_COL.ACCOUNT_NUMBER- 1] || '').toString();
    const inst      = (r[JG_COL.INSTITUTION   - 1] || '').toString();
    if ([acct, acctNum, inst].some(v => v.toUpperCase().includes('RY BANK'))) continue;

    // 1) skip if tagged ImportedRY
    const tagsCell = (r[JG_COL.TAGS - 1] || '').toString();
    if (tagsCell.toUpperCase().includes('IMPORTEDRY')) continue;

    // 2) skip if transaction date is before startDate
    const rawDate = r[JG_COL.DATE - 1];
    const txDate  = rawDate ? new Date(rawDate) : now;
    if (txDate < startDate) continue;

    // 3) only flag rows that have at least one tag starting with "RJ"
    const isRJTag = tagsCell
      .split(',')
      .some(t => t.trim().toUpperCase().startsWith('RJ'));
    const flag = isRJTag;

    // parse ownership %: blank → 50
    let pct = parseFloat(r[JG_COL.RJ_OWNERSHIP - 1]);
    if (isNaN(pct)) pct = 50;

    const amt     = parseFloat(r[JG_COL.AMOUNT           - 1]) || 0;
    const reimb   = parseFloat(r[JG_COL.AMOUNT_REIMBURSED - 1]) || 0;
    const modVal  = r[JG_COL.LAST_MODIFIED               - 1];
    const mod     = modVal ? new Date(modVal) : now;
    const lastOut = new Date(r[JG_COL.LAST_SYNCED_OUT    - 1] || 0);
    const lastIn  = new Date(r[JG_COL.LAST_SYNCED_IN     - 1] || 0);

    // only skip on flag and on LAST_SYNCED_IN
if (!flag || /*mod <= lastOut ||*/ mod <= lastIn) continue;


    // 5) generate TransactionID on-demand
    let txId = r[JG_COL.TRANSACTIONID - 1];
    if (!txId) {
      txId = Utilities.getUuid();
      jgSh.getRange(rowIdx, JG_COL.TRANSACTIONID).setValue(txId);
      logs.push(`Row ${rowIdx}: generated missing TransactionID ${txId}`);
      cLog(`exportToRY: Row ${rowIdx} generated ${txId}`);
    }

    // 6) build the RY payload (invert signs for RY)
    const max     = Math.max(...Object.values(RY_COL));
    const payload = Array(max).fill('');

    // common fields
    payload[RY_COL.DATE        - 1] = Utilities.formatDate(txDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    payload[RY_COL.DESCRIPTION - 1] = r[JG_COL.DESCRIPTION - 1];

    // special Venmo case: full amount → R_OWES_J
    const descUpper = (r[JG_COL.DESCRIPTION - 1] || '').toString().toUpperCase();
    if (descUpper.includes('VENMO')) {
      payload[RY_COL.AMOUNT       - 1] = -amt;
      payload[RY_COL.R_OWES_J     - 1] = -amt;
    } else {
      payload[RY_COL.AMOUNT       - 1] = -amt;
      payload[RY_COL.R_OWES_J     - 1] = -(amt * (pct / 100));
    }

    if (reimb !== 0) {
      payload[RY_COL.AMOUNT_REIMBURSED - 1] = -reimb;
    }

    payload[RY_COL.TRANSACTION_ID  - 1] = txId;
    payload[RY_COL.LAST_MODIFIED   - 1] = mod.toISOString();
    payload[RY_COL.NEEDS_SYNC_IN   - 1] = true;
    payload[RY_COL.COMMENTS        - 1] = r[JG_COL.FIX_COMMENT  - 1];
    payload[RY_COL.CATEGORY        - 1] = r[JG_COL.CATEGORY     - 1];
    payload[RY_COL.TAGS            - 1] = r[JG_COL.TAGS         - 1];

    // 7) update vs collect for append
    const idx = ryIds.indexOf(txId);
    if (idx > -1) {
      for (let c = 0; c < payload.length; c++) {
        if (payload[c] !== '') rySh.getRange(idx + 1, c + 1).setValue(payload[c]);
      }
      updatedCount++;
    } else {
      newRows.push({ payload, date: txDate });
      newCount++;
    }

    // 8) clear our flag & stamp outbound
    jgSh.getRange(rowIdx, JG_COL.NEEDS_SYNC_OUT ).setValue(false);
    jgSh.getRange(rowIdx, JG_COL.LAST_SYNCED_OUT).setValue(now);
  }

  // 9) append new rows in chronological order
  if (newRows.length) {
    newRows.sort((a, b) => a.date - b.date);
    newRows.forEach(n => rySh.appendRow(n.payload));
  }

  // 10) summary
  const total = updatedCount + newCount;
  logs.push(`Export complete: ${total} row(s).`);
  logs.push(`↳ Rows updated: ${updatedCount}`);
  logs.push(`↳ Rows appended: ${newCount}`);
  logs.forEach(e => cLog(e));
  return logs;
}


// === BACK-FILL MISSING TRANSACTIONIDs OVER THE WHOLE SHEET ===
function ensureTransactionIds() {
  const sh     = SpreadsheetApp.getActive().getSheetByName(JG_SHEET_NAME);
  const last   = sh.getLastRow();
  if (last < 2) return [];
  const range  = sh.getRange(2, JG_COL.TRANSACTIONID, last - 1, 1);
  const values = range.getValues();
  const now    = new Date();
  const logs   = [];

  values.forEach((row, i) => {
    if (!row[0]) {
      const id = Utilities.getUuid();
      sh.getRange(i + 2, JG_COL.TRANSACTIONID).setValue(id);
      logs.push(`Row ${i + 2}: filled missing TransactionID ${id}`);
    }
  });

  return logs;
}

// === IMPORT FROM RY (optimized, with comments & detailed logging) ===
function importFromRY() {
  const logs     = [];
  const now      = new Date();
  const jgSh     = SpreadsheetApp.getActive().getSheetByName(JG_SHEET_NAME);
  const rySh     = SpreadsheetApp.openById(RY_SHEET_ID).getSheetByName(RY_SHEET_NAME);
  const ryData   = rySh.getDataRange().getValues();
  const jgData   = jgSh.getDataRange().getValues();
  const jgIds    = jgData.map(r => r[JG_COL.TRANSACTIONID-1]);
  logs.push(`importFromRY start at ${now.toISOString()}`);

  let total        = 0;
  let skipNoChange = 0;
  let skipWrongDir = 0;
  let updatedCount = 0;
  let newCount     = 0;
  const appendRows = [];
  const updatedRows = [];  // { idx,rowData }
  const ryFlagUpdates = []; // [rowIdx, needsSyncIn, lastSyncedIn]

// for loop starts at row 4 (skipping the triple header)
  for (let i = 3; i < ryData.length; i++) {
    total++;
    const rowIdx = i + 1;
    const r      = ryData[i];

    // IDs & timestamps
    let id     = r[RY_COL.TRANSACTION_ID-1] || Utilities.getUuid();
    let modVal = r[RY_COL.LAST_MODIFIED-1];
    let lastIn = new Date(r[RY_COL.LAST_SYNCED_IN-1] || 0);
    let lastOut= new Date(r[RY_COL.LAST_SYNCED_OUT-1]|| 0);
    let mod    = modVal ? new Date(modVal) : now;

    // ensure RY has id & mod filled
    if (!r[RY_COL.TRANSACTION_ID-1]) rySh.getRange(rowIdx, RY_COL.TRANSACTION_ID).setValue(id);
    if (!modVal)                          rySh.getRange(rowIdx, RY_COL.LAST_MODIFIED).setValue(mod);

    // skip if nothing new
    if (!(mod > lastIn && mod > lastOut)) {
      skipNoChange++;
      continue;
    }

    // ownership filter
    const rawAmt = r[RY_COL.AMOUNT-1];
    const ORIG_jO = parseFloat(r[RY_COL.ORIG_J_OWES_R-1]) || 0;
    const ORIG_rO = parseFloat(r[RY_COL.ORIG_R_OWES_J-1]) || 0;
    const jO     = parseFloat(r[RY_COL.J_OWES_R-1]) || 0;
    const rO     = parseFloat(r[RY_COL.R_OWES_J-1]) || 0;
    if (!(jO > 0 && rO === 0)) {
      skipWrongDir++;
      continue;
    }

    // compute JG payload
    const reimb = parseFloat(r[RY_COL.AMOUNT_REIMBURSED-1]) || 0;

    // --- UPDATED FIX START ---
    const useOrig = reimb !== 0;
    const jSlice  = useOrig ? ORIG_jO : jO;
    const rSlice  = useOrig ? ORIG_rO : rO;
    const totalAmt = rawAmt !== ''
      ? rawAmt
      : (jSlice + rSlice) * 2;
    const pctJG   = totalAmt ? (jSlice * 100) / totalAmt : 0;
    const pctRJ   = 100 - pctJG;
    // --- UPDATED FIX END ---

    const signed   = -jO;
    const desc     = r[RY_COL.DESCRIPTION-1] || '';
    const dateVal  = r[RY_COL.DATE-1];
    const comment  = r[RY_COL.COMMENTS-1] || '';

    // build row
    const max = Math.max(...Object.values(JG_COL));
    const pl  = Array(max).fill('');
    // tags
    const existingTags = jgIds.indexOf(id) > -1
      ? jgSh.getRange(jgIds.indexOf(id)+1, JG_COL.TAGS).getValue()
      : '';
    pl[JG_COL.TAGS-1]            = existingTags.includes('ImportedRY') 
      ? existingTags 
      : (existingTags? existingTags+',ImportedRY' : 'ImportedRY');
    pl[JG_COL.DATE-1]            = dateVal;
    pl[JG_COL.DESCRIPTION-1]     = desc;
    pl[JG_COL.FULL_DESCRIPTION-1]= desc;
    pl[JG_COL.AMOUNT-1]          = signed;
    pl[JG_COL.TOTAL_AMOUNT_PAID-1]= -totalAmt;
    pl[JG_COL.ACCOUNT-1]         = 'RY Bank';
    pl[JG_COL.ACCOUNT_NUMBER-1]  = 'RY Bank';
    pl[JG_COL.INSTITUTION-1]     = 'RY Bank';
    pl[JG_COL.DATE_ADDED-1]      = now;
    pl[JG_COL.TRANSACTIONID-1]   = id;
    pl[JG_COL.RJ_OWNERSHIP-1]    = pctRJ;
    pl[JG_COL.JG_OWNERSHIP-1]    = pctJG;
    pl[JG_COL.AMOUNT_REIMBURSED-1]= reimb;
    pl[JG_COL.COMMENTS-1]        = comment;
    pl[JG_COL.LAST_MODIFIED-1]   = mod;
    pl[JG_COL.NEEDS_SYNC_OUT-1]  = true;

    // --- NEW: Month & Week ---
    var { month, week } = getMonthAndWeekDates(dateVal);
    pl[JG_COL.MONTH-1] = month;
    pl[JG_COL.WEEK-1]  = week;

    // update vs append
    const existIdx = jgIds.indexOf(id);
    if (existIdx > -1) {
      updatedRows.push({ idx: existIdx+1, data: pl });
      updatedCount++;
    } else {
      appendRows.push(pl);
      newCount++;
    }

    // mark RY as synced
    ryFlagUpdates.push([rowIdx, false, now]);
  }

  // batch-append new
  if (appendRows.length) {
    const start = jgSh.getLastRow() + 1;
    jgSh.getRange(start, 1, appendRows.length, appendRows[0].length)
        .setValues(appendRows);
  }

  // apply updates
  updatedRows.forEach(u => {
    jgSh.getRange(u.idx, 1, 1, u.data.length).setValues([u.data]);
  });

  // batch-update RY sync flags & timestamps
  ryFlagUpdates.forEach(u => {
    const [r, flag, ts] = u;
    rySh.getRange(r, RY_COL.NEEDS_SYNC_IN).setValue(flag);
    rySh.getRange(r, RY_COL.LAST_SYNCED_IN).setValue(ts);
  });

  // summary
  logs.push(`Total RY rows scanned: ${total}`);
  logs.push(`↳ Skipped (no new changes): ${skipNoChange}`);
  logs.push(`↳ Skipped (wrong direction): ${skipWrongDir}`);
  logs.push(`Rows UPDATED in JG: ${updatedCount}`);
  logs.push(`Rows APPENDED to JG: ${newCount}`);
  logs.forEach(msg => console.info(msg));
  return logs;
}

// === CONFLICT REVIEW ===
function getConflictIds() {
  const sh = SpreadsheetApp.getActive().getSheetByName(JG_SHEET_NAME);
  return sh.getDataRange().getValues().slice(1)
    .filter(r =>
      r[JG_COL.NEEDS_SYNC_OUT-1] &&
      new Date(r[JG_COL.LAST_MODIFIED-1]) >
      new Date(r[JG_COL.LAST_SYNCED_IN-1])
    )
    .map(r => r[JG_COL.TRANSACTIONID-1]);
}

function showConflictSidebar() {
  const ids = getConflictIds();
  let html = '<div style="padding:16px;font-family:Arial,sans-serif;"><h3>Conflicts</h3>';
  html += ids.length
    ? '<ul>' + ids.map(i=>`<li>${i}</li>`).join('') + '</ul>'
    : '<p>None 🎉</p>';
  html += '</div>';
  return html;
}

// === APPLY REIMBURSEMENT ===
function applyReimbursement(sh, row) {
  // — A) stash originals on first reimbursement —
  if (!sh.getRange(row, COL.ORIG_J).getValue() && !sh.getRange(row, COL.ORIG_R).getValue()) {
    storeOriginal(sh, row);
  }

  // — B) pull reimbursement amount —
  const M = parseFloat(sh.getRange(row, COL.AMOUNT_REIMBURSED).getValue()) || 0;
  if (!M) return;  // nothing to do if reimbursement cleared or zero

  // — C) pull original helper values (may be positive or negative) —
  const origJ = parseFloat(sh.getRange(row, COL.ORIG_J).getValue()) || 0;
  const origR = parseFloat(sh.getRange(row, COL.ORIG_R).getValue()) || 0;

  // work with absolute balances
  const absJ = Math.abs(origJ);
  const absR = Math.abs(origR);
  const total = absJ + absR;
  if (total === 0) return;

  // — D) determine each share on that original total —
  const pctJ = absJ / total;
  const pctR = absR / total;

  // — E) apply reimbursement against those absolutes, then re-sign —
  if (absJ > 0) {
    // original was J owes R → we add R’s share back to J_OWES_R
    const share = M * pctR;
    const newJ  = absJ + share;
    sh.getRange(row, COL.J_OWES_R).setValue(newJ.toFixed(2));
  } else {
    // original was R owes J → we subtract J’s share from R_OWES_J
    const share = M * pctJ;
    const newR  = absR - share;
    // R_OWES_J is stored as negative
    sh.getRange(row, COL.R_OWES_J).setValue((-newR).toFixed(2));
  }
}

/**
 * Given a JS Date, returns the first-of-month date
 * and the week-start (Sunday) date just like:
 *
 * Month: =DATE(YEAR(B), MONTH(B), 1)
 * Week:  =B - WEEKDAY(B) + 1  (so that Sundays map to themselves)
 *
 * @param {Date} date  A JS Date object (transaction date)
 * @return {{ month: Date, week: Date }}
 */
function getMonthAndWeekDates(date) {
  if (!(date instanceof Date)) {
    date = new Date(date);  // allow strings, etc.
  }
  
  // --- Month: first day of that month ---
  const year  = date.getFullYear();
  const month = date.getMonth();       // 0-indexed
  const monthDate = new Date(year, month, 1);

  // --- Week: back up to the most recent Sunday ---
  // JS getDay(): Sunday=0, Monday=1, … Saturday=6
  const dayOfWeek = date.getDay();
  // subtract dayOfWeek days (in ms) from the date
  const msPerDay  = 24 * 60 * 60 * 1000;
  const weekDate  = new Date(date.getTime() - dayOfWeek * msPerDay);

  return {
    month: monthDate,
    week:  weekDate
  };
}

// === EXPORT TO RY (UUID-on-demand), with Venmo special case & ImportedRY handling ===
function exportToRY2() {
  const logs   = [];
  const jgSh   = SpreadsheetApp.getActive().getSheetByName(JG_SHEET_NAME);
  const rySh   = SpreadsheetApp.openById(RY_SHEET_ID).getSheetByName(RY_SHEET_NAME);
  const jgData = jgSh.getDataRange().getValues();
  const ryData = rySh.getDataRange().getValues();
  const ryIds  = ryData.map(r => r[RY_COL.TRANSACTION_ID - 1]);
  const now    = new Date();

  const startDate = new Date('2025-03-01');  // ← adjust as needed
  let updatedCount = 0, newCount = 0;
  const newRows = [];

  logs.push(`exportToRY start: JG rows=${jgData.length - 1}, RY rows=${ryData.length - 1}`);
  cLog(logs[logs.length - 1]);

  for (let i = 1; i < jgData.length; i++) {
    const rowIdx   = i + 1;
    const r        = jgData[i];
    const tagsCell = (r[JG_COL.TAGS - 1] || '').toString();
    const tagsArr  = tagsCell.split(',').map(t => t.trim().toUpperCase());
    const isRJTag      = tagsArr.some(t => t.startsWith('RJ'));
    const isImportedRY = tagsArr.includes('IMPORTEDRY') || tagsArr.includes('IMPORT RY');

    // include only RJ- or ImportedRY-tagged
    if (!isRJTag && !isImportedRY) continue;

    const rawDate = r[JG_COL.DATE - 1];
    const txDate  = rawDate ? new Date(rawDate) : now;
    if (txDate < startDate) continue;

    const modVal = r[JG_COL.LAST_MODIFIED - 1];
    const mod    = modVal ? new Date(modVal) : now;
    const lastIn = new Date(r[JG_COL.LAST_SYNCED_IN - 1] || 0);
    if (mod <= lastIn) continue;

    // ensure TransactionID
    let txId = r[JG_COL.TRANSACTIONID - 1];
    if (!txId) {
      txId = Utilities.getUuid();
      jgSh.getRange(rowIdx, JG_COL.TRANSACTIONID).setValue(txId);
      logs.push(`Row ${rowIdx}: generated missing TransactionID ${txId}`);
      cLog(`exportToRY: Row ${rowIdx} generated ${txId}`);
    }

    // SPECIAL CASE: ImportedRY-only payload (takes precedence over RJ tags)
    if (isImportedRY) {
      const max     = Math.max(...Object.values(RY_COL));
      const payload = Array(max).fill('');
      payload[RY_COL.TRANSACTION_ID  - 1] = txId;
      payload[RY_COL.CATEGORY        - 1] = r[JG_COL.CATEGORY     - 1];
      payload[RY_COL.TAGS            - 1] = r[JG_COL.TAGS         - 1];
      payload[RY_COL.LAST_MODIFIED   - 1] = mod.toISOString();
      payload[RY_COL.NEEDS_SYNC_IN   - 1] = true;

      const idx = ryIds.indexOf(txId);
      if (idx > -1) {
        rySh.getRange(idx + 1, RY_COL.CATEGORY     ).setValue(payload[RY_COL.CATEGORY     - 1]);
        rySh.getRange(idx + 1, RY_COL.TAGS         ).setValue(payload[RY_COL.TAGS         - 1]);
        rySh.getRange(idx + 1, RY_COL.LAST_MODIFIED).setValue(payload[RY_COL.LAST_MODIFIED - 1]);
        rySh.getRange(idx + 1, RY_COL.NEEDS_SYNC_IN).setValue(true);
        updatedCount++;
      } else {
        newRows.push({ payload, date: txDate });
        newCount++;
      }

      jgSh.getRange(rowIdx, JG_COL.NEEDS_SYNC_OUT ).setValue(false);
      jgSh.getRange(rowIdx, JG_COL.LAST_SYNCED_OUT).setValue(now);
      continue;
    }

    // RJ-tagged rows continue with existing logic:
    let pct   = parseFloat(r[JG_COL.RJ_OWNERSHIP - 1]);
    if (isNaN(pct)) pct = 50;
    const amt    = parseFloat(r[JG_COL.AMOUNT           - 1]) || 0;
    const reimb  = parseFloat(r[JG_COL.AMOUNT_REIMBURSED - 1]) || 0;
    const max    = Math.max(...Object.values(RY_COL));
    const payload = Array(max).fill('');

    payload[RY_COL.DATE        - 1] = Utilities.formatDate(txDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    payload[RY_COL.DESCRIPTION - 1] = r[JG_COL.DESCRIPTION - 1];
    const descUpper = (r[JG_COL.DESCRIPTION - 1] || '').toUpperCase();
    if (descUpper.includes('VENMO')) {
      payload[RY_COL.AMOUNT   - 1] = -amt;
      payload[RY_COL.R_OWES_J - 1] = -amt;
    } else {
      payload[RY_COL.AMOUNT   - 1] = -amt;
      payload[RY_COL.R_OWES_J - 1] = -(amt * (pct / 100));
    }
    if (reimb !== 0) payload[RY_COL.AMOUNT_REIMBURSED - 1] = -reimb;

    payload[RY_COL.TRANSACTION_ID  - 1] = txId;
    payload[RY_COL.LAST_MODIFIED   - 1] = mod.toISOString();
    payload[RY_COL.NEEDS_SYNC_IN   - 1] = true;
    payload[RY_COL.COMMENTS        - 1] = r[JG_COL.FIX_COMMENT  - 1];
    payload[RY_COL.CATEGORY        - 1] = r[JG_COL.CATEGORY     - 1];
    payload[RY_COL.TAGS            - 1] = r[JG_COL.TAGS         - 1];

    const idx2 = ryIds.indexOf(txId);
    if (idx2 > -1) {
      for (let c = 0; c < payload.length; c++) {
        if (payload[c] !== '') rySh.getRange(idx2 + 1, c + 1).setValue(payload[c]);
      }
      updatedCount++;
    } else {
      newRows.push({ payload, date: txDate });
      newCount++;
    }

    jgSh.getRange(rowIdx, JG_COL.NEEDS_SYNC_OUT ).setValue(false);
    jgSh.getRange(rowIdx, JG_COL.LAST_SYNCED_OUT).setValue(now);
  }

  if (newRows.length) {
    newRows.sort((a, b) => a.date - b.date);
    newRows.forEach(n => rySh.appendRow(n.payload));
  }

  const total = updatedCount + newCount;
  logs.push(`Export complete: ${total} row(s).`);
  logs.push(`↳ Rows updated: ${updatedCount}`);
  logs.push(`↳ Rows appended: ${newCount}`);
  logs.forEach(e => cLog(e));
  return logs;
}
