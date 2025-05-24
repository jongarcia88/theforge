// ─── CONFIGURATION ────────────────────────────────────────────────────────────
const EVENT_SHEET_NAME = 'Event Report';
const TX_START_COL     = 10;  // column J is 10

// Editable columns for your Event Report
const EVENT_EDITABLE = {
  [TX_START_COL + 1]: 'Description',     // K
  [TX_START_COL + 3]: 'JG Ownership %',  // M
  [TX_START_COL + 4]: 'RJ Ownership %',  // N
  [TX_START_COL + 5]: 'Category',        // O
  [TX_START_COL + 6]: 'Tags'             // P
};

// ─── onEdit ENTRYPOINT ────────────────────────────────────────────────────────
function onEdit(e) {
  const ss    = e.source;
  let   logSh = ss.getSheetByName('SyncLog');
  const ts    = new Date();

  // ─── ENSURE SyncLog SHEET & HEADER ───────────────────────────────────────────
  if (!logSh) {
    logSh = ss.insertSheet('SyncLog');
    logSh.appendRow([
      'Timestamp','Phase','Sheet','Cell','Row','Col',
      'Field','TransactionID','NewValue'
    ]);
  }

  const range  = e.range;
  const sheet  = range.getSheet();
  const name   = sheet.getName();
  const row    = range.getRow();
  const col    = range.getColumn();
  const a1     = range.getA1Notation();
  const newVal = range.getValue();

  // ─── UNIVERSAL LOG ENTRY ──────────────────────────────────────────────────────
  Logger.log(`onEdit → sheet="${name}", cell=${a1}, newValue="${newVal}"`);
  SpreadsheetApp.getActive().toast(`onEdit: ${name}!${a1}`, 'DEBUG', 1);
  logSh.appendRow([
    ts, 'ENTER', name, a1, row, col, '', '', newVal
  ]);

  // ─── EVENT REPORT LOGIC ──────────────────────────────────────────────────────
  if (name === EVENT_SHEET_NAME) {
    // a) I1 event‐tag change
    if (a1 === 'I1') {
      if (newVal !== lastSelectedEventTag) {
        lastSelectedEventTag = newVal;
        sheet.getRange('J1').setValue('⚡ Refresh Needed!');
      }
      logSh.appendRow([ts,'EventTag',name,a1,row,col,'Event‐tag','',newVal]);
      Logger.log(`EventTag changed to "${newVal}"`);
      return;
    }
    // b) ignore above row 5
    if (row < 5) return;
    // c) only these cols editable
    const field = EVENT_EDITABLE[col];
    if (!field) return;
    // d) hidden TransactionID in column R (J+8)
    const txId = sheet.getRange(row, TX_START_COL + 8).getValue();
    if (!txId) return;
    // e) split‐sync vs single‐field
    if (field === 'JG Ownership %' || field === 'RJ Ownership %') {
      let p = parseFloat(newVal);
      if (isNaN(p) || p < 0) p = 0;
      if (p > 100) p = 100;
      const isJG = field === 'JG Ownership %';
      const jg   = isJG ? p : (100 - p);
      const rj   = isJG ? (100 - p) : p;
      range.setValue(isJG ? jg : rj);
      sheet.getRange(row, TX_START_COL + (isJG ? 4 : 3))
           .setValue(isJG ? rj : jg);
      batchSaveFields(txId, {
        'JG Ownership %': jg,
        'RJ Ownership %': rj
      });
      logSh.appendRow([ts,'EventSplit',name,a1,row,col,field,txId,`${jg}/${rj}`]);
      Logger.log(`Split‐sync ${field} → JG=${jg}, RJ=${rj} for ${txId}`);
    } else {
      batchSaveFields(txId, { [field]: newVal });
      logSh.appendRow([ts,'EventSingle',name,a1,row,col,field,txId,newVal]);
      Logger.log(`Batch‐saved ${field}="${newVal}" for ${txId}`);
    }
    return;
  }

  // ─── JG → RY SYNC‐FLAGGING & ON‐DEMAND UUID ─────────────────────────────────
  if (name === JG_SHEET_NAME && row > 1) {
    // 1) Ensure single‐row TransactionID (col V → 22)
    const idCell = sheet.getRange(row, JG_COL.TRANSACTIONID);
    let txId = idCell.getValue();
    if (!txId) {
      txId = Utilities.getUuid();
      idCell.setValue(txId);
      logSh.appendRow([ts,'EnsureID',name,'',row,'','',txId,'']);
      Logger.log(`Row ${row}: generated TransactionID ${txId}`);
    }

    // 2) Mark LastModified (col AB → 28)
    sheet.getRange(row, JG_COL.LAST_MODIFIED).setValue(ts);
    // 3) Flag NeedsSyncOut (col AE → 31)
    sheet.getRange(row, JG_COL.NEEDS_SYNC_OUT).setValue(true);

    // 4) Log and toast
    logSh.appendRow([
      ts, 'FlagSync', name, a1, row, col,
      '', txId, newVal
    ]);
    SpreadsheetApp.getActive().toast(`✔ Marked row ${row}`, 'sync flag', 2);
    Logger.log(`Flagged row ${row} (TX=${txId}) for sync`);
  }
}
