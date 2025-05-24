// === CONFIGURATION ===
// Sheet names and constants used throughout the script
const SHEET_NAME = 'Sheet1';
const MAPPING_SHEET_NAME = 'EventTagColors';
const HEADER_ROWS = 3;  // Skip first three rows (header + extra)

/**
 * Reads the first row of headers and returns a map of headerName -> columnIndex.
 * @param {Sheet} sh - The sheet to read headers from.
 * @returns {Object} headerMap
 */
function getHeaderMap(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h != null && h !== '') {
      map[h.toString().trim()] = i + 1;
    }
  });
  return map;
}

// === MAIN EDIT TRIGGER ===
/**
 * Runs on every user edit. Handles value mirroring, auto-half splits, reimbursement logic,
 * and marks rows as dirty with timestamp and transaction ID.
 * @param {Event} e - The onEdit event object.
 */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;

  const headerMap    = getHeaderMap(sh);
  const colJ         = headerMap['J Owes R'];
  const colR         = headerMap['R Owes J'];
  const colAmt       = headerMap['Amount'];
  const colReimb     = headerMap['Amount Reimbursed'];
  const colOrigJ     = headerMap['Orig J OWES R'] || headerMap['Orig J'];
  const colOrigR     = headerMap['Orig R OWES J'] || headerMap['Orig R'];
  const colID        = headerMap['TransactionID'];
  const colLastMod   = headerMap['LastModified'];
  const colNeedsSync = headerMap['NeedsSyncIn'];

  const startRow = e.range.getRow();
  const startCol = e.range.getColumn();
  const numRows  = e.range.getNumRows();
  const numCols  = e.range.getNumColumns();
  const values   = e.range.getValues();

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row <= HEADER_ROWS) continue;

    let rowChanged = false;

    // A) Mirror or clear helper columns when J or R is edited
    for (let j = 0; j < numCols; j++) {
      const col    = startCol + j;
      const newVal = values[i][j];

      if (col === colJ) {
        if (!newVal) sh.getRange(row, colOrigJ).clearContent();
        else        sh.getRange(row, colOrigJ).setValue(newVal);
        rowChanged = true;
      }
      if (col === colR) {
        if (!newVal) sh.getRange(row, colOrigR).clearContent();
        else        sh.getRange(row, colOrigR).setValue(newVal);
        rowChanged = true;
      }

      // B) Auto-split amount into J and R if fresh Amount entered
      if (col === colAmt) {
        const amt   = parseFloat(newVal);
        const jCell = sh.getRange(row, colJ);
        const rCell = sh.getRange(row, colR);
        if (newVal && !isNaN(amt) && !jCell.getValue() && !rCell.getValue()) {
          jCell.setValue((amt / 2).toFixed(2));
          rowChanged = true;
        }
      }

      // D) Handle reimbursements: reset or reapply splits
      if (col === colReimb) {
        if (!newVal) {
          resetToOriginal(sh, row, colOrigJ, colOrigR, colJ, colR);
        } else {
          applyReimbursement(sh, row, colAmt, colReimb, colOrigJ, colOrigR, colJ, colR);
        }
        rowChanged = true;
      }
    }

    // E) Mark any other user edit as dirty (excluding internal columns)
    if (!rowChanged) {
      const editedCols = Array.from({ length: numCols }, (_, k) => startCol + k);
      const internalCols = [colID, colLastMod, colNeedsSync];
      if (!editedCols.some(c => internalCols.includes(c))) {
        rowChanged = true;
      }
    }

    // C) If rowChanged, ensure ID, flag sync and timestamp
    if (rowChanged) {
      ensureId(sh, row, colID);
      sh.getRange(row, colNeedsSync).setValue(true);
      sh.getRange(row, colLastMod).setValue(new Date());
    }
  }
}

// === HELPER: Ensure a unique Transaction ID ===
function ensureId(sh, row, colID) {
  const idCell = sh.getRange(row, colID);
  if (!idCell.getValue()) idCell.setValue(Utilities.getUuid());
}

// === HELPER: Reset J/R values back to their originals ===
function resetToOriginal(sh, row, colOrigJ, colOrigR, colJ, colR) {
  const origJ = sh.getRange(row, colOrigJ).getValue();
  const origR = sh.getRange(row, colOrigR).getValue();
  if (origJ) sh.getRange(row, colJ).setValue(origJ);
  if (origR) sh.getRange(row, colR).setValue(origR);
  sh.getRange(row, colOrigJ).clearContent();
  sh.getRange(row, colOrigR).clearContent();
}

// === HELPER: Apply reimbursement by adjusting J/R splits proportionally ===
function applyReimbursement(sh, row, colAmt, colReimb, colOrigJ, colOrigR, colJ, colR) {
  if (!sh.getRange(row, colOrigJ).getValue() && !sh.getRange(row, colOrigR).getValue()) {
    const jVal = sh.getRange(row, colJ).getValue();
    const rVal = sh.getRange(row, colR).getValue();
    if (jVal) sh.getRange(row, colOrigJ).setValue(jVal);
    if (rVal) sh.getRange(row, colOrigR).setValue(rVal);
  }

  let A = parseFloat(sh.getRange(row, colAmt).getValue()) || 0;
  const M = parseFloat(sh.getRange(row, colReimb).getValue()) || 0;
  if (!M) return;

  const origJ = parseFloat(sh.getRange(row, colOrigJ).getValue()) || 0;
  const origR = parseFloat(sh.getRange(row, colOrigR).getValue()) || 0;

  if (A <= 0 && (origJ || origR)) {
    A = (origJ || origR) * 2;
    sh.getRange(row, colAmt).setValue(A.toFixed(2));
  }
  if (A <= 0) return;

  let jPct = origJ ? (origJ / A) * 100 : 50;
  let rPct = 100 - jPct;
  if (origR && !origJ) {
    rPct = (origR / A) * 100;
    jPct = 100 - rPct;
  }

  if (origJ) {
    const share = M * (jPct / 100);
    sh.getRange(row, colJ).setValue((origJ + share).toFixed(2));
  } else if (origR) {
    const share = M * (rPct / 100);
    sh.getRange(row, colR).setValue((origR - share).toFixed(2));
  }
}

/**
 * Batch color rows by RJ-prefixed event tag, converting any HSL colors
 * in the mapping sheet to hex and then using hex for backgrounds.
 */
function applyEventTagColors() {
  Logger.log('applyEventTagColors: start');
  const ss = SpreadsheetApp.getActive();
  Logger.log('Spreadsheet ID: ' + ss.getId());

  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + SHEET_NAME + '" not found');
  Logger.log('Data sheet: ' + sh.getName());

  const headerMap = getHeaderMap(sh);
  Logger.log('Header map: ' + JSON.stringify(headerMap));

  const colTag = headerMap['Tags'];
  Logger.log('Tags column index: ' + colTag);
  if (!colTag) throw new Error('Could not find "Tags" column in header row');

  const mapSheet = getOrCreateMapSheet(ss);
  Logger.log('Mapping sheet: ' + mapSheet.getName());

  // Read and sanitize existing map entries (HSL → hex)
  const rawMap = mapSheet.getDataRange().getValues();
  const colorMap = {};
  for (let i = 1; i < rawMap.length; i++) {
    const tag    = rawMap[i][0];
    let   colStr = rawMap[i][1];
    if (!tag || !colStr) continue;

    let hex;
    if (/^hsl/i.test(colStr)) {
      hex = hslToHex(colStr);
      mapSheet.getRange(i + 1, 2).setValue(hex);
      Logger.log(`Converted ${colStr} → ${hex}`);
    } else {
      hex = colStr;
    }
    colorMap[tag] = hex;
  }
  Logger.log('Color map keys: ' + Object.keys(colorMap).join(', '));

  const lastRow = sh.getLastRow();
  Logger.log('Last data row: ' + lastRow);
  if (lastRow <= HEADER_ROWS) {
    Logger.log('No data rows to process, exiting');
    return;
  }
  const lastCol     = sh.getLastColumn();
  const dataStart   = HEADER_ROWS + 1;
  const numDataRows = lastRow - HEADER_ROWS;

  Logger.log(`Processing rows ${dataStart}–${lastRow}, columns 1–${lastCol}`);
  const tagsA = sh.getRange(dataStart, colTag, numDataRows, 1).getValues();

  // Build a 2D array of background colors
  const bgColors = tagsA.map((row, i) => {
    const cellTag = row[0];
    const tag     = extractEventTag(cellTag);
    Logger.log(`Row ${dataStart + i}: cell tag="${cellTag}", extracted="${tag}"`);
    if (tag) {
      if (!colorMap[tag]) {
        const newColor = getRandomSoftHex();
        colorMap[tag] = newColor;
        mapSheet.appendRow([tag, newColor]);
        Logger.log(`Added new tag "${tag}" with color ${newColor}`);
      }
      return Array(lastCol).fill(colorMap[tag]);
    }
    return Array(lastCol).fill(null);
  });

  sh.getRange(dataStart, 1, bgColors.length, lastCol).setBackgrounds(bgColors);
  Logger.log('applyEventTagColors: done');
}
/**
 * Get or create the tag-color mapping sheet.
 */
function getOrCreateMapSheet(ss) {
  let sheet = ss.getSheetByName(MAPPING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAPPING_SHEET_NAME);
    sheet.appendRow(['Tag', 'Color']);
  }
  return sheet;
}

/**
 * Extracts the first RJ-prefixed token, either bracketed [RJ…] or
 * from comma-separated values, ignoring a bare "RJ".
 */
function extractEventTag(tagStr) {
  if (!tagStr) return null;
  const str = String(tagStr);

  // 1) bracketed tags
  const bracketMatches = [...str.matchAll(/\[([^\]]+)\]/g)]
    .map(m => m[1])
    .filter(t => t.startsWith('RJ') && t !== 'RJ');
  if (bracketMatches.length) {
    Logger.log('extractEventTag: bracket match "' + bracketMatches[0] + '"');
    return bracketMatches[0];
  }

  // 2) fallback to comma-separated tokens
  const parts = str.split(',').map(p => p.trim());
  for (const part of parts) {
    if (part.startsWith('RJ') && part !== 'RJ') {
      Logger.log('extractEventTag: fallback token "' + part + '"');
      return part;
    }
  }
Logger.log('extractEventTag: no RJ tag extracted from "' + str + '"');
  return null;
}


// === OPTIONAL: Install a trigger for applyEventTagColors ===
function installApplyEventTagColorsTrigger() {
  ScriptApp.newTrigger('applyEventTagColors')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// === OPTIONAL: Clear existing applyEventTagColors triggers ===
function clearApplyEventTagColorsTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'applyEventTagColors') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * Converts "hsl(H,S%,L%)" into "#rrggbb"
 */
function hslToHex(hslStr) {
  const m = hslStr.match(/hsl\((\d+),\s*(\d+)%?,\s*(\d+)%?\)/i);
  if (!m) return hslStr;
  const h = +m[1], s = +m[2] / 100, l = +m[3] / 100;
  const [r, g, b] = hslToRgb(h, s, l);
  return '#' + toHex(r) + toHex(g) + toHex(b);
}

function hslToRgb(h, s, l) {
  h /= 360;
  let r, g, b;
  if (s === 0) {
    r = g = b = l;
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, h + 1/3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1/3);
  }
  return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}

function hue2rgb(p, q, t) {
  if (t < 0) t++;
  if (t > 1) t--;
  if (t < 1/6) return p + (q - p) * 6 * t;
  if (t < 1/2) return q;
  if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
  return p;
}

function toHex(n) {
  const h = n.toString(16);
  return h.length < 2 ? '0' + h : h;
}

/**
 * Generates a random soft (pastel) HSL color.
 * Varies both saturation (60–80%) and lightness (85–95%) for more distinct tints.
 */
function getRandomSoftColor() {
  const hue   = Math.floor(Math.random() * 360);
  const sat   = Math.floor(60 + Math.random() * 20);  // 60–80%
  const light = Math.floor(85 + Math.random() * 10);  // 85–95%
  return `hsl(${hue},${sat}%,${light}%)`;
}

/**
 * Generates a random pastel hex color.
 * Broadens the RGB range (180–255) so you get noticeable variation.
 */
function getRandomSoftHex() {
  const r = Math.floor(180 + Math.random() * 75);  // 180–255
  const g = Math.floor(180 + Math.random() * 75);
  const b = Math.floor(180 + Math.random() * 75);
  return '#' + toHex(r) + toHex(g) + toHex(b);
}