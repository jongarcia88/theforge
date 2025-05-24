/**
 * CONFIGURATION: fill these in
 */
const ALIAS_ADDRESS       = 'locations+log@yourdomain.com';
const DRIVE_FOLDER_ID     = '1bS3t8pvRO-0z7TsR3z6zjK_VssD3deEB';
const PROCESSED_FOLDER_ID = '1ocY9ZZCdqk-1aOLr9mQCe5UD-8diINH1';
const SPREADSHEET_ID      = '1sDKJn69zYapZrD343-iziOVpJr9SH4hvKEh3_QVXfU8ByWefQ2fDQ2pJ';
const SHEET_NAME          = 'History';

// Global month map for date parsing
const MONTH_MAP = {
  january:1, february:2, march:3, april:4, may:5, june:6,
  july:7, august:8, september:9, october:10, november:11, december:12
};

/**
 * 1. Process unread emails: extract screenshot attachments
 */
function processLocationEmails() {
  Logger.log('Starting processLocationEmails');
  const query   = `in:inbox is:unread to:${ALIAS_ADDRESS}`;
  const threads = GmailApp.search(query);
  Logger.log('Found %s threads', threads.length);
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      if (msg.isUnread()) {
        Logger.log('Email subject: %s', msg.getSubject());
        processOneMessage(msg);
        msg.markRead();
      }
    });
    thread.moveToArchive();
  });
  Logger.log('Finished processLocationEmails');
}

/**
 * 2. Handle one email: parse date, save & OCR attachments, summarize
 */
function processOneMessage(msg) {
  const subject = msg.getSubject();
  const m = subject.match(/(\d{4}-\d{2}-\d{2})/);
  if (!m) {
    Logger.log('No date in subject: %s', subject);
    return;
  }
  const dateString = m[1];
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const atts = msg.getAttachments();
  Logger.log('Attachments: %s', atts.length);
  const recs = atts.map((blob, i) => {
    const ext    = blob.getName().split('.').pop();
    const name   = `${dateString}_${i+1}.${ext}`;
    const file   = folder.createFile(blob.copyBlob()).setName(name);
    const url    = file.getUrl();
    Logger.log('Saved %s', name);

    // OCR via Vision API
    const visionResp = Vision.Images.annotate({
      requests: [{ image: { content: blob.getBytes() }, features: [{ type: 'TEXT_DETECTION', maxResults: 1 }] }]
    });
    const text     = visionResp.responses[0].fullTextAnnotation?.text || '';
    const lines    = text.split('\n').map(l=>l.trim()).filter(l=>l);
    const location = lines[0] || 'Unknown';
    const tmMatch  = text.match(/(\d{1,2}:\d{2}\s?(?:AM|PM))/i);
    const timeText = tmMatch ? tmMatch[1].toUpperCase() : '';
    return { dateString, location, timeText, url };
  });
  appendSummary(buildSummary(recs));
}

/**
 * 3. Build summary rows from OCR records
 */
function buildSummary(items) {
  const tz = Session.getScriptTimeZone();
  function parseDT(d, t) {
    return Utilities.parseDate(`${d} ${t}`, tz, 'yyyy-MM-dd hh:mm a');
  }
  const byLoc = {};
  items.forEach(r => {
    byLoc[r.location] = byLoc[r.location] || [];
    byLoc[r.location].push({ ...r, dt: parseDT(r.dateString, r.timeText) });
  });
  return Object.keys(byLoc).flatMap(loc => {
    const arr   = byLoc[loc].sort((a,b) => a.dt - b.dt);
    const first = arr[0], last = arr[arr.length - 1];
    const mins  = Math.round((last.dt - first.dt) / 60000);
    return [[
      first.dateString,
      loc,
      Utilities.formatDate(first.dt, tz, 'hh:mm a'),
      Utilities.formatDate(last.dt, tz, 'hh:mm a'),
      mins,
      first.url
    ]];
  });
}

/**
 * 4. Append summary rows to History sheet
 */
function appendSummary(rows) {
  if (!rows.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Generate Clean History: parse raw History into sorted sessions
 */

/**
 * Rebuilds “Clean History” with:
 *  - normalized dates for Today/Yesterday
 *  - ImageID + ImageURL columns
 *  - sequence‐gap warnings for secondaries
 */
function generateCleanHistory() {
  Logger.log('Generating Clean History');
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const raw   = sheet.getDataRange().getValues();
  const tz    = ss.getSpreadsheetTimeZone();

  // (Re)create the Clean History sheet
  let clean = ss.getSheetByName('Clean History');
  if (!clean) clean = ss.insertSheet('Clean History');
  else         clean.clear();

  // New header: added PrimaryID before ImageID
  clean.appendRow([
    'Date','Location','Start time','End time','Duration',
    'PrimaryID','ImageID','Image','Warning'
  ]);

  // Regexes for date detection
  const fullDateRe = /\b(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)\w*,\s*(Jan\w*|Feb\w*|Mar\w*|Apr\w*|May|Jun\w*|Jul\w*|Aug\w*|Sep\w*|Oct\w*|Nov\w*)\s+(\d{1,2})\b/i;
  const keywordRe  = /\b(?:Earlier today|Today|Yesterday)\b/i;

  // Session‐parsing regex
  const sessRe = /(?:\|\s*([^|]+?)\s*\|\s*)?(\d{1,2}:\d{2}\s?(?:AM|PM))\s*[-–]\s*(\d{1,2}:\d{2}\s?(?:AM|PM))\s*\(([^)]+)\)/gi;

  const sessions = [];

  // Trackers for primary/secondary logic
  let currentDate       = '';
  let currentPrimaryNum = -Infinity;
  let lastImageNum      = -Infinity;
  let currentPrimaryID  = '';

  raw.slice(1).forEach(row => {
    const txt     = String(row[3] || '');
    const imgCell = row[5] || row[4] || '';

    // --- pull ImageID and numeric suffix ---
    let imageID = '';
    const idMatch = String(imgCell).match(/\/d\/([^/]+)\b/);
    if (idMatch) {
      try {
        imageID = DriveApp.getFileById(idMatch[1]).getName();
      } catch (e) {
        imageID = '';
      }
    }
    const numMatch = imageID.match(/_(\d+)\./);
    const imgNum   = numMatch ? parseInt(numMatch[1], 10) : null;

    // --- detect a new primary row ---
    let dm;
    if ((dm = txt.match(fullDateRe))) {
      // e.g. "Tuesday, April 8"
      currentDate       = Utilities.formatDate(
        new Date(
          new Date().getFullYear(),
          MONTH_MAP[dm[1].toLowerCase()] - 1,
          parseInt(dm[2], 10)
        ),
        tz,
        'M/d/yyyy'
      );
      currentPrimaryNum = (imgNum !== null ? imgNum : -Infinity);
      lastImageNum      = currentPrimaryNum;
      currentPrimaryID  = imageID;
    } else if (keywordRe.test(txt)) {
      // "Today", "Earlier today", "Yesterday"
      const lc = txt.toLowerCase();
      if (lc === 'today' || lc === 'earlier today') {
        currentDate = Utilities.formatDate(new Date(), tz, 'M/d/yyyy');
      } else {
        const d = new Date();
        d.setDate(d.getDate() - 1);
        currentDate = Utilities.formatDate(d, tz, 'M/d/yyyy');
      }
      currentPrimaryNum = (imgNum !== null ? imgNum : -Infinity);
      lastImageNum      = currentPrimaryNum;
      currentPrimaryID  = imageID;
    }

    // --- parse each session in the text ---
    let m;
    while ((m = sessRe.exec(txt)) !== null) {
      const loc = m[1]
        ? m[1].trim()
        : (() => {
            const before = txt.slice(0, m.index);
            const p      = before.lastIndexOf('|');
            return p > -1 ? before.slice(p + 1).trim() : before.trim();
          })();
      const start = m[2];
      const end   = m[3];
      const dur   = m[4];

      // --- sequence‐gap warning for secondaries ---
      let warning = '';
      if (imgNum !== null && imgNum > currentPrimaryNum) {
        if (lastImageNum !== -Infinity && imgNum !== lastImageNum + 1) {
          warning = `Gap: expected ${lastImageNum + 1}, got ${imgNum}`;
        }
        lastImageNum = imgNum;
      }

      sessions.push({
        date:      currentDate,
        loc,
        start,
        end,
        dur,
        primaryID: currentPrimaryID,
        imageID,
        imageURL:  imgCell,
        warning
      });
    }
  });

  // --- sort by start time ---
  sessions.sort((a, b) => {
    const toMin = t => {
      const mm = t.match(/(\d{1,2}):(\d{2})\s?(AM|PM)/i);
      let h = +mm[1], n = +mm[2], p = mm[3].toUpperCase();
      if (p === 'PM' && h < 12) h += 12;
      if (p === 'AM' && h === 12) h = 0;
      return h * 60 + n;
    };
    return toMin(a.start) - toMin(b.start);
  });

  // --- write values and set green background for primaries ---
  if (sessions.length) {
    // write all columns
    clean
      .getRange(2, 1, sessions.length, 9)
      .setValues(sessions.map(s => [
        s.date, s.loc, s.start, s.end, s.dur,
        s.primaryID, s.imageID, s.imageURL, s.warning
      ]));

    // highlight primary ImageID cells
    const bg = sessions.map(s => [
      s.imageID === s.primaryID ? '#00FF00' : ''
    ]);
    clean
      .getRange(2, 7, sessions.length, 1)
      .setBackgrounds(bg);
  }

  Logger.log('Clean History rebuilt with PrimaryID, ImageID highlights, and warnings.');
}

/**
 * Send processing report by email
 */
function sendProcessingReport(processed, failed) {
  const recipient = Session.getEffectiveUser().getEmail();
  const subject   = `Processing Report: ${processed.length} OK, ${failed.length} Errors`;
  const body      = `Processed files:\n${processed.join('\n')}\n\nFailures:\n${failed.map(f=>f.name+' - '+f.error).join('\n')}`;
  MailApp.sendEmail(recipient, subject, body);
}

/**
 * Bulk process folder files, move processed, regenerate & report
 */
function processFolderScreenshots() {
  Logger.log('Starting batch process');
  let src, dst;
  try { src = DriveApp.getFolderById(DRIVE_FOLDER_ID); }
  catch (e) { Logger.log('Src folder error: ' + e); return; }
  try { dst = DriveApp.getFolderById(PROCESSED_FOLDER_ID); }
  catch (e) { Logger.log('Dst folder error: ' + e); return; }

  const files    = src.getFiles();
  const processed = [], failed = [];
  let count = 0, BATCH = 10;
  while (files.hasNext() && count < BATCH) {
    const f = files.next();
    try {
      processScreenshotFile(f);
      processed.push(f.getName());
      dst.addFile(f);
      src.removeFile(f);
    } catch (e) {
      Logger.log(`Error ${f.getName()}: ${e}`);
      failed.push({ name: f.getName(), error: e.toString() });
    }
    count++;
  }
  Logger.log(`Batch done: ${processed.length} processed, ${failed.length} failed`);
  generateCleanHistory();
  sendProcessingReport(processed, failed);
}

/**
 * OCR one file and append raw to History
 */
function processScreenshotFile(file) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const blob = file.getBlob();
  const url  = file.getUrl();
  Logger.log('OCR’ing %s', file.getName());

  // OCR the image by creating a temporary Google Doc
  const resource = { title: `ocr-temp-${Date.now()}`, mimeType: blob.getContentType() };
  const ocrFile  = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'en', convert: true });
  const fullText = DocumentApp.openById(ocrFile.id).getBody().getText();
  Drive.Files.remove(ocrFile.id);

  const firstLine = fullText.split('\n')[0] || '';
  const tmMatch   = fullText.match(/(\d{1,2}:\d{2}\s?(?:AM|PM))/i);
  const timeText  = tmMatch ? tmMatch[1].toUpperCase() : '';

  // --- EXTENDED DATE LOGIC ---
  const today = new Date();
  const todayFormatted = Utilities.formatDate(today, tz, 'M/d/yyyy');
  const yesterdayDate = new Date(today);
  yesterdayDate.setDate(yesterdayDate.getDate() - 1);
  const yesterdayFormatted = Utilities.formatDate(yesterdayDate, tz, 'M/d/yyyy');

  let dateCell = '';
  // 1) Full "Weekday, Month D" e.g. "Tuesday, April 8"
  const dm = fullText.match(/\b(?:Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday),\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})\b/i);
  if (dm) {
    dateCell = Utilities.formatDate(
      new Date(
        today.getFullYear(),
        MONTH_MAP[dm[1].toLowerCase()] - 1,
        parseInt(dm[2], 10)
      ),
      tz,
      'M/d/yyyy'
    );
  } else if (/\b(?:Earlier Today|Today)\b/i.test(fullText)) {
    // 2) "Earlier Today" or "Today"
    dateCell = todayFormatted;
  } else if (/\bYesterday\b/i.test(fullText)) {
    // 3) "Yesterday"
    dateCell = yesterdayFormatted;
  } else {
    // 4) Single weekday name -> most recent within last 7 days
    const dowMatch = fullText.match(/\b(?:Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)\b/i);
    if (dowMatch) {
      const weekdayName = dowMatch[0].toLowerCase();
      const WEEKDAY_MAP = { sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6 };
      const targetDay = WEEKDAY_MAP[weekdayName];
      let d = new Date(today);
      for (let i = 0; i < 7; i++) {
        if (d.getDay() === targetDay) break;
        d.setDate(d.getDate() - 1);
      }
      dateCell = Utilities.formatDate(d, tz, 'M/d/yyyy');
    }
  }
  // --- END DATE LOGIC ---

  const imageID   = file.getName();
  const primaryID = dateCell ? imageID : '';

  const hist = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  hist.appendRow([
    dateCell,
    imageID,
    primaryID,
    firstLine,
    timeText,
    fullText.replace(/\n/g, ' | '),
    url
  ]);
  Logger.log('Appended %s', file.getName());
}


/**
 * Create auto trigger every minute
 */
function createFolderProcessingTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processFolderScreenshots')
    .forEach(ScriptApp.deleteTrigger);
  ScriptApp.newTrigger('processFolderScreenshots').timeBased().everyMinutes(1).create();
  Logger.log('Trigger set up for every minute');
}

/**
 * Custom menu
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Location OCR')
    .addItem('Process Emails','processLocationEmails')
    .addItem('Run OCR Test','runOcrTest')
    .addItem('Generate Clean History','generateCleanHistory')
    .addItem('Process Folder','processFolderScreenshots')
    .addItem('Enable Auto','createFolderProcessingTrigger')
    .addToUi();
}

/**
 * Alias runOcrTest to bulk process for quick test
 */
function runOcrTest() {
  processFolderScreenshots();
}
