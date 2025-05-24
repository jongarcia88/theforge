// Custody Evidence Tracker - Final Script with Event-Specific Raw Notes
// Multiple Events, Date Fallback, Source Column, Impact-Oriented Category, Rich Text Hyperlinks

const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const DRIVE_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
const EMAIL_LABEL = '_Tracker/TO_PROCESS';
const PROCESSED_LABEL = '_Tracker/Processed';

function processCustodyEmails() {
  Logger.log('Starting custody email processing');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Event Log setup
  let logSheet = ss.getSheetByName('Event Log');
  if (!logSheet) logSheet = ss.insertSheet('Event Log');
  if (logSheet.getLastRow() < 1) {
    logSheet.appendRow([
      'Processed Timestamp','Who','Category','Summary','Evidence Row','Attachments Count','Success','Error','GPT JSON'
    ]);
  }

  // Evidence sheet must include 'Source' as column K
  const evidenceSheet = ss.getSheetByName('Evidence');
  if (!evidenceSheet) throw new Error("Sheet tab 'Evidence' not found.");

  const dateRegex = /(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2})(AM|PM)?)?/i;
  const threads = GmailApp.search(`label:"${EMAIL_LABEL}"`);
  Logger.log(`Found ${threads.length} threads`);

  const toProcessLabel = GmailApp.getUserLabelByName(EMAIL_LABEL);
  let processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!processedLabel) processedLabel = GmailApp.createLabel(PROCESSED_LABEL);

  threads.forEach((thread, ti) => {
    Logger.log(`Processing thread ${ti+1}/${threads.length}`);
    let threadSuccess = true;
    thread.getMessages().forEach((message, mi) => {
      Logger.log(`Processing message ${mi+1}`);
      const subject = message.getSubject().trim();
      let subjectDate = null;
      const m = subject.match(dateRegex);
      if (m) {
        let mo = +m[1], da = +m[2], yr = +m[3];
        if (yr < 100) yr += 2000;
        let hr = 0, min = 0;
        if (m[4]) {
          hr = +m[4];
          const ap = (m[5]||'').toUpperCase();
          if (ap==='PM'&&hr<12) hr+=12;
          if (ap==='AM'&&hr===12) hr=0;
        }
        subjectDate = new Date(yr, mo-1, da, hr, min);
      }
      const attachments = message.getAttachments();
      const rawBody = message.getPlainBody().trim();
      let splitResult;
      try {
        splitResult = splitAndProcessEvents(rawBody, subject);
      } catch (e) {
        threadSuccess = false;
        logSheet.appendRow([new Date(), subject, '', '', '', attachments.length, false, 'Split error', e.rawJson||e.toString()]);
        return;
      }
      const { events, rawJson } = splitResult;

      events.forEach((evt, ei) => {
        let success = true, err = '', evidenceRow = '';
        try {
          // Determine date
          let eventDate = message.getDate();
          if (evt.date && !isNaN(Date.parse(evt.date))) {
            eventDate = new Date(evt.date);
          } else if (subjectDate) {
            eventDate = subjectDate;
          }

          // Involved parties
          const involved = Array.isArray(evt.involved_parties)
            ? evt.involved_parties.join(', ')
            : (evt.involved_parties||'');

          // Upload attachments
          const urlList = attachments.map(f => uploadToDrive(f));

          // Build rich text hyperlinks
          const textUrls = urlList.join('\n');
          const rtBuilder = SpreadsheetApp.newRichTextValue().setText(textUrls);
          let pos = 0;
          urlList.forEach(url => {
            const len = url.length;
            rtBuilder.setLinkUrl(pos, pos+len, url);
            pos += len+1;
          });
          const linkRichText = rtBuilder.build();

          // Use event-specific raw_text only
          const sourceText = evt.raw_text || '';

          // Append to Evidence, placeholder for links in col J
          const row = [
            eventDate,
            evt.category || '',
            evt.who || subject,
            evt.summary || '',
            evt.description || '',
            involved,
            evt.impact_children || '',
            evt.impact_household || '',
            evt.impact_finance || '',
            '',
            sourceText
          ];
          evidenceSheet.appendRow(row);
          evidenceRow = evidenceSheet.getLastRow();
          evidenceSheet.getRange(evidenceRow, 10).setRichTextValue(linkRichText);
        } catch (e) {
          success = false;
          err = e.toString();
          threadSuccess = false;
        }
        logSheet.appendRow([
          new Date(), subject, evt.category||'', evt.summary||'', evidenceRow, attachments.length, success, err, rawJson
        ]);
      });
    });
    if (threadSuccess) {
      if (toProcessLabel) thread.removeLabel(toProcessLabel);
      thread.addLabel(processedLabel);
    }
  });
  Logger.log('Processing complete');
}

function splitAndProcessEvents(body, subject) {
  const prompt = `You are a legal documentation assistant. The email body may describe multiple distinct events.\nEmail Body:\n${body}\n\nFor each event, infer a date in ISO format (null if none); extract who (default subject), a vivid category that reflects the event's impact on the children, the household, and family finances; one-sentence summary; detailed description; involved_parties (array); impact_children (describe how the event benefits or harms the children); impact_household (describe how it affects daily household routines); impact_finance (describe financial consequences); and raw_text snippet -- just the text describing that specific event. Return a pure JSON array of objects with all keys present (empty or null if missing).`;
  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${OPENAI_API_KEY}` },
    payload: JSON.stringify({ model: 'gpt-4', messages: [{ role: 'user', content: prompt }] })
  });
  if (resp.getResponseCode() !== 200) {
    const err = new Error(`OpenAI error ${resp.getResponseCode()}`);
    err.rawJson = resp.getContentText();
    throw err;
  }
  let content = JSON.parse(resp.getContentText()).choices[0].message.content;
  content = content.replace(/^```(?:json)?\s*/i,'').replace(/```$/i,'').trim();
  let events;
  try {
    events = JSON.parse(content);
  } catch (e) {
    const err = new Error('Invalid JSON');
    err.rawJson = content;
    throw err;
  }
  events = events.map(evt => {
    evt.who = evt.who ? evt.who.replace(/\b(the narrator|tracker)\b/gi, 'Jonathan') : 'Jonathan';
    return evt;
  });
  return { events, rawJson: content };
}

function uploadToDrive(file) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const f = folder.createFile(file);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getUrl();
}

// Setup: configure OPENAI_API_KEY & DRIVE_FOLDER_ID; ensure 'Evidence' sheet with 'Source'; create trigger manually.
