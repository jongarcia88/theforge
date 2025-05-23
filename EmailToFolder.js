const ALIAS              = 'garcia.jonathan+evidence@gmail.com';
const EVIDENCE_FOLDER_ID = '1qdlv2s8bGildbB0GWyQzzoQFo5YrAYzH';
const LOG_SHEET_ID       = '1HRFI5x6dbTkknGhX7hCrTit-Q-SwSwR_vmLwJe_OfH4';
const PROCESSED_LABEL    = 'Evidence-Processed';
const LOG_SHEET_NAME     = 'Evidence Log';


function checkGmailAndSaveEvidence() {
  // Build Gmail search for alias + attachments, skipping already-processed
  const query = [
    `to:${ALIAS}`,
    'has:attachment',
    `-label:${PROCESSED_LABEL}`
  ].join(' ');

  const threads = GmailApp.search(query, 0, 50);
  if (threads.length === 0) return;

  // Drive folder & Gmail label
  const folder = DriveApp.getFolderById(EVIDENCE_FOLDER_ID);
  const label  = GmailApp.getUserLabelByName(PROCESSED_LABEL)
                || GmailApp.createLabel(PROCESSED_LABEL);

  // Open your log sheet and ensure the “Evidence Log” tab exists
  const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    // Header row
    sheet.appendRow([
      'Timestamp',
      'File Name',
      'Tags',
      'Sender',
      'Subject',
      'Drive URL'
    ]);
  }

  // Process each thread & attachment
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const subject = msg.getSubject() || '';
      const sender  = msg.getFrom();
      const tags    = extractHashtags(subject);

      msg.getAttachments().forEach(att => {
        const file = folder.createFile(att.copyBlob());
        file.setName(att.getName());
        sheet.appendRow([
          new Date(),
          att.getName(),
          tags.join(','),
          sender,
          subject,
          file.getUrl()
        ]);
      });
    });
    thread.addLabel(label);
  });
}

/**
 * returns an array of lowercase tags (no #), e.g.
 * "#alpha #parenting fun" → ['alpha','parenting']
 */
function extractHashtags(text) {
  const matches = text.match(/#([A-Za-z0-9_]+)/g) || [];
  return matches.map(m => m.substring(1).toLowerCase());
}