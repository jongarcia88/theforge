const SOURCE_FOLDER_ID   = '1qdlv2s8bGildbB0GWyQzzoQFo5YrAYzH';  // ← your “unassigned” folder
const ASSIGNED_FOLDER_ID = '1OfIjvE8jaaxukihr9ALHDbRO7mHZYnmc';  // ← your “Assigned” folder

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📎 Receipts')
    .addItem('Attach Receipt…', 'showPicker')
    .addToUi();
}

function showPicker() {
const tpl = HtmlService.createTemplateFromFile('Picker');
tpl.sourceFolderId   = SOURCE_FOLDER_ID;
tpl.assignedFolderId = ASSIGNED_FOLDER_ID;
  
  const html = tpl.evaluate()
    .setWidth(700)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Receipt');
}

/**
 * @return [{ id, name, created }]
 */
function getImageFilesInFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const out    = [];
  const it     = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getMimeType().startsWith('image/')) {
      out.push({
        id:      f.getId(),
        name:    f.getName(),
        created: f.getDateCreated().toISOString()
      });
    }
  }
  return out;
}

function moveAndInsert(fileId, targetFolderId) {
  // 1) move the file
  const file   = DriveApp.getFileById(fileId);
  const target = DriveApp.getFolderById(targetFolderId);
  file.moveTo(target);
  const url    = file.getUrl();

  // 2) identify the sheet, header, and target column
  const sheet    = SpreadsheetApp.getActiveSheet();
  const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0];
  const colIndex = headers.indexOf("Evidence Link or Notes") + 1;
  if (colIndex === 0) {
    throw new Error('Cannot find header "Evidence Link or Notes" in row 1');
  }
  const rowIndex = sheet.getActiveCell().getRow();
  const cell     = sheet.getRange(rowIndex, colIndex);

  // 3) build the new full text (append to any existing lines)
  const prevText = (cell.getValue() + '').trim();
  const lines    = prevText
    ? prevText.split('\n').concat([url])
    : [ url ];
  const fullText = lines.join('\n');

  // 4) build a RichTextValue so each line – if it starts with "http…" – becomes a link
  const builder = SpreadsheetApp.newRichTextValue().setText(fullText);
  let offset = 0;
  lines.forEach(line => {
    const len = line.length;
    if (/^https?:\/\//.test(line)) {
      builder.setLinkUrl(offset, offset + len, line);
    }
    offset += len + 1;  // move past the "\n"
  });

  // 5) write it back
  cell.setRichTextValue(builder.build());
}


/**
 * Move every file in fileIds to targetFolderId,
 * then append all of their URLs (one per line) into the
 * "Evidence Link or Notes" column of the active row, as clickable links.
 */
function moveAndInsertAll(fileIds, targetFolderId) {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0];
  const colIdx  = headers.indexOf("Evidence Link or Notes") + 1;
  if (colIdx === 0) {
    throw new Error('Header "Evidence Link or Notes" not found in row 1');
  }
  const rowIdx  = sheet.getActiveCell().getRow();
  const cell    = sheet.getRange(rowIdx, colIdx);

  // Gather existing lines
  const prevText = (cell.getValue() + '').trim();
  const lines    = prevText ? prevText.split('\n') : [];

  // Move each file & collect its URL
  fileIds.forEach(id => {
    const file   = DriveApp.getFileById(id);
    const folder = DriveApp.getFolderById(targetFolderId);
    file.moveTo(folder);
    lines.push(file.getUrl());
  });

  // Build a single RichTextValue with clickable links
  const fullText = lines.join('\n');
  const builder  = SpreadsheetApp.newRichTextValue().setText(fullText);
  let offset = 0;
  lines.forEach(line => {
    if (/^https?:\/\//.test(line)) {
      builder.setLinkUrl(offset, offset + line.length, line);
    }
    offset += line.length + 1;  // +1 for the newline
  });
  cell.setRichTextValue(builder.build());
}
