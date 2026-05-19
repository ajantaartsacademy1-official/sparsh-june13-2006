// ═══════════════════════════════════════════════════════════════
//  Stage Show — Google Apps Script Backend
//  Paste this entire file into script.google.com
//
//  DEPLOY STEPS:
//  1. Click "Deploy" → "New deployment"
//  2. Type: Web app
//  3. Execute as: Me
//  4. Who has access: Anyone
//  5. Click Deploy → copy the Web App URL
//  6. Paste that URL into the app's "Connect Google Drive" screen
// ═══════════════════════════════════════════════════════════════

const FILE_NAME = 'StageShow_Songs.json';
const XLS_NAME  = 'StageShow_Setlist.xlsx';

function doGet(e) {
  const action = (e.parameter.action || 'read');

  if (action === 'read')      return read();
  if (action === 'write')     return write(e.parameter.data);
  if (action === 'exportXls') return exportXls(e.parameter.data);

  return json({ error: 'Unknown action: ' + action });
}

// ── READ: return all songs as JSON ──────────────────────────────
function read() {
  const file = getFile(FILE_NAME);
  const data = file ? file.getBlob().getDataAsString() : '[]';
  return json({ songs: JSON.parse(data) });
}

// ── WRITE: save songs JSON to Drive ────────────────────────────
function write(dataStr) {
  if (!dataStr) return json({ error: 'No data provided' });
  const file = getFile(FILE_NAME);
  if (file) {
    file.setContent(dataStr);
  } else {
    DriveApp.createFile(FILE_NAME, dataStr, MimeType.PLAIN_TEXT);
  }
  return json({ ok: true });
}

// ── EXPORT XLS: build Excel and save to Drive ──────────────────
function exportXls(dataStr) {
  if (!dataStr) return json({ error: 'No data provided' });

  const songs = JSON.parse(dataStr);

  // Create a temporary Google Sheet
  const ss = SpreadsheetApp.create('__StageShow_tmp__');
  const sh = ss.getActiveSheet();
  sh.setName('Setlist');

  // Header row
  sh.appendRow(['#', 'Song Name', 'Singers', 'Karaoke Link', 'Added']);

  // Data rows
  songs.forEach((s, i) => {
    sh.appendRow([
      i + 1,
      s.name || '',
      (s.singers || []).join(', '),
      s.karaoke || '',
      s.addedAt ? new Date(s.addedAt).toLocaleDateString('en-IN') : ''
    ]);
  });

  // Style header
  const hdrRange = sh.getRange(1, 1, 1, 5);
  hdrRange.setFontWeight('bold')
          .setBackground('#7C3AED')
          .setFontColor('#ffffff')
          .setHorizontalAlignment('center');

  // Auto-resize columns
  sh.autoResizeColumns(1, 5);
  sh.setFrozenRows(1);

  // Export as xlsx blob
  const ssId   = ss.getId();
  const xlsBlob = Drive.Files.export(
    ssId,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  // Delete the temporary Sheet
  DriveApp.getFileById(ssId).setTrashed(true);

  // Replace old xlsx if exists
  const old = getFile(XLS_NAME);
  if (old) old.setTrashed(true);

  // Save new xlsx to Drive root
  DriveApp.createFile(XLS_NAME, xlsBlob);

  return json({ ok: true, file: XLS_NAME });
}

// ── HELPER: find file by name in Drive root ─────────────────────
function getFile(name) {
  const files = DriveApp.getFilesByName(name);
  return files.hasNext() ? files.next() : null;
}

// ── HELPER: return JSON response with CORS headers ──────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
