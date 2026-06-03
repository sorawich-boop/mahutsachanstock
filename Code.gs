// ============================================================
// Mahutsachan Cafe — Stock Tracker API  v2
// Google Apps Script (standalone project)
// ============================================================
// SETUP:
//   1. Create TWO new blank Google Sheets (one for bar, one for containers)
//   2. Go to https://script.google.com → New project → paste this file
//   3. Fill in the four IDs below
//   4. Deploy → New deployment → Web App
//      · Execute as: Me
//      · Who has access: Anyone
//   5. Paste the Web App URL into stock-tracker-v2.html
//
// The script auto-creates Config + Log tabs on first use.
// ============================================================

// ── YOUR IDs ────────────────────────────────────────────────
const BAR_SPREADSHEET_ID       = '1mnZZoNfg4MEmtwd-xzEzgma78OJpsDMasCA0KnQbaJY';
const CONTAINER_SPREADSHEET_ID = '1BGBIfk-XAGO4TH41lT-WMK3_8BsDTw4HbGc5_9dcOv0';
const BAR_DRIVE_FOLDER_ID       = '1kkjz8bN8SITKzcZAKQBoguorGNp3xFRH';
const CONTAINER_DRIVE_FOLDER_ID = '1o5zVa-L0q75CkAzR53P_r4Vkl54EOOrN';
// ────────────────────────────────────────────────────────────

// ── SHEET STRUCTURE ─────────────────────────────────────────
// Config tab  →  name | unit | cat | min | order
// Log tab     →  date | by | item | recv | used | remaining | kind | photos | ts
// ────────────────────────────────────────────────────────────

const CONFIG_HEADERS = ['name','unit','cat','min','order'];
const LOG_HEADERS    = ['date','by','item','recv','used','remaining','kind','photos','ts'];

// ============================================================
// HTTP HANDLERS
// ============================================================

// ── Run this from the editor to authorize Drive access ──────
function testDrive() {
  try {
    const barFolder  = DriveApp.getFolderById(BAR_DRIVE_FOLDER_ID);
    const contFolder = DriveApp.getFolderById(CONTAINER_DRIVE_FOLDER_ID);
    Logger.log('✅ Bar folder: ' + barFolder.getName());
    Logger.log('✅ Container folder: ' + contFolder.getName());
    Logger.log('Drive access OK');
  } catch(e) {
    Logger.log('❌ Drive error: ' + e.message);
  }
}

function doGet(e) {
  // Guard: e is undefined when run from the editor directly — that's normal
  if (!e || !e.parameter) {
    return jsonOut({ ok: false, error: 'Run via Web App URL, not from editor' });
  }
  // Quick test: open WebAppURL?action=test in browser
  if (e.parameter.action === 'test') {
    try {
      const barSS  = SpreadsheetApp.openById(BAR_SPREADSHEET_ID).getName();
      const contSS = SpreadsheetApp.openById(CONTAINER_SPREADSHEET_ID).getName();
      return jsonOut({ ok: true, bar: barSS, container: contSS, ts: new Date().toISOString() });
    } catch(err) {
      return jsonOut({ ok: false, error: err.message });
    }
  }
  if (e.parameter.action === 'testDrive') {
    try {
      const barFolder  = DriveApp.getFolderById(BAR_DRIVE_FOLDER_ID).getName();
      const contFolder = DriveApp.getFolderById(CONTAINER_DRIVE_FOLDER_ID).getName();
      return jsonOut({ ok: true, barFolder: barFolder, contFolder: contFolder, msg: 'Drive OK via Web App' });
    } catch(err) {
      return jsonOut({ ok: false, error: err.message });
    }
  }
  return jsonOut(handleRequest('GET', e.parameter, null));
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(_) {}
  return jsonOut(handleRequest('POST', e.parameter, body));
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ROUTER
// ============================================================

function handleRequest(method, params, body) {
  const action = (method === 'GET' ? params.action : body?.action) || 'getStock';

  try {
    switch(action) {

      case 'getStock':
        return {
          success: true,
          bar:              getCurrentStock(BAR_SPREADSHEET_ID),
          container:        getCurrentStock(CONTAINER_SPREADSHEET_ID),
          barCatalog:       getCatalog(BAR_SPREADSHEET_ID),
          containerCatalog: getCatalog(CONTAINER_SPREADSHEET_ID),
          barHistory:       getRecentLog(BAR_SPREADSHEET_ID, 80),
          containerHistory: getRecentLog(CONTAINER_SPREADSHEET_ID, 80),
          lastUpdated:      new Date().toISOString()
        };

      case 'addEntry': {
        const { type, date, by, recv, used } = body;
        if (!type || !date) return err('Missing type or date');
        const ssId = ssFor(type);
        return addEntry(ssId, date, by || '', recv || {}, used || {});
      }

      case 'saveCatalog': {
        const { type, catalog } = body;
        if (!type || !catalog) return err('Missing type or catalog');
        return saveCatalog(ssFor(type), catalog);
      }

      case 'physicalCount': {
        const { type, date, by, actual, label } = body;
        if (!type || !date || !actual) return err('Missing fields');
        return addPhysicalCount(ssFor(type), date, by || '', actual, label || '');
      }

      case 'adjustStock': {
        // Compensate for a deleted history entry by writing reverse rows to Log
        const { type, date, by, adjustments } = body;
        if (!type || !adjustments) return err('Missing fields');
        return writeAdjustments(ssFor(type), date, by || 'Deleted entry', adjustments);
      }

      case 'initStock': {
        // Push initial stock snapshot when Log is empty (first setup)
        const { type, stockSnapshot, by } = body;
        if (!type || !stockSnapshot) return err('Missing type or stockSnapshot');
        return initStock(ssFor(type), stockSnapshot, by || 'System');
      }

      case 'uploadPhoto': {
        const { type, date, itemName, filename, mimeType, data } = body;
        if (!type || !date || !itemName || !data) return err('Missing fields');
        return uploadPhoto(type, date, itemName, filename || 'photo.jpg', mimeType || 'image/jpeg', data);
      }

      case 'deletePhotos': {
        // fileUrls = array of Google Drive file URLs to delete
        const { fileUrls } = body;
        if (!fileUrls || !fileUrls.length) return { success: true, deleted: 0 };
        return deletePhotos(fileUrls);
      }

      case 'setStock': {
        // Directly set a single item's stock via an 'adjust' Log row
        const { type, itemName, newStock, by } = body;
        if (!type || !itemName || newStock === undefined) return err('Missing fields');
        return setStock(ssFor(type), itemName, parseFloat(newStock), by || 'Manual edit');
      }

      case 'deleteFolders': {
        // Delete item-level Drive folders for a specific date entry
        const { type, date, itemNames } = body;
        if (!type || !date || !itemNames || !itemNames.length) return { success: true, deleted: 0 };
        return deleteFolders(type, date, itemNames);
      }

      default:
        return err('Unknown action: ' + action);
    }
  } catch(e) {
    return err(e.message);
  }
}

function ssFor(type) {
  return type === 'bar' ? BAR_SPREADSHEET_ID : CONTAINER_SPREADSHEET_ID;
}

function err(msg) { return { success: false, error: msg }; }

// ============================================================
// SHEET INITIALISATION
// ============================================================

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    // Style header row
    const hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight('bold');
    hRange.setBackground('#1c1916');
    hRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getConfigSheet(ssId) {
  return getOrCreateSheet(SpreadsheetApp.openById(ssId), 'Config', CONFIG_HEADERS);
}

function getLogSheet(ssId) {
  return getOrCreateSheet(SpreadsheetApp.openById(ssId), 'Log', LOG_HEADERS);
}

// ============================================================
// CATALOG  (Config tab)
// ============================================================

function getCatalog(ssId) {
  const sheet = getConfigSheet(ssId);
  if (sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      name:  String(r[0]).trim(),
      unit:  String(r[1]).trim(),
      cat:   String(r[2]).trim() || 'อื่นๆ',
      min:   parseFloat(r[3])   || 0,
      order: parseInt(r[4])     || 0
    }))
    .sort((a, b) => a.order - b.order);
}

function saveCatalog(ssId, catalog) {
  const sheet = getConfigSheet(ssId);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
  }

  if (!catalog.length) return { success: true, count: 0 };

  const rows = catalog.map((item, i) => [
    item.name, item.unit, item.cat || 'อื่นๆ', item.min || 0, i + 1
  ]);
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  return { success: true, count: rows.length };
}

// ============================================================
// CURRENT STOCK  (reads last Log row per item)
// ============================================================

function getCurrentStock(ssId) {
  const sheet = getLogSheet(ssId);
  if (sheet.getLastRow() < 2) return { items: [] };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  // For each item, track the last 'remaining' value we've seen
  const latest = {};
  rows.forEach(r => {
    const item = String(r[2]).trim();
    if (!item) return;
    const kind = String(r[6]).trim();
    if (kind === 'entry' || kind === 'count' || kind === 'baseline' || kind === 'adjust') {
      latest[item] = parseFloat(r[5]) || 0;
    }
  });

  const catalog = getCatalog(ssId);
  const items = catalog.map(c => ({
    name:  c.name,
    unit:  c.unit,
    stock: latest[c.name] !== undefined ? latest[c.name] : 0
  }));

  return { items };
}

// ============================================================
// ADD ENTRY  (appends rows to Log tab)
// ============================================================

function addEntry(ssId, date, by, recv, used) {
  const catalog = getCatalog(ssId);
  if (!catalog.length) return err('Catalog is empty — add items first');

  // Build current stock map
  const stockData = getCurrentStock(ssId);
  const stockMap  = {};
  stockData.items.forEach(i => { stockMap[i.name] = i.stock; });

  const sheet = getLogSheet(ssId);
  const ts    = new Date().toISOString();
  const rows  = [];

  catalog.forEach(item => {
    const r = parseFloat(recv[item.name]) || 0;
    const u = parseFloat(used[item.name]) || 0;
    if (r === 0 && u === 0) return; // skip unchanged items

    const prev      = stockMap[item.name] || 0;
    const remaining = Math.max(0, prev + r - u);
    stockMap[item.name] = remaining; // keep running total for this batch

    rows.push([date, by, item.name, r || '', u || '', remaining, 'entry', '', ts]);
  });

  if (!rows.length) return { success: false, error: 'No changes' };
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);

  return { success: true, count: rows.length };
}

// ============================================================
// PHYSICAL COUNT  (appends count rows to Log tab)
// ============================================================

function addPhysicalCount(ssId, date, by, actual, label) {
  const catalog = getCatalog(ssId);
  if (!catalog.length) return err('Catalog is empty');

  const stockData = getCurrentStock(ssId);
  const stockMap  = {};
  stockData.items.forEach(i => { stockMap[i.name] = i.stock; });

  const sheet = getLogSheet(ssId);
  const ts    = new Date().toISOString();
  const rows  = [];
  const diffs = {};

  catalog.forEach(item => {
    const actVal = actual[item.name] !== undefined
      ? parseFloat(actual[item.name])
      : stockMap[item.name] || 0;
    const sys  = stockMap[item.name] || 0;
    const diff = actVal - sys;
    diffs[item.name] = diff;

    rows.push([
      date, by, item.name,
      sys,  '',          // recv = system stock at time of count (for diff reconstruction)
      actVal,            // remaining = actual count
      'count',
      label || '',       // reuse photos col for label
      ts
    ]);
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  }

  return { success: true, diffs, count: rows.length };
}

// ============================================================
// ADJUST STOCK  (compensate for deleted history entry)
// adjustments = [{ name, recv, used }, ...] — the ORIGINAL entry values
// We reverse them: used becomes recv, recv becomes used
// ============================================================

function writeAdjustments(ssId, date, by, adjustments) {
  const stockData = getCurrentStock(ssId);
  const stockMap  = {};
  stockData.items.forEach(i => { stockMap[i.name] = i.stock; });

  const sheet = getLogSheet(ssId);
  const ts    = new Date().toISOString();
  const rows  = [];

  adjustments.forEach(a => {
    const origRecv = parseFloat(a.recv) || 0;
    const origUsed = parseFloat(a.used) || 0;
    if (origRecv === 0 && origUsed === 0) return;

    // Reverse: subtract what was received, add back what was used
    const prev      = stockMap[a.name] !== undefined ? stockMap[a.name] : 0;
    const remaining = prev - origRecv + origUsed;
    stockMap[a.name] = remaining;

    rows.push([date, by, a.name, origUsed || '', origRecv || '', remaining, 'adjust', '', ts]);
  });

  if (!rows.length) return { success: true, count: 0 };
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  return { success: true, count: rows.length };
}

// ============================================================
// INIT STOCK  (write baseline snapshot to Log on first setup)
// stockSnapshot = [{ name, stock }, ...]
// ============================================================

function initStock(ssId, stockSnapshot, by) {
  const sheet = getLogSheet(ssId);

  // Only allow if Log is truly empty (header row only)
  if (sheet.getLastRow() > 1) {
    return { success: false, error: 'Log already has data — use addEntry instead' };
  }

  if (!stockSnapshot.length) return { success: true, count: 0 };

  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const ts   = new Date().toISOString();

  const rows = stockSnapshot
    .filter(i => i.name && i.stock !== undefined)
    .map(i => [date, by, i.name, '', '', parseFloat(i.stock) || 0, 'baseline', '', ts]);

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  }

  return { success: true, count: rows.length };
}

// ============================================================
// PHOTO UPLOAD  (Google Drive)
// ============================================================

function uploadPhoto(type, date, itemName, filename, mimeType, base64Data) {
  try {
    const folderId = type === 'bar' ? BAR_DRIVE_FOLDER_ID : CONTAINER_DRIVE_FOLDER_ID;
    const root     = DriveApp.getFolderById(folderId);
    const dateDir  = findOrCreate(root, date);
    const itemDir  = findOrCreate(dateDir, itemName);

    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType, filename);
    const file    = itemDir.createFile(blob);

    // Try to make shareable — non-fatal if org policy restricts external sharing
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(shareErr) {
      Logger.log('setSharing skipped: ' + shareErr.message);
    }

    const fileId  = file.getId();
    const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
    return { success: true, fileId: fileId, fileUrl: fileUrl };
  } catch(e) {
    return err(e.message);
  }
}

// ============================================================
// SET STOCK  (writes a single adjust row to override one item's stock)
// ============================================================

function setStock(ssId, itemName, newStock, by) {
  const sheet = getLogSheet(ssId);
  const date  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const ts    = new Date().toISOString();
  sheet.appendRow([date, by, itemName, '', '', newStock, 'adjust', '', ts]);
  return { success: true, stock: newStock };
}

// ============================================================
// DELETE FOLDERS  (removes item-level Drive folders by path)
// Folder structure: root / dd-mm-yyyy / itemName
// ============================================================

function deleteFolders(type, date, itemNames) {
  try {
    const folderId = type === 'bar' ? BAR_DRIVE_FOLDER_ID : CONTAINER_DRIVE_FOLDER_ID;
    const root = DriveApp.getFolderById(folderId);

    // Convert yyyy-MM-dd → dd-mm-yyyy (folder naming convention used at upload)
    var parts = String(date).split('-');
    var folderDate = parts.length === 3
      ? parts[2] + '-' + parts[1] + '-' + parts[0]
      : date;

    // Find the date folder
    var dateFolders = root.getFoldersByName(folderDate);
    if (!dateFolders.hasNext()) {
      return { success: true, deleted: 0, msg: 'Date folder not found: ' + folderDate };
    }
    var dateFolder = dateFolders.next();

    var deleted = 0;
    itemNames.forEach(function(itemName) {
      var it = dateFolder.getFoldersByName(itemName);
      if (it.hasNext()) {
        it.next().setTrashed(true);
        deleted++;
      }
    });

    // If date folder is now empty, trash it too
    var remaining = dateFolder.getFolders();
    var hasFiles  = dateFolder.getFiles();
    if (!remaining.hasNext() && !hasFiles.hasNext()) {
      dateFolder.setTrashed(true);
    }

    return { success: true, deleted: deleted };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// DELETE PHOTOS  (removes Drive files by URL)
// ============================================================

function deletePhotos(fileUrls) {
  var deleted = 0, failed = 0;
  fileUrls.forEach(function(url) {
    try {
      // Extract file ID from Drive URL formats:
      //   https://drive.google.com/file/d/FILE_ID/view
      //   https://drive.google.com/open?id=FILE_ID
      var match = url.match(/\/d\/([^\/\?]+)/) || url.match(/[?&]id=([^&]+)/);
      if (!match) { failed++; return; }
      var fileId = match[1];
      DriveApp.getFileById(fileId).setTrashed(true);
      deleted++;
    } catch(e) {
      Logger.log('deletePhotos: failed for ' + url + ': ' + e.message);
      failed++;
    }
  });
  return { success: true, deleted: deleted, failed: failed };
}

function findOrCreate(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ============================================================
// RECENT LOG HISTORY  (returns grouped entry objects)
// ============================================================

function getRecentLog(ssId, limit) {
  const sheet = getLogSheet(ssId);
  if (sheet.getLastRow() < 2) return [];

  // Read ALL rows so we can accurately track running stock for count diffs
  const totalDataRows = sheet.getLastRow() - 1;
  const data = sheet.getRange(2, 1, totalDataRows, 9).getValues();

  const groups    = {}; // key → entry object
  const keys      = []; // ordered keys (chronological)
  const lastStock = {}; // tracks latest remaining per item as we iterate

  data.forEach(function(r) {
    const date   = r[0] instanceof Date
      ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(r[0]).trim();
    const by     = String(r[1]).trim();
    const item   = String(r[2]).trim();
    const recv   = parseFloat(r[3]) || 0;
    const used   = parseFloat(r[4]) || 0;
    const remain = parseFloat(r[5]) || 0;
    const kind   = String(r[6]).trim();
    const col8   = String(r[7]).trim();
    const ts     = String(r[8]).trim();

    if (!item) return;

    // Capture system stock BEFORE this row updates the tracker
    var prevStock = lastStock[item]; // undefined if first time seeing this item

    // Update running stock tracker for all row types
    lastStock[item] = remain;

    if (kind === 'baseline' || kind === 'adjust') return; // skip from history display

    var key = ts || (date + '\xA7' + by + '\xA7' + kind);

    if (!groups[key]) {
      groups[key] = {
        date:    date,
        by:      by,
        kind:    kind === 'count' ? 'physicalCount' : 'entry',
        label:   kind === 'count' ? col8 : '',
        entries: [],
        ts:      ts || key
      };
      keys.push(key);
    }

    if (kind === 'count') {
      // System stock = value tracked just before this count row
      // If recv col has a value, it was stored explicitly by addPhysicalCount (new format)
      var storedSys = parseFloat(r[3]);
      var sysStock  = !isNaN(storedSys) ? storedSys
                    : (prevStock !== undefined ? prevStock : remain);
      var diff = remain - sysStock;
      groups[key].entries.push({ name: item, system: sysStock, actual: remain, diff: diff });
    } else {
      var photos = col8 ? col8.split(',').map(function(p){ return p.trim(); }).filter(Boolean) : [];
      groups[key].entries.push({ name: item, recv: recv, used: used, photos: photos });
    }
  });

  // Return the most recent `limit` entries, newest first
  return keys.slice(-limit).reverse().map(function(k){ return groups[k]; });
}
