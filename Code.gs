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
// ── Kitchen (ครัว): create a NEW blank Google Sheet and paste its ID below ──
// (No Drive folder needed — kitchen entries have no photos)
const KITCHEN_SPREADSHEET_ID    = 'PASTE_KITCHEN_SHEET_ID_HERE';
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

      case 'getStock': {
        // Read each sheet once — catalog + log data — derive stock & history from same read
        var barCat   = getCatalog(BAR_SPREADSHEET_ID);
        var contCat  = getCatalog(CONTAINER_SPREADSHEET_ID);
        var barData  = getSheetData(BAR_SPREADSHEET_ID);
        var contData = getSheetData(CONTAINER_SPREADSHEET_ID);
        // Kitchen is optional — read defensively so a missing/placeholder ID
        // never breaks bar/container sync.
        var kitCat = [], kitData = { rows: [] };
        if (kitchenConfigured()) {
          try {
            kitCat  = getCatalog(KITCHEN_SPREADSHEET_ID);
            kitData = getSheetData(KITCHEN_SPREADSHEET_ID);
          } catch (kitErr) { /* leave kitchen empty */ }
        }
        return {
          success:          true,
          bar:              { items: deriveStock(barData.rows,  barCat)  },
          container:        { items: deriveStock(contData.rows, contCat) },
          kitchen:          { items: deriveStock(kitData.rows,  kitCat)  },
          barCatalog:       barCat,
          containerCatalog: contCat,
          kitchenCatalog:   kitCat,
          kitchenReady:     kitchenConfigured(),
          barHistory:       deriveHistory(barData.rows,  80),
          containerHistory: deriveHistory(contData.rows, 80),
          kitchenHistory:   deriveHistory(kitData.rows,  80),
          lastUpdated:      new Date().toISOString()
        };
      }

      case 'addEntry': {
        const { type, date, by, recv, used } = body;
        if (!type || !date) return err('Missing type or date');
        const ssId = ssFor(type);
        return addEntry(ssId, date, by || '', recv || {}, used || {});
      }

      case 'kitchenEntry': {
        // Kitchen: client sends { name, recv, remaining } per item.
        // remaining (คงเหลือ) is authoritative; used (ใช้ไป) is computed server-side.
        const { date, by, entries } = body;
        if (!date || !entries) return err('Missing date or entries');
        return addKitchenEntry(KITCHEN_SPREADSHEET_ID, date, by || '', entries);
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

      case 'deleteLogEntry': {
        // Delete all Log rows matching a given timestamp (removes a count/entry session)
        const { type, ts } = body;
        if (!type || !ts) return err('Missing type or ts');
        return deleteLogEntry(ssFor(type), String(ts));
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
  if (type === 'bar')     return BAR_SPREADSHEET_ID;
  if (type === 'kitchen') return KITCHEN_SPREADSHEET_ID;
  return CONTAINER_SPREADSHEET_ID;
}

// True only once a real kitchen spreadsheet ID has been pasted in above.
function kitchenConfigured() {
  return !!KITCHEN_SPREADSHEET_ID && KITCHEN_SPREADSHEET_ID.indexOf('PASTE_') !== 0;
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
// SINGLE-READ DATA LAYER
// Reads the Log sheet ONCE and returns raw rows.
// deriveStock + deriveHistory both operate on the same rows.
// ============================================================

// How many rows to read for history (recent entries only).
// Stock derivation always scans all rows for accuracy.
var HISTORY_ROW_LIMIT = 500;

function getSheetData(ssId) {
  var sheet = getLogSheet(ssId);
  if (sheet.getLastRow() < 2) return { rows: [] };
  var totalRows = sheet.getLastRow() - 1;
  var rows = sheet.getRange(2, 1, totalRows, 9).getValues();
  return { rows: rows };
}

// Derive current stock levels from all rows (full scan — required for accuracy)
function deriveStock(rows, catalog) {
  var latest = {};
  rows.forEach(function(r) {
    var item = String(r[2]).trim();
    if (!item) return;
    var kind = String(r[6]).trim();
    if (kind === 'entry' || kind === 'count' || kind === 'baseline' || kind === 'adjust') {
      latest[item] = parseFloat(r[5]) || 0;
    }
  });
  return catalog.map(function(c) {
    return { name: c.name, unit: c.unit, stock: latest[c.name] !== undefined ? latest[c.name] : 0 };
  });
}

// Derive history entries from the last HISTORY_ROW_LIMIT rows only
// System stock for count rows is read from the recv col (stored by addPhysicalCount)
function deriveHistory(rows, limit) {
  // Only process recent rows for history display
  var recentRows = rows.length > HISTORY_ROW_LIMIT ? rows.slice(-HISTORY_ROW_LIMIT) : rows;

  // We still need running stock at the START of our window for count diffs
  // Build lastStock from ALL rows up to where recentRows begins
  var lastStock = {};
  if (rows.length > HISTORY_ROW_LIMIT) {
    var priorRows = rows.slice(0, rows.length - HISTORY_ROW_LIMIT);
    priorRows.forEach(function(r) {
      var item = String(r[2]).trim();
      if (!item) return;
      lastStock[item] = parseFloat(r[5]) || 0;
    });
  }

  var groups = {};
  var keys   = [];

  recentRows.forEach(function(r) {
    var date   = r[0] instanceof Date
      ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(r[0]).trim();
    var by     = String(r[1]).trim();
    var item   = String(r[2]).trim();
    var recv   = parseFloat(r[3]) || 0;
    var used   = parseFloat(r[4]) || 0;
    var remain = parseFloat(r[5]) || 0;
    var kind   = String(r[6]).trim();
    var col8   = String(r[7]).trim();
    var ts     = String(r[8]).trim();

    if (!item) return;

    var prevStock = lastStock[item];
    lastStock[item] = remain;

    if (kind === 'baseline' || kind === 'adjust') return;

    var key = ts || (date + '\xA7' + by + '\xA7' + kind);
    if (!groups[key]) {
      groups[key] = { date: date, by: by, kind: kind === 'count' ? 'physicalCount' : 'entry',
        label: kind === 'count' ? col8 : '', entries: [], ts: ts || key };
      keys.push(key);
    }

    if (kind === 'count') {
      var storedSys = parseFloat(r[3]);
      var sysStock  = !isNaN(storedSys) ? storedSys : (prevStock !== undefined ? prevStock : remain);
      groups[key].entries.push({ name: item, system: sysStock, actual: remain, diff: remain - sysStock });
    } else {
      var photos = col8 ? col8.split(',').map(function(p){ return p.trim(); }).filter(Boolean) : [];
      groups[key].entries.push({ name: item, recv: recv, used: used, photos: photos });
    }
  });

  return keys.slice(-limit).reverse().map(function(k){ return groups[k]; });
}

// ============================================================
// CURRENT STOCK  (kept for internal use by other actions)
// ============================================================

function getCurrentStock(ssId) {
  var catalog = getCatalog(ssId);
  var data    = getSheetData(ssId);
  return { items: deriveStock(data.rows, catalog) };
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
// KITCHEN ENTRY  (remaining/คงเหลือ is authoritative; used/ใช้ไป computed)
// entries = [{ name, recv, remaining }, ...]
// Every kitchen entry is effectively a physical count, so we trust the
// counted "remaining" and back-calculate usage: used = prev + recv − remaining.
// ============================================================

function addKitchenEntry(ssId, date, by, entries) {
  if (!kitchenConfigured()) return err('Kitchen sheet not configured');
  const catalog = getCatalog(ssId);
  if (!catalog.length) return err('Catalog is empty — add items first');

  // Current stock map (server-side truth) for computing used
  const stockData = getCurrentStock(ssId);
  const stockMap  = {};
  stockData.items.forEach(i => { stockMap[i.name] = i.stock; });

  const sheet = getLogSheet(ssId);
  const ts    = new Date().toISOString();
  const rows  = [];

  entries.forEach(e => {
    const name = String(e.name || '').trim();
    if (!name) return;
    const recv = parseFloat(e.recv) || 0;
    const remaining = parseFloat(e.remaining);
    if (isNaN(remaining)) return;                 // a count is required
    const prev = stockMap[name] || 0;
    let used = prev + recv - remaining;
    if (used < 0) used = 0;                        // counted more than expected — don't log negative usage
    stockMap[name] = remaining;                    // remaining is authoritative
    rows.push([date, by, name, recv || '', used || '', remaining, 'entry', '', ts]);
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
// DELETE LOG ENTRY  (removes all rows matching a timestamp from the Log tab)
// ============================================================

function deleteLogEntry(ssId, ts) {
  var sheet = getLogSheet(ssId);
  if (sheet.getLastRow() < 2) return { success: true, deleted: 0 };

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  // Collect row numbers that match the ts (column index 8), in reverse order
  var rowsToDelete = [];
  data.forEach(function(r, i) {
    if (String(r[8]).trim() === ts) {
      rowsToDelete.push(i + 2); // +2: 1-indexed + header row
    }
  });

  // Delete from bottom to top so row indices stay valid
  rowsToDelete.reverse().forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });

  return { success: true, deleted: rowsToDelete.length };
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


