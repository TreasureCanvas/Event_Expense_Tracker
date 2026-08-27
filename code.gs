/**
 * ===================================================================
 *  ระบบจัดทำและรวบรวมเอกสารเบิกค่าเดินทางเชื่อมโยงกิจกรรม
 *  (Event Expense Tracker) - Backend (Google Apps Script)
 * ===================================================================
 */

// ---------- CONFIG ----------
const SHEET_EVENTS = 'Events';
const SHEET_SUBMISSIONS = 'Submissions';
const SHEET_EVENT_EMPLOYEES = 'EventEmployees';
const SHEET_EVENT_FUEL_PRICES = 'EventFuelPrices';
const DRIVE_FOLDER_NAME = 'EventExpenseTracker_Files';
const FUEL_RECEIPTS_SUBFOLDER_NAME = 'Fuel_Receipts';

// ---------- SETUP ----------
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function ensureSheet(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return sh;
  }

  ensureSheet(SHEET_EVENTS, [
    'event_id', 'event_name', 'owner_name', 'venue', 'start_date', 'end_date', 'form_link', 'event_token', 'created_at'
  ]);

  ensureSheet(SHEET_SUBMISSIONS, [
    'submission_id', 'event_id', 'employee_name', 'department',
    'submission_status', 'updated_at', 'total_amount',
    'summary_text', 'pdf_file_url', 'trip_details_json'
  ]);

  ensureSheet(SHEET_EVENT_EMPLOYEES, [
    'event_id', 'employee_name', 'department'
  ]);

  ensureSheet(SHEET_EVENT_FUEL_PRICES, [
    'event_id', 'date', 'price', 'receipt_url'
  ]);
  // Force the 'date' column to stay plain text (YYYY-MM-DD) — otherwise
  // Sheets auto-converts date-looking strings into Date serials, which
  // breaks the simple lexicographic date comparisons used for the
  // "closest previous date" fuel-price fallback in form.html.
  const fpSheet = ss.getSheetByName(SHEET_EVENT_FUEL_PRICES);
  fpSheet.getRange('B2:B').setNumberFormat('@');

  getOrCreateDriveFolder();
  Logger.log('Setup complete.');
}

function getOrCreateDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ---------- Per-event Drive folder structure ----------
// Root ("EventExpenseTracker_Files")
//   └── <event folder, e.g. "EV-2026-ABCD1234 - อบรมการตลาด">   <- travel receipts + PDFs live here (unchanged location)
//         └── "Fuel_Receipts"                                    <- NEW: daily fuel price receipt images live here, kept
//                                                                    separate from travel receipt images
function getOrCreateSubfolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

function getOrCreateEventFolder(eventId, eventName) {
  const root = getOrCreateDriveFolder();
  const safeEventName = String(eventName || '').replace(/[\\/:*?"<>|]/g, '_').trim();
  const folderName = safeEventName ? (eventId + ' - ' + safeEventName) : eventId;
  return getOrCreateSubfolder_(root, folderName);
}

function getOrCreateFuelReceiptsFolder(eventId, eventName) {
  const eventFolder = getOrCreateEventFolder(eventId, eventName);
  return getOrCreateSubfolder_(eventFolder, FUEL_RECEIPTS_SUBFOLDER_NAME);
}

// ---------- SAFE SCHEMA MIGRATION ----------
function migrateEventsSchema() {
  const sheet = getSheet(SHEET_EVENTS);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];

  function ensureColumnBefore(headerName, beforeHeaderName) {
    if (headers.indexOf(headerName) > -1) return;
    let insertAt = headers.indexOf(beforeHeaderName) + 1;
    if (insertAt === 0) insertAt = headers.length + 1;
    sheet.insertColumnBefore(insertAt);
    sheet.getRange(1, insertAt).setValue(headerName);
    headers.splice(insertAt - 1, 0, headerName);
  }

  ensureColumnBefore('venue', 'start_date');
  ensureColumnBefore('form_link', 'event_token');

  const data = sheet.getDataRange().getValues();
  const newHeaders = data[0];
  const idCol = newHeaders.indexOf('event_id');
  const linkCol = newHeaders.indexOf('form_link');
  const scriptUrl = ScriptApp.getService().getUrl();

  for (let r = 1; r < data.length; r++) {
    if (!data[r][linkCol] && data[r][idCol]) {
      sheet.getRange(r + 1, linkCol + 1).setValue(scriptUrl + '?page=form&event=' + data[r][idCol]);
    }
  }

  Logger.log('Migration complete.');
}

// ---------- WEB ENTRY POINTS ----------
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'home';
  let tmpl;
  if (page === 'form') {
    tmpl = HtmlService.createTemplateFromFile('form'); // แก้จาก 'index' เป็น 'form'
  } else {
    tmpl = HtmlService.createTemplateFromFile('index'); // แก้จาก 'home' เป็น 'index'
  }
  tmpl.eventId = (e && e.parameter && e.parameter.event) || '';
  return tmpl.evaluate()
    .setTitle('Event Expense Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// รองรับ CORS Preflight Requests จาก Vercel
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// doPost — NO LockService ANYWHERE (temporarily removed to eliminate it as a suspect entirely)
function doPost(e) {
  let body = {};
  if (e && e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonOut({ ok: false, error: 'รูปแบบข้อมูล JSON ไม่ถูกต้อง' });
    }
  } else if (e && e.parameter && e.parameter.payload) {
    try {
      body = JSON.parse(e.parameter.payload);
    } catch (pErr) {
      body = e.parameter;
    }
  } else if (e && e.parameter) {
    body = e.parameter;
  }

  const action = body.action;
  try {
    return jsonOut(routeAction(action, body.payload));
  } catch (err) {
    return jsonOut({ ok: false, error: 'Internal Server Error: ' + String(err) });
  }
}

function routeAction(action, payload) {
  switch (action) {
    case 'createEvent':
      return apiCreateEvent(payload);
    case 'updateEvent':
      return apiUpdateEvent(payload);
    case 'listAllEvents':
      return apiListAllEvents();
    case 'getEventInfo':
      return apiGetEventInfo(payload);
    case 'getEventRoster':
      return apiGetEventRoster(payload);
    case 'submitExpense':
      return apiSubmitExpense(payload);
    case 'getMySubmission':
      return apiGetMySubmission(payload);
    case 'getFuelPrice':
      return apiGetFuelPrice();
    case 'getFileBase64':
      return apiGetFileAsBase64(payload);
    case 'getFuelReceiptImages':
      return apiGetFuelReceiptImages(payload);
    default:
      return { ok: false, error: 'ไม่พบ Action ที่ระบุ: ' + action };
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- HELPERS ----------
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values.shift();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ---------- Response caching (CacheService) ----------
// Read-heavy endpoints (event list, event info, roster) re-scan whole
// sheets on every call. When several people hit the same event in quick
// succession (e.g. multiple staff opening the same form link, or an admin
// refreshing the event list), a short-lived cache lets every call after
// the first be served instantly with zero Sheets reads. TTLs are kept
// short (30–60s) so nothing stays stale for long, and every mutation below
// explicitly clears the relevant key(s) so writes are always reflected
// immediately rather than waiting for the cache to expire.
function cacheGetJson_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    return null; // cache miss/unavailable — caller just falls back to a real read
  }
}
function cacheSetJson_(key, obj, ttlSeconds) {
  try {
    // CacheService values are capped at 100KB — silently skip caching
    // (rather than fail the request) if a response is ever too large.
    CacheService.getScriptCache().put(key, JSON.stringify(obj), ttlSeconds);
  } catch (err) {
    // too large or cache temporarily unavailable — fine, just not cached
  }
}
function cacheRemove_(key) {
  try {
    CacheService.getScriptCache().remove(key);
  } catch (err) { /* no-op */ }
}

function generateId(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function nowStr() {
  return Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy - HH:mm');
}

function dateStr(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(d, 'GMT+7', 'dd/MM/yyyy');
}

// Normalizes any date value (plain 'YYYY-MM-DD' string, a Date object that
// Sheets may have auto-parsed despite the '@' text format, or anything else)
// into a strict 'YYYY-MM-DD' string. Used for the fuel-price table so dates
// can be compared/matched with simple lexicographic string comparisons.
function isoDateStr(d) {
  if (!d) return '';
  if (d instanceof Date) {
    if (isNaN(d)) return '';
    return Utilities.formatDate(d, 'GMT+7', 'yyyy-MM-dd');
  }
  const str = String(d).trim();
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  const parsed = new Date(str);
  if (!isNaN(parsed)) return Utilities.formatDate(parsed, 'GMT+7', 'yyyy-MM-dd');
  return str;
}

function safeSetPublicSharing(file) {
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log('Warning: Cannot set public file sharing policy: ' + err.toString());
  }
}

// ---------- Drive file <-> Base64 helpers (for edit-mode receipt preview) ----------
// Extracts the Drive file ID out of whatever URL shape we stored
// (file.getUrl() typically returns .../file/d/<ID>/view?usp=drivesdk).
function driveFileIdFromUrl_(url) {
  if (!url) return null;
  const m = /\/file\/d\/([a-zA-Z0-9_-]+)/.exec(url) || /[?&]id=([a-zA-Z0-9_-]+)/.exec(url);
  return m ? m[1] : null;
}

// Reads a previously-saved receipt image straight out of Drive on the
// server side (no CORS, no client-side fetch of a Google login-gated page)
// and returns it as a ready-to-embed Base64 data URL, in the exact same
// shape the client uses for freshly-picked files ({ base64, fileName,
// mimeType }). Returns null if the file can't be read (deleted, no access,
// bad URL, etc.) so the caller can skip it gracefully instead of failing
// the whole request.
function driveFileToBase64Attachment_(url) {
  const fileId = driveFileIdFromUrl_(url);
  if (!fileId) return null;
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType() || 'image/jpeg';
    const base64 = 'data:' + mimeType + ';base64,' + Utilities.base64Encode(blob.getBytes());
    return {
      base64: base64,
      fileName: file.getName(),
      mimeType: mimeType,
      isExisting: true,   // flags this as an already-uploaded receipt (see apiSubmitExpense)
      existingUrl: url     // keeps the original Drive URL so we can skip re-uploading it
    };
  } catch (err) {
    Logger.log('Existing receipt fetch error for ' + url + ': ' + err.toString());
    return null;
  }
}

// ---------- Daily fuel price + receipt persistence ----------
// payload.fuelPrices shape (per day): {
//   date: 'YYYY-MM-DD',
//   price: number,
//   receipt: null
//        | { isExisting:true, existingUrl:'<drive url>' }   <- previously uploaded, keep as-is (never re-uploaded)
//        | { base64:'data:image/...;base64,...', fileName, mimeType }  <- freshly picked file to upload
// }
// Mirrors the isExisting/existingUrl pattern already used for travel-
// receipt attachments in apiSubmitExpense, so editing an event never
// silently deletes or re-encodes a fuel receipt that wasn't touched.
function saveFuelPrices_(eventId, eventName, fuelPrices) {
  const sheet = getSheet(SHEET_EVENT_FUEL_PRICES);
  if (!sheet) return;

  // Replace this event's price rows wholesale — safe because every row we
  // write below either keeps the receipt's existing Drive URL untouched or
  // uploads a fresh one; nothing is ever silently dropped unless the client
  // omitted that date/receipt on purpose (i.e. the user removed it in the UI).
  //
  // PERFORMANCE: previously this issued one deleteRow() call per existing
  // row PLUS one appendRow() call per new row — each one a separate,
  // relatively slow Sheets API round-trip. Instead, read the whole sheet
  // once, build the final data set in memory (every other event's rows +
  // this event's new rows), and write it all back with a single batched
  // setValues() call.
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('event_id');
  const otherRows = data.slice(1).filter(row => row[idCol] !== eventId);

  const newRows = [];
  if (Array.isArray(fuelPrices) && fuelPrices.length > 0) {
    let fuelFolder = null; // lazily created — only if we actually need to upload a new receipt

    fuelPrices.forEach(fp => {
      if (!fp || !fp.date) return;
      let receiptUrl = '';

      if (fp.receipt) {
        if (fp.receipt.isExisting && fp.receipt.existingUrl) {
          receiptUrl = fp.receipt.existingUrl;
        } else if (fp.receipt.base64) {
          try {
            if (!fuelFolder) fuelFolder = getOrCreateFuelReceiptsFolder(eventId, eventName);
            const match = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/.exec(fp.receipt.base64);
            const mime = match ? match[1] : 'image/jpeg';
            const dataPart = match ? match[2] : fp.receipt.base64;
            const ext = (mime.split('/')[1] || 'jpg');
            const blob = Utilities.newBlob(
              Utilities.base64Decode(dataPart), mime,
              'fuel_' + eventId + '_' + fp.date + '.' + ext
            );
            const file = fuelFolder.createFile(blob);
            safeSetPublicSharing(file);
            receiptUrl = file.getUrl();
          } catch (err) {
            Logger.log('Fuel receipt save error (' + fp.date + '): ' + err.toString());
          }
        }
      }

      newRows.push([eventId, String(fp.date), Math.round((Number(fp.price) || 0) * 100) / 100, receiptUrl]);
    });
  }

  const finalRows = otherRows.concat(newRows);
  // Clear just the data area (keep the header row + its formatting intact)
  // then write everything back in one shot.
  const numCols = headers.length;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, numCols).setValues(finalRows);
  }
}

// Fast path — plain sheet read only, NO Drive calls. Used by
// apiGetEventInfo, which fires on every single form.html/edit-event page
// load. Each receipt image previously required its own synchronous
// DriveApp.getFileById()+getBlob() round-trip *every time this ran* even
// though the image itself is only actually needed once — at PDF-generation
// time (form.html) or when an admin explicitly clicks "🔍 ดูรูปภาพ"
// (index.html). Skipping that here is the single biggest speed win
// available, since it turns "load the form" from N Drive fetches down to
// zero.
function getEventFuelPrices_(eventId) {
  const rows = sheetToObjects(getSheet(SHEET_EVENT_FUEL_PRICES)).filter(x => x.event_id === eventId);
  return rows.map(fp => {
    const receiptUrl = fp.receipt_url || '';
    return {
      date: isoDateStr(fp.date),
      price: Math.round((Number(fp.price) || 0) * 100) / 100,
      receiptUrl: receiptUrl,
      receiptImage: null // fetched on demand — see getEventFuelPricesWithImages_
    };
  }).sort((a, b) => a.date.localeCompare(b.date));
}

// Slow path — same data as above, but with each receipt image resolved to
// Base64 via a Drive round-trip. Only called on demand: right before
// form.html generates a PDF that actually needs a fuel receipt embedded,
// or when index.html's edit-event view loads receipt previews in the
// background after the fast fields have already rendered.
function getEventFuelPricesWithImages_(eventId) {
  return getEventFuelPrices_(eventId).map(fp => Object.assign({}, fp, {
    receiptImage: fp.receiptUrl ? driveFileToBase64Attachment_(fp.receiptUrl) : null
  }));
}

// ---------- API: สร้างกิจกรรมใหม่ ----------
function apiCreateEvent(payload) {
  if (!payload || !payload.eventName || !payload.ownerName) {
    return { ok: false, error: 'กรุณากรอกชื่อกิจกรรมและชื่อผู้สร้าง' };
  }

  const events = getSheet(SHEET_EVENTS);
  const eventId = 'EV-' + new Date().getFullYear() + '-' + generateId('').replace('-', '');

  // ✅ แก้ไขลิงก์ให้ชี้ไปที่ Vercel URL
  const vercelBaseUrl = 'https://event-expense-tracker-chi.vercel.app';
  const formLink = vercelBaseUrl + '/form.html?event=' + eventId;

  events.appendRow([
    eventId, 
    String(payload.eventName).trim(), 
    String(payload.ownerName).trim(), 
    String(payload.venue || '').trim(),
    payload.startDate || '', 
    payload.endDate || payload.startDate || '',
    formLink, 
    '', 
    nowStr()
  ]);

  const empSheet = getSheet(SHEET_EVENT_EMPLOYEES);
  if (Array.isArray(payload.employees) && payload.employees.length > 0) {
    // Batch write: build the full set of employee rows in memory and write
    // them in one setValues() call instead of one appendRow() per person.
    const empRows = payload.employees
      .filter(emp => emp && emp.name)
      .map(emp => [eventId, String(emp.name).trim(), String(emp.department || '').trim()]);
    if (empRows.length > 0) {
      const startRow = empSheet.getLastRow() + 1;
      empSheet.getRange(startRow, 1, empRows.length, 3).setValues(empRows);
    }
  }

  saveFuelPrices_(eventId, payload.eventName, payload.fuelPrices);

  cacheRemove_(CACHE_KEY_EVENTS_LIST);

  return {
    ok: true,
    eventId: eventId,
    formLink: formLink
  };
}

// ---------- API: แก้ไขข้อมูลกิจกรรม ----------
function apiUpdateEvent(payload) {
  if (!payload || !payload.eventId) return { ok: false, error: 'ไม่พบรหัสกิจกรรมที่จะแก้ไข' };
  if (!payload.eventName || !payload.ownerName) {
    return { ok: false, error: 'กรุณากรอกชื่อกิจกรรมและชื่อผู้สร้าง' };
  }

  const events = getSheet(SHEET_EVENTS);
  const data = events.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('event_id');
  let rowIndex = -1;
  
  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === payload.eventId) { rowIndex = r; break; }
  }
  if (rowIndex === -1) return { ok: false, error: 'ไม่พบกิจกรรมนี้ในระบบ' };

  const nameCol = headers.indexOf('event_name');
  const ownerCol = headers.indexOf('owner_name');
  const venueCol = headers.indexOf('venue');
  const startCol = headers.indexOf('start_date');
  const endCol = headers.indexOf('end_date');

  // Batch write: edit the row's values in memory (preserving every other
  // column untouched), then write the whole row back with a single
  // setValues() call instead of up to 5 separate setValue() round-trips.
  const updatedRow = data[rowIndex].slice();
  updatedRow[nameCol] = String(payload.eventName).trim();
  updatedRow[ownerCol] = String(payload.ownerName).trim();
  if (venueCol > -1) updatedRow[venueCol] = String(payload.venue || '').trim();
  updatedRow[startCol] = payload.startDate || '';
  updatedRow[endCol] = payload.endDate || payload.startDate || '';
  events.getRange(rowIndex + 1, 1, 1, updatedRow.length).setValues([updatedRow]);

  // ลบรายชื่อพนักงานเดิม แล้วอัปเดตชุดใหม่ลงไป
  // PERFORMANCE: same batching pattern as saveFuelPrices_ — read once,
  // filter out this event's old rows in memory, append the new set, and
  // write everything back with a single setValues() call instead of one
  // deleteRow()/appendRow() per employee.
  const empSheet = getSheet(SHEET_EVENT_EMPLOYEES);
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData[0];
  const empIdCol = empHeaders.indexOf('event_id');
  const otherEmpRows = empData.slice(1).filter(row => row[empIdCol] !== payload.eventId);

  const newEmpRows = Array.isArray(payload.employees)
    ? payload.employees
        .filter(emp => emp && emp.name)
        .map(emp => [payload.eventId, String(emp.name).trim(), String(emp.department || '').trim()])
    : [];

  const finalEmpRows = otherEmpRows.concat(newEmpRows);
  const empLastRow = empSheet.getLastRow();
  if (empLastRow > 1) empSheet.getRange(2, 1, empLastRow - 1, empHeaders.length).clearContent();
  if (finalEmpRows.length > 0) {
    empSheet.getRange(2, 1, finalEmpRows.length, empHeaders.length).setValues(finalEmpRows);
  }

  // Photo persistence guarantee: saveFuelPrices_ only re-uploads receipts
  // flagged as freshly-picked (fp.receipt.base64); anything flagged
  // isExisting/existingUrl keeps pointing at its original Drive file, so
  // editing the event never drops previously uploaded fuel receipts unless
  // the client omitted them on purpose (user removed them in the UI).
  saveFuelPrices_(payload.eventId, payload.eventName, payload.fuelPrices);

  cacheRemove_(CACHE_KEY_EVENTS_LIST);
  cacheRemove_('event_info_' + payload.eventId);
  cacheRemove_('roster_' + payload.eventId);

  return { ok: true, eventId: payload.eventId };
}

// ---------- API: รายการกิจกรรมทั้งหมด ----------
const CACHE_KEY_EVENTS_LIST = 'events_list_v1';

function apiListAllEvents() {
  const cached = cacheGetJson_(CACHE_KEY_EVENTS_LIST);
  if (cached) return cached;

  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const employees = sheetToObjects(getSheet(SHEET_EVENT_EMPLOYEES));

  // Count staff per event in a single pass instead of re-filtering the
  // full employee list once per event (O(n) instead of O(events × employees),
  // which matters once an org has built up many events/staff over time).
  const staffCountByEvent = {};
  employees.forEach(e => {
    staffCountByEvent[e.event_id] = (staffCountByEvent[e.event_id] || 0) + 1;
  });

  const list = events.map(ev => {
    return {
      eventId: ev.event_id,
      eventName: ev.event_name,
      ownerName: ev.owner_name,
      venue: ev.venue || '',
      startDate: dateStr(ev.start_date),
      endDate: dateStr(ev.end_date),
      formLink: ev.form_link || (ScriptApp.getService().getUrl() + '?page=form&event=' + ev.event_id),
      staffCount: staffCountByEvent[ev.event_id] || 0,
      createdAt: ev.created_at
    };
  }).sort((a, b) => (new Date(b.createdAt || 0) - new Date(a.createdAt || 0)));

  const result = { ok: true, events: list };
  cacheSetJson_(CACHE_KEY_EVENTS_LIST, result, 45);
  return result;
}

// ---------- API: ดึงข้อมูลกิจกรรมและรายชื่อพนักงานสำหรับ Auto-complete ----------
function apiGetEventInfo(payload) {
  if (!payload || !payload.eventId) return { ok: false, error: 'ไม่ได้ระบุรหัสกิจกรรม' };

  const eventId = payload.eventId;
  const cacheKey = 'event_info_' + eventId;
  const cached = cacheGetJson_(cacheKey);
  if (cached) return cached;

  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้ในระบบ (Event ID ไม่ถูกต้อง)' };

  const employees = sheetToObjects(getSheet(SHEET_EVENT_EMPLOYEES))
    .filter(x => x.event_id === eventId)
    .map(x => ({ name: x.employee_name, department: x.department }));

  const result = {
    ok: true,
    event: {
      eventId: ev.event_id, eventName: ev.event_name, ownerName: ev.owner_name, venue: ev.venue || '',
      startDate: dateStr(ev.start_date), endDate: dateStr(ev.end_date)
    },
    employees: employees,
    // Per-day Gasohol 91 price (receipt IMAGES are intentionally excluded
    // here — see getEventFuelPrices_ — and fetched on demand instead) —
    // used by index.html to pre-populate the edit form, and by form.html
    // for the personal-vehicle auto-calculation.
    fuelPrices: getEventFuelPrices_(eventId)
  };
  // This endpoint fires on every single form.html page load, so caching it
  // briefly means every concurrent staff member opening the same event's
  // form link gets served instantly instead of re-scanning both sheets.
  cacheSetJson_(cacheKey, result, 30);
  return result;
}

// ---------- API: ดึงสถานะการส่งเอกสารของพนักงานทุกคนในกิจกรรม ----------
function apiGetEventRoster(payload) {
  if (!payload || !payload.eventId) return { ok: false, error: 'ไม่ได้ระบุรหัสกิจกรรม' };

  const eventId = payload.eventId;
  const cacheKey = 'roster_' + eventId;
  const cached = cacheGetJson_(cacheKey);
  if (cached) return cached;

  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้' };

  const roster = sheetToObjects(getSheet(SHEET_EVENT_EMPLOYEES)).filter(x => x.event_id === eventId);
  const submissions = sheetToObjects(getSheet(SHEET_SUBMISSIONS)).filter(s => s.event_id === eventId);

  // จำนวนวันที่เบิก (Claimed Days Count) — number of unique travel dates
  // found across that staff member's submitted trip details.
  function countClaimedDays_(sub) {
    if (!sub || !sub.trip_details_json) return 0;
    try {
      const details = JSON.parse(sub.trip_details_json);
      const dates = new Set();
      (details.trips || []).forEach(t => { if (t.date) dates.add(String(t.date)); });
      return dates.size;
    } catch (err) {
      return 0;
    }
  }

  // Structured per-leg trip data (date, ขาไป/ขากลับ label, segments with
  // travel mode / origin / destination / amount) — used by index.html to
  // build the "Copy Excel Template Summary" text without having to
  // re-parse or re-fetch anything else.
  function getTripDetails_(sub) {
    if (!sub || !sub.trip_details_json) return [];
    try {
      const details = JSON.parse(sub.trip_details_json);
      return details.trips || [];
    } catch (err) {
      return [];
    }
  }

  const rows = roster.map(emp => {
    const sub = submissions.find(s => String(s.employee_name).trim() === String(emp.employee_name).trim());
    if (!sub) {
      return {
        employeeName: emp.employee_name, department: emp.department,
        submitted: false, status: 'ยังไม่ส่งเอกสาร', updatedAt: '', totalAmount: 0,
        pdfUrl: '', summaryText: '', claimedDays: 0, tripDetails: []
      };
    }
    return {
      employeeName: emp.employee_name, department: emp.department,
      submitted: true, status: sub.submission_status, updatedAt: sub.updated_at,
      totalAmount: Math.round((Number(sub.total_amount) || 0) * 100) / 100, pdfUrl: sub.pdf_file_url, summaryText: sub.summary_text,
      claimedDays: countClaimedDays_(sub), tripDetails: getTripDetails_(sub)
    };
  });

  const result = {
    ok: true,
    event: {
      eventId: ev.event_id, eventName: ev.event_name, ownerName: ev.owner_name, venue: ev.venue || '',
      startDate: dateStr(ev.start_date), endDate: dateStr(ev.end_date)
    },
    roster: rows
  };
  // Short TTL — long enough to absorb a burst of admins/browsers refreshing
  // the same event's roster panel, short enough that a fresh submission
  // never stays hidden for long even if the explicit cacheRemove_ below
  // were ever missed.
  cacheSetJson_(cacheKey, result, 20);
  return result;
}

// ---------- API: บันทึก / แก้ไข การส่งเอกสารเบิกจ่าย ----------
function apiSubmitExpense(payload) {
  if (!payload || !payload.eventId || !payload.employeeName) {
    return { ok: false, error: 'ข้อมูลไม่ครบถ้วน (ต้องระบุ eventId และ employeeName)' };
  }

  const submissions = getSheet(SHEET_SUBMISSIONS);
  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === payload.eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้ในระบบ' };

  // Travel-receipt images & the generated PDF live directly inside this
  // event's own folder (kept apart from daily fuel-price receipts, which
  // are filed under that event folder's "Fuel_Receipts" subfolder instead).
  const folder = getOrCreateEventFolder(payload.eventId, ev.event_name);

  // 1. บันทึกไฟล์ PDF เอกสารเบิกเงินลง Drive (ทำก่อนล็อก เพราะ Drive I/O ใช้เวลานาน)
  let pdfUrl = '';
  if (payload.pdfBase64) {
    try {
      const cleanBase64 = payload.pdfBase64.replace(/^data:application\/pdf;base64,/, '');
      const blob = Utilities.newBlob(
        Utilities.base64Decode(cleanBase64),
        'application/pdf',
        payload.pdfFileName || ('Expense_' + payload.employeeName + '.pdf')
      );
      const file = folder.createFile(blob);
      safeSetPublicSharing(file);
      pdfUrl = file.getUrl();
    } catch (pdfErr) {
      Logger.log('PDF Save Error: ' + pdfErr.toString());
    }
  }

  // 2. บันทึกรูปสลิปแนบลง Drive
  //    seg.attachments now carries BOTH freshly-picked files AND previously
  //    submitted receipts that were converted back to Base64 for preview
  //    (flagged isExisting + existingUrl by apiGetMySubmission /
  //    driveFileToBase64Attachment_). We only need to actually re-upload the
  //    genuinely new ones — existing receipts just get their original URL
  //    kept as-is, so editing/resubmitting a report doesn't create
  //    duplicate copies of the same image in Drive every time.
  const tripDetails = (payload.tripDetails || []).map((trip, ti) => {
    const segments = (trip.segments || []).map((seg, si) => {
      const out = Object.assign({}, seg);
      const keptUrls = Array.isArray(seg.attachmentUrls) ? seg.attachmentUrls.slice() : [];
      const allAttachments = Array.isArray(seg.attachments) ? seg.attachments : [];

      allAttachments.forEach((att, ai) => {
        if (!att) return;

        // Already-uploaded receipt (came back from apiGetMySubmission as
        // Base64 purely for preview/PDF purposes) — reuse its Drive URL
        // instead of re-encoding and re-uploading the same image.
        if (att.isExisting && att.existingUrl) {
          keptUrls.push(att.existingUrl);
          return;
        }

        if (!att.base64) return;
        try {
          const match = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/.exec(att.base64);
          const mime = match ? match[1] : 'image/jpeg';
          const data = match ? match[2] : att.base64;
          const ext = (mime.split('/')[1] || 'jpg');
          const imgBlob = Utilities.newBlob(
            Utilities.base64Decode(data), mime,
            'slip_' + payload.employeeName + '_trip' + (ti + 1) + '_seg' + (si + 1) + '_' + (ai + 1) + '.' + ext
          );
          const imgFile = folder.createFile(imgBlob);
          safeSetPublicSharing(imgFile);
          keptUrls.push(imgFile.getUrl());
        } catch (imgErr) {
          Logger.log('Image Save Error: ' + imgErr.toString());
        }
      });

      out.attachmentUrls = keptUrls;
      delete out.attachments;
      delete out.existingAttachments; // preview-only field, never persisted
      return out;
    });
    return Object.assign({}, trip, { segments: segments });
  });

  const storedDetails = {
    purpose: payload.purpose || '',
    fuelRefPrice: payload.fuelRefPrice || '',
    trips: tripDetails
  };

  // ตรวจสอบการส่งซ้ำ + เขียนแถวลง Sheet — NO LockService (temporarily removed)
  const all = sheetToObjects(submissions);
  const existingIndex = all.findIndex(
    s => s.event_id === payload.eventId && String(s.employee_name).trim() === String(payload.employeeName).trim()
  );

  const isResubmit = existingIndex > -1;
  const timeNow = nowStr();
  const status = isResubmit
    ? 'มีการแก้ไขล่าสุดเมื่อ (' + timeNow + ' น.)'
    : 'แนบเอกสารแล้ว (' + timeNow + ' น.)';

  const submissionId = isResubmit ? all[existingIndex].submission_id : generateId('SUB');
  // Round to strictly 2 decimal places (e.g. 150.00, 1250.50) before
  // persisting, so stored amounts never carry stray floating-point digits.
  const safeTotalAmount = Math.round((Number(payload.totalAmount) || 0) * 100) / 100;

  const rowData = [
    submissionId,
    payload.eventId,
    String(payload.employeeName).trim(),
    String(payload.department || '').trim(),
    status,
    timeNow,
    safeTotalAmount,
    payload.summaryText || '',
    pdfUrl || (isResubmit ? all[existingIndex].pdf_file_url : ''),
    JSON.stringify(storedDetails)
  ];

  if (isResubmit) {
    submissions.getRange(existingIndex + 2, 1, 1, rowData.length).setValues([rowData]);
  } else {
    submissions.appendRow(rowData);
  }

  // The roster panel (index.html) must always reflect a just-submitted
  // claim immediately — never wait out the cache TTL.
  cacheRemove_('roster_' + payload.eventId);

  return { ok: true, submissionId: submissionId, status: status, pdfUrl: pdfUrl, resubmitted: isResubmit };
}

// ---------- API: ดึงประวัติการส่งเอกสารของพนักงานคนนั้นๆ ----------
function apiGetMySubmission(payload) {
  if (!payload || !payload.eventId || !payload.employeeName) {
    return { ok: false, error: 'ระบุข้อมูลไม่สมบูรณ์' };
  }

  const subs = sheetToObjects(getSheet(SHEET_SUBMISSIONS));
  const sub = subs.find(
    s => s.event_id === payload.eventId && String(s.employee_name).trim() === String(payload.employeeName).trim()
  );
  if (!sub) return { ok: true, exists: false };

  let details = { purpose: '', fuelRefPrice: '', trips: [] };
  try {
    details = JSON.parse(sub.trip_details_json);
  } catch (err) { }

  // Pull previously-uploaded receipt images back down from Drive as Base64
  // so the edit form can (a) show a real <img> preview of what was already
  // submitted, and (b) feed them straight into the same seg.attachments
  // array the PDF generator already knows how to render — no client-side
  // Drive fetch, no CORS, no broken thumbnails.
  const trips = (details.trips || []).map(trip => {
    const segments = (trip.segments || []).map(seg => {
      const urls = Array.isArray(seg.attachmentUrls) ? seg.attachmentUrls : [];
      const existingAttachments = urls
        .map(url => driveFileToBase64Attachment_(url))
        .filter(att => att !== null);
      return Object.assign({}, seg, { existingAttachments: existingAttachments });
    });
    return Object.assign({}, trip, { segments: segments });
  });

  return {
    ok: true, 
    exists: true,
    data: {
      department: sub.department || '',
      purpose: details.purpose || '',
      fuelRefPrice: details.fuelRefPrice || '',
      trips: trips,
      status: sub.submission_status,
      updatedAt: sub.updated_at
    }
  };
}

// ---------- API: ดึงไฟล์ใดๆ ใน Drive กลับมาเป็น Base64 ----------
// Used by index.html's bulk "PDF -> ZIP" export: fetching a Drive share
// link directly from the browser returns an HTML viewer page (not the raw
// file bytes) and is blocked by CORS anyway, so the client asks the backend
// to read the file server-side and hand back ready-to-embed Base64 instead.
// Reuses the same helper already trusted for receipt-image previews.
function apiGetFileAsBase64(payload) {
  if (!payload || !payload.url) return { ok: false, error: 'ไม่ได้ระบุ URL ไฟล์' };
  const att = driveFileToBase64Attachment_(payload.url);
  if (!att) return { ok: false, error: 'ไม่สามารถอ่านไฟล์นี้ได้ (อาจถูกลบ หรือไม่มีสิทธิ์เข้าถึง)' };
  return { ok: true, base64: att.base64, fileName: att.fileName, mimeType: att.mimeType };
}

// ---------- API: ดึงรูปใบเสร็จราคาน้ำมัน (แบบ Base64) เฉพาะตอนที่ต้องใช้จริง ----------
// Deliberately separate from apiGetEventInfo (see getEventFuelPrices_ above)
// so the common "just load the form/event" path never pays for Drive
// image fetches it doesn't need yet.
function apiGetFuelReceiptImages(payload) {
  if (!payload || !payload.eventId) return { ok: false, error: 'ไม่ได้ระบุรหัสกิจกรรม' };
  return { ok: true, fuelPrices: getEventFuelPricesWithImages_(payload.eventId) };
}

// ---------- API: ราคาน้ำมันอ้างอิง ----------
function apiGetFuelPrice() {
  return { 
    ok: true, 
    price: 34.57, 
    source: 'manual_fallback', 
    note: 'ราคานี้เป็นค่าอ้างอิงที่ปรับปรุงด้วยมือ ไม่ใช่ราคาสด กรุณาตรวจสอบราคาจริงและแก้ไขได้ในฟอร์มเสมอ' 
  };
}
