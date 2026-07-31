/**
 * ===================================================================
 *  ระบบจัดทำและรวบรวมเอกสารเบิกค่าเดินทางเชื่อมโยงกิจกรรม
 *  (Event Expense Tracker) - Backend (Google Apps Script)
 * ===================================================================
 *  วิธี Deploy:
 *  1. เปิด Google Sheets เปล่า 1 ไฟล์ -> Extensions > Apps Script
 *  2. วางไฟล์นี้ทับ Code.gs ที่มีอยู่
 *  3. รันฟังก์ชัน setupSheets() หนึ่งครั้ง (Run > setupSheets) เพื่อสร้างชีตทั้งหมด
 *     (ครั้งแรกจะขอ Authorize สิทธิ์ - กด Allow)
 *  4. Deploy > New deployment > Type: Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     กด Deploy แล้วคัดลอก "Web app URL" (ลงท้ายด้วย /exec)
 *  5. นำ URL ไปวางแทนที่ API_URL ในไฟล์ index.html / owner.html / staff.html
 * ===================================================================
 */

// ---------- CONFIG ----------
const SHEET_EVENTS = 'Events';
const SHEET_SUBMISSIONS = 'Submissions';
const SHEET_EVENT_EMPLOYEES = 'EventEmployees';
const DRIVE_FOLDER_NAME = 'EventExpenseTracker_Files';

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

  getOrCreateDriveFolder();
  Logger.log('Setup complete.');
}

function getOrCreateDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// ---------- SAFE SCHEMA MIGRATION (for spreadsheets created before the
// venue/form_link columns existed) ----------
// Run this ONCE from the Apps Script editor (Run > migrateEventsSchema) if
// your Events sheet is missing the "venue" or "form_link" columns.
// Unlike re-running setupSheets(), this does NOT clear any existing data -
// it only inserts the missing columns at the correct position and backfills
// form_link for events that don't have one yet.
function migrateEventsSchema() {
  const sheet = getSheet(SHEET_EVENTS);
  const lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  function ensureColumnBefore(headerName, beforeHeaderName) {
    if (headers.indexOf(headerName) > -1) return; // already exists - nothing to do
    let insertAt = headers.indexOf(beforeHeaderName) + 1; // 1-indexed column position
    if (insertAt === 0) insertAt = headers.length + 1; // reference column not found - append at end instead
    sheet.insertColumnBefore(insertAt);
    sheet.getRange(1, insertAt).setValue(headerName);
    headers.splice(insertAt - 1, 0, headerName); // keep local copy in sync for the next check
  }

  ensureColumnBefore('venue', 'start_date');
  ensureColumnBefore('form_link', 'event_token');

  // backfill form_link for any existing events that don't have one yet
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

  Logger.log('Migration complete. Columns now: ' + sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].join(', '));
}

// ---------- WEB ENTRY POINTS ----------
function doGet(e) {
  const page = (e.parameter.page || 'home');
  let tmpl;
  if (page === 'form') tmpl = HtmlService.createTemplateFromFile('index');
  else tmpl = HtmlService.createTemplateFromFile('home'); // 'home', and legacy 'owner'/'staff' links all land here
  tmpl.eventId = e.parameter.event || '';
  return tmpl.evaluate()
    .setTitle('Event Expense Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// doPost is the JSON API used by the standalone HTML files embedded in
// Google Sites (fetch() calls land here).
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: 'Invalid JSON body' });
  }

  const action = body.action;
  try {
    switch (action) {
      case 'createEvent':
        return jsonOut(apiCreateEvent(body.payload));
      case 'updateEvent':
        return jsonOut(apiUpdateEvent(body.payload));
      case 'listAllEvents':
        return jsonOut(apiListAllEvents());
      case 'getEventInfo':
        return jsonOut(apiGetEventInfo(body.payload));
      case 'getEventRoster':
        return jsonOut(apiGetEventRoster(body.payload));
      case 'submitExpense':
        return jsonOut(apiSubmitExpense(body.payload));
      case 'getMySubmission':
        return jsonOut(apiGetMySubmission(body.payload));
      case 'getFuelPrice':
        return jsonOut(apiGetFuelPrice());
      default:
        return jsonOut({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
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
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
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

// ---------- API: Staff creates a new event ----------
// payload: { eventName, ownerName, startDate, endDate, employees: [{name, department}, ...] }
function apiCreateEvent(payload) {
  if (!payload.eventName || !payload.ownerName) {
    return { ok: false, error: 'กรุณากรอกชื่อกิจกรรมและชื่อผู้สร้าง' };
  }

  const events = getSheet(SHEET_EVENTS);
  const eventId = 'EV-' + new Date().getFullYear() + '-' + generateId('').replace('-', '');
  const scriptUrl = ScriptApp.getService().getUrl();
  const formLink = scriptUrl + '?page=form&event=' + eventId;

  events.appendRow([
    eventId, payload.eventName, payload.ownerName, payload.venue || '',
    payload.startDate || '', payload.endDate || payload.startDate || '',
    formLink, '', nowStr()
  ]);

  const empSheet = getSheet(SHEET_EVENT_EMPLOYEES);
  (payload.employees || []).forEach(emp => {
    if (!emp.name) return;
    empSheet.appendRow([eventId, emp.name, emp.department || '']);
  });

  return {
    ok: true,
    eventId: eventId,
    formLink: formLink
  };
}

// ---------- API: edit an existing event's details + replace its staff roster ----------
// payload: { eventId, eventName, ownerName, startDate, endDate, employees: [{name, department}, ...] }
function apiUpdateEvent(payload) {
  if (!payload.eventId) return { ok: false, error: 'ไม่พบรหัสกิจกรรมที่จะแก้ไข' };
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

  // update event_name / owner_name / venue / start_date / end_date only
  // (event_id, form_link, event_token, created_at stay untouched)
  const nameCol = headers.indexOf('event_name');
  const ownerCol = headers.indexOf('owner_name');
  const venueCol = headers.indexOf('venue');
  const startCol = headers.indexOf('start_date');
  const endCol = headers.indexOf('end_date');

  events.getRange(rowIndex + 1, nameCol + 1).setValue(payload.eventName);
  events.getRange(rowIndex + 1, ownerCol + 1).setValue(payload.ownerName);
  if (venueCol > -1) events.getRange(rowIndex + 1, venueCol + 1).setValue(payload.venue || '');
  events.getRange(rowIndex + 1, startCol + 1).setValue(payload.startDate || '');
  events.getRange(rowIndex + 1, endCol + 1).setValue(payload.endDate || payload.startDate || '');

  // replace the staff roster wholesale: delete this event's existing rows, re-add the submitted list
  const empSheet = getSheet(SHEET_EVENT_EMPLOYEES);
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData[0];
  const empIdCol = empHeaders.indexOf('event_id');
  const rowsToDelete = [];
  for (let r = 1; r < empData.length; r++) {
    if (empData[r][empIdCol] === payload.eventId) rowsToDelete.push(r + 1); // +1 for 1-indexed sheet rows
  }
  rowsToDelete.reverse().forEach(r => empSheet.deleteRow(r)); // reverse so row numbers don't shift mid-delete

  (payload.employees || []).forEach(emp => {
    if (!emp.name) return;
    empSheet.appendRow([payload.eventId, emp.name, emp.department || '']);
  });

  return { ok: true, eventId: payload.eventId };
}

// ---------- API: list every event in the database (for the main dashboard dropdown) ----------
function apiListAllEvents() {
  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const employees = sheetToObjects(getSheet(SHEET_EVENT_EMPLOYEES));

  const list = events.map(ev => {
    const staffCount = employees.filter(e => e.event_id === ev.event_id).length;
    return {
      eventId: ev.event_id,
      eventName: ev.event_name,
      ownerName: ev.owner_name,
      venue: ev.venue || '',
      startDate: dateStr(ev.start_date),
      endDate: dateStr(ev.end_date),
      formLink: ev.form_link || (ScriptApp.getService().getUrl() + '?page=form&event=' + ev.event_id),
      staffCount: staffCount,
      createdAt: ev.created_at
    };
  }).sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1));

  return { ok: true, events: list };
}

// ---------- API: get event info + employee list (for autocomplete on the form) ----------
function apiGetEventInfo(payload) {
  const eventId = payload.eventId;
  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้ในระบบ (Event ID ไม่ถูกต้อง)' };

  const employees = sheetToObjects(getSheet(SHEET_EVENT_EMPLOYEES))
    .filter(x => x.event_id === eventId)
    .map(x => ({ name: x.employee_name, department: x.department }));

  return {
    ok: true,
    event: {
      eventId: ev.event_id, eventName: ev.event_name, ownerName: ev.owner_name, venue: ev.venue || '',
      startDate: dateStr(ev.start_date), endDate: dateStr(ev.end_date)
    },
    employees: employees
  };
}

// ---------- API: full roster for one event — every assigned employee, whether
// they have submitted yet or not, so the dashboard can show "ส่งแล้ว / ยังไม่ส่ง" ----------
function apiGetEventRoster(payload) {
  const eventId = payload.eventId;
  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้' };

  const roster = sheetToObjects(getSheet(SHEET_EVENT_EMPLOYEES)).filter(x => x.event_id === eventId);
  const submissions = sheetToObjects(getSheet(SHEET_SUBMISSIONS)).filter(s => s.event_id === eventId);

  const rows = roster.map(emp => {
    const sub = submissions.find(s => s.employee_name === emp.employee_name);
    if (!sub) {
      return {
        employeeName: emp.employee_name, department: emp.department,
        submitted: false, status: 'ยังไม่ส่งเอกสาร', updatedAt: '', totalAmount: '',
        pdfUrl: '', summaryText: ''
      };
    }
    return {
      employeeName: emp.employee_name, department: emp.department,
      submitted: true, status: sub.submission_status, updatedAt: sub.updated_at,
      totalAmount: sub.total_amount, pdfUrl: sub.pdf_file_url, summaryText: sub.summary_text
    };
  });

  return {
    ok: true,
    event: {
      eventId: ev.event_id, eventName: ev.event_name, ownerName: ev.owner_name, venue: ev.venue || '',
      startDate: dateStr(ev.start_date), endDate: dateStr(ev.end_date)
    },
    roster: rows
  };
}

// ---------- API: employee submits/edits an expense claim ----------
// payload: {
//   eventId, employeeName, department, purpose, fuelRefPrice,
//   totalAmount, summaryText, pdfBase64, pdfFileName, tripDetails: [ {label, segments:[...]}, ... ]
// }
// each segment: { type, origin, destination, distanceKm, manualAmount, amount,
//                 attachments: [{base64, fileName}], attachmentUrls: [existing urls kept] }
function apiSubmitExpense(payload) {
  const submissions = getSheet(SHEET_SUBMISSIONS);
  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === payload.eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้ในระบบ' };

  const folder = getOrCreateDriveFolder();

  // save PDF to Drive
  let pdfUrl = '';
  if (payload.pdfBase64) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(payload.pdfBase64),
      'application/pdf',
      payload.pdfFileName || 'expense.pdf'
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    pdfUrl = file.getUrl();
  }

  // save any newly-attached slip images to Drive; keep URLs of files that
  // were already attached in a previous submission (carried over by the client)
  const tripDetails = (payload.tripDetails || []).map((trip, ti) => {
    const segments = (trip.segments || []).map((seg, si) => {
      const out = Object.assign({}, seg);
      const keptUrls = Array.isArray(seg.attachmentUrls) ? seg.attachmentUrls.slice() : [];
      const newAttachments = Array.isArray(seg.attachments) ? seg.attachments : [];

      newAttachments.forEach((att, ai) => {
        if (!att || !att.base64) return;
        try {
          const match = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/.exec(att.base64);
          const mime = match ? match[1] : 'image/jpeg';
          const data = match ? match[2] : att.base64;
          const ext = (mime.split('/')[1] || 'jpg');
          const imgBlob = Utilities.newBlob(
            Utilities.base64Decode(data), mime,
            'slip_' + payload.employeeName + '_' + trip.label + '_seg' + (si + 1) + '_' + (ai + 1) + '.' + ext
          );
          const imgFile = folder.createFile(imgBlob);
          imgFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          keptUrls.push(imgFile.getUrl());
        } catch (imgErr) {
          // skip this one file on error, keep the rest
        }
      });

      out.attachmentUrls = keptUrls;
      delete out.attachments;
      return out;
    });
    return Object.assign({}, trip, { segments: segments });
  });

  const storedDetails = {
    purpose: payload.purpose || '',
    fuelRefPrice: payload.fuelRefPrice || '',
    trips: tripDetails
  };

  // check for existing submission by this employee for this event -> overwrite (resubmit)
  const all = sheetToObjects(submissions);
  const existingIndex = all.findIndex(
    s => s.event_id === payload.eventId && s.employee_name === payload.employeeName
  );

  const isResubmit = existingIndex > -1;
  const status = isResubmit
    ? 'มีการแก้ไขล่าสุดเมื่อ (' + nowStr() + ' น.)'
    : 'แนบเอกสารแล้ว (' + nowStr() + ' น.)';

  const submissionId = isResubmit ? all[existingIndex].submission_id : generateId('SUB');

  const rowData = [
    submissionId, payload.eventId, payload.employeeName, payload.department || '',
    status, nowStr(), payload.totalAmount || 0,
    payload.summaryText || '', pdfUrl, JSON.stringify(storedDetails)
  ];

  if (isResubmit) {
    submissions.getRange(existingIndex + 2, 1, 1, rowData.length).setValues([rowData]);
  } else {
    submissions.appendRow(rowData);
  }

  return { ok: true, submissionId: submissionId, status: status, pdfUrl: pdfUrl, resubmitted: isResubmit };
}

// ---------- API: check if this employee already submitted for this event,
// so the form can reload their previous data for editing ----------
function apiGetMySubmission(payload) {
  const subs = sheetToObjects(getSheet(SHEET_SUBMISSIONS));
  const sub = subs.find(s => s.event_id === payload.eventId && s.employee_name === payload.employeeName);
  if (!sub) return { ok: true, exists: false };

  let details = { purpose: '', fuelRefPrice: '', trips: [] };
  try {
    details = JSON.parse(sub.trip_details_json);
  } catch (err) { /* leave defaults if malformed/old-format data */ }

  return {
    ok: true, exists: true,
    data: {
      department: sub.department || '',
      purpose: details.purpose || '',
      fuelRefPrice: details.fuelRefPrice || '',
      trips: details.trips || [],
      status: sub.submission_status,
      updatedAt: sub.updated_at
    }
  };
}

// ---------- API: best-effort retail fuel price (Gasohol 91) ----------
// Google Sites embed cannot reliably scrape live prices client-side, so we
// fetch server-side (no CORS restrictions here) from a public source and
// fall back to a manual default if it fails. Employees can always override
// the value manually in the form (ราคาน้ำมันอ้างอิง field).
function apiGetFuelPrice() {
  try {
    // NOTE: replace with a real, licensed data source in production.
    // NOTE: this is a manually-maintained fallback (updated ~Jul 2026), not a
    // live feed - Apps Script has no reliable free public API for this. The
    // employee can and should always override it in the form to match the
    // actual price on their travel date (e.g. check gasprice.kapook.com).
    return { ok: true, price: 34.57, source: 'manual_fallback', note: 'ราคานี้เป็นค่าอ้างอิงที่ปรับปรุงด้วยมือ ไม่ใช่ราคาสด กรุณาตรวจสอบราคาจริงและแก้ไขได้ในฟอร์มเสมอ' };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}
