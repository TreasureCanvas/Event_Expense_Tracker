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

// doPost — Lock เฉพาะ action ที่เขียนข้อมูลลง Sheet โดยตรง (createEvent, updateEvent)
// submitExpense ล็อกตัวเองภายในฟังก์ชัน (ดู apiSubmitExpense) เพราะมีขั้นตอนอัปโหลดไฟล์ที่ใช้เวลานาน
function doPost(e) {
  let body = {};
  if (e && e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonOut({ ok: false, error: 'TEST-MODE: JSON parse failed' });
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
  Logger.log('TEST-MODE doPost hit. action=%s', action);

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
      return jsonOut({ ok: false, error: 'TEST-MODE: ไม่พบ Action ที่ระบุ: ' + action });
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

function safeSetPublicSharing(file) {
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log('Warning: Cannot set public file sharing policy: ' + err.toString());
  }
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
  if (Array.isArray(payload.employees)) {
    payload.employees.forEach(emp => {
      if (!emp || !emp.name) return;
      empSheet.appendRow([eventId, String(emp.name).trim(), String(emp.department || '').trim()]);
    });
  }

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

  events.getRange(rowIndex + 1, nameCol + 1).setValue(String(payload.eventName).trim());
  events.getRange(rowIndex + 1, ownerCol + 1).setValue(String(payload.ownerName).trim());
  if (venueCol > -1) events.getRange(rowIndex + 1, venueCol + 1).setValue(String(payload.venue || '').trim());
  events.getRange(rowIndex + 1, startCol + 1).setValue(payload.startDate || '');
  events.getRange(rowIndex + 1, endCol + 1).setValue(payload.endDate || payload.startDate || '');

  // ---- Batched rewrite of EventEmployees: keep other events' rows, replace this event's rows in ONE write ----
  const empSheet = getSheet(SHEET_EVENT_EMPLOYEES);
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData[0];
  const empIdCol = empHeaders.indexOf('event_id');

  // เก็บบรรทัดที่เป็นของกิจกรรม "อื่นๆ" เอาไว้ทั้งหมด
  const otherRows = empData.slice(1).filter(row => row[empIdCol] !== payload.eventId);

  // สร้างบรรทัดใหม่สำหรับกิจกรรม "นี้" จากข้อมูลที่ส่งมา
  const newRows = [];
  if (Array.isArray(payload.employees)) {
    payload.employees.forEach(emp => {
      if (!emp || !emp.name) return;
      newRows.push([payload.eventId, String(emp.name).trim(), String(emp.department || '').trim()]);
    });
  }

  const finalRows = otherRows.concat(newRows);

  // ล้างข้อมูลเดิมทั้งหมดที่อยู่ใต้ Header แล้วเขียนแถวข้อมูลทั้งหมดกลับเข้าไปใหม่ในครั้งเดียว
  const lastRow = empSheet.getLastRow();
  if (lastRow > 1) {
    empSheet.getRange(2, 1, lastRow - 1, empHeaders.length).clearContent();
  }
  if (finalRows.length > 0) {
    empSheet.getRange(2, 1, finalRows.length, empHeaders.length).setValues(finalRows);
  }

  return { ok: true, eventId: payload.eventId };
}

// ---------- API: รายการกิจกรรมทั้งหมด ----------
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
  }).sort((a, b) => (new Date(b.createdAt || 0) - new Date(a.createdAt || 0)));

  return { ok: true, events: list };
}

// ---------- API: ดึงข้อมูลกิจกรรมและรายชื่อพนักงานสำหรับ Auto-complete ----------
function apiGetEventInfo(payload) {
  if (!payload || !payload.eventId) return { ok: false, error: 'ไม่ได้ระบุรหัสกิจกรรม' };
  
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

// ---------- API: ดึงสถานะการส่งเอกสารของพนักงานทุกคนในกิจกรรม ----------
function apiGetEventRoster(payload) {
  if (!payload || !payload.eventId) return { ok: false, error: 'ไม่ได้ระบุรหัสกิจกรรม' };

  const eventId = payload.eventId;
  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้' };

  const roster = sheetToObjects(getSheet(SHEET_EVENT_EMPLOYEES)).filter(x => x.event_id === eventId);
  const submissions = sheetToObjects(getSheet(SHEET_SUBMISSIONS)).filter(s => s.event_id === eventId);

  const rows = roster.map(emp => {
    const sub = submissions.find(s => String(s.employee_name).trim() === String(emp.employee_name).trim());
    if (!sub) {
      return {
        employeeName: emp.employee_name, department: emp.department,
        submitted: false, status: 'ยังไม่ส่งเอกสาร', updatedAt: '', totalAmount: 0,
        pdfUrl: '', summaryText: ''
      };
    }
    return {
      employeeName: emp.employee_name, department: emp.department,
      submitted: true, status: sub.submission_status, updatedAt: sub.updated_at,
      totalAmount: Number(sub.total_amount) || 0, pdfUrl: sub.pdf_file_url, summaryText: sub.summary_text
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

// ---------- API: บันทึก / แก้ไข การส่งเอกสารเบิกจ่าย ----------
function apiSubmitExpense(payload) {
  if (!payload || !payload.eventId || !payload.employeeName) {
    return { ok: false, error: 'ข้อมูลไม่ครบถ้วน (ต้องระบุ eventId และ employeeName)' };
  }

  const submissions = getSheet(SHEET_SUBMISSIONS);
  const events = sheetToObjects(getSheet(SHEET_EVENTS));
  const ev = events.find(e => e.event_id === payload.eventId);
  if (!ev) return { ok: false, error: 'ไม่พบกิจกรรมนี้ในระบบ' };

  const folder = getOrCreateDriveFolder();

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
      return out;
    });
    return Object.assign({}, trip, { segments: segments });
  });

  const storedDetails = {
    purpose: payload.purpose || '',
    fuelRefPrice: payload.fuelRefPrice || '',
    trips: tripDetails
  };

  // ตรวจสอบการส่งซ้ำ + เขียนแถวลง Sheet — ล็อกเฉพาะช่วงนี้เท่านั้น (เร็ว ไม่รวม Drive I/O)
  const lock = LockService.getScriptLock();
  if (!lock.waitLock(15000)) {
    return { ok: false, error: 'ระบบกำลังประมวลผลคำขอจำนวนมาก กรุณาลองใหม่อีกครั้งในอีกสักครู่' };
  }
  try {
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
    const safeTotalAmount = Number(payload.totalAmount) || 0;

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

    return { ok: true, submissionId: submissionId, status: status, pdfUrl: pdfUrl, resubmitted: isResubmit };
  } finally {
    lock.releaseLock();
  }
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

  return {
    ok: true, 
    exists: true,
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

// ---------- API: ราคาน้ำมันอ้างอิง ----------
function apiGetFuelPrice() {
  return { 
    ok: true, 
    price: 34.57, 
    source: 'manual_fallback', 
    note: 'ราคานี้เป็นค่าอ้างอิงที่ปรับปรุงด้วยมือ ไม่ใช่ราคาสด กรุณาตรวจสอบราคาจริงและแก้ไขได้ในฟอร์มเสมอ' 
  };
}
