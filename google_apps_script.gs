/**
 * ═══════════════════════════════════════════════════════
 *  FIELD EXPENSE TRACKER — Google Apps Script Backend
 *  v3 — Delta Sync + Soft Delete + Multi-user safe
 * ═══════════════════════════════════════════════════════
 *
 *  วิธีติดตั้ง / อัปเดต:
 *  1. Extensions → Apps Script → วางโค้ดนี้แทนที่ของเดิม
 *  2. Deploy → New deployment (Web App, Execute as: Me, Anyone)
 *  3. คัดลอก URL ใหม่ใส่ใน Dashboard → Settings
 *  4. กด "สร้าง Sheet โครงสร้าง" เพื่อเพิ่ม columns ใหม่
 *
 *  การเปลี่ยนแปลง v3:
 *  • Expenses เพิ่มคอลัมน์ updatedAt (col 26) และ deletedAt (col 27)
 *  • deleteExpense → Soft delete (ไม่ลบ row จริง แค่ set deletedAt)
 *  • GET action "getDeltas?since=ISO" → คืน delta ตั้งแต่ timestamp นั้น
 *  • addExpense / updateExpense set updatedAt = now() ทุกครั้ง
 *  • getExpenses กรอง deletedAt ออกโดยอัตโนมัติ
 * ═══════════════════════════════════════════════════════
 */

// ─────────────────────────────────────
//  CORS + ROUTING
// ─────────────────────────────────────
// ─────────────────────────────────────
//  UTILITIES
// ─────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetToJson(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getRootUploadFolder() {
  const folderName = 'FieldExpenseTracker';
  const iter = DriveApp.getFoldersByName(folderName);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(folderName);
}

function getOrCreateSubFolder(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function getTimestamp(dateStr) {
  try {
    const d = dateStr ? new Date(dateStr) : new Date();
    const y  = d.getFullYear();
    const mo = String(d.getMonth()+1).padStart(2,'0');
    const dy = String(d.getDate()).padStart(2,'0');
    const h  = String(d.getHours()).padStart(2,'0');
    const mi = String(d.getMinutes()).padStart(2,'0');
    const s  = String(d.getSeconds()).padStart(2,'0');
    return y + mo + dy + '_' + h + mi + s;
  } catch(_) { return String(Date.now()); }
}

// ─────────────────────────────────────
//  CORS + ROUTING
// ─────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'getData';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let result = {};

    if (action === 'ping') {
      result = { success: true, pong: true, serverTime: new Date().toISOString() };

    } else if (action === 'countDeleted') {
      // นับจำนวน soft-deleted rows ใน Expenses sheet
      const sheet = ss.getSheetByName('Expenses');
      let count = 0;
      if (sheet && sheet.getLastRow() >= 2) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const iDel = headers.indexOf('deletedAt');
        if (iDel >= 0) {
          count = data.slice(1).filter(r => r[iDel] && String(r[iDel]).trim() !== '').length;
        }
      }
      result = { success: true, count };

    } else if (action === 'getData') {
      // Full pull — projects/activities/categories (ไม่รวม expenses)
      result = {
        success:    true,
        serverTime: new Date().toISOString(),
        projects:   sheetToJson(ss, 'Projects'),
        activities: sheetToJson(ss, 'Activities'),
        categories: sheetToJson(ss, 'Categories'),
      };

    } else if (action === 'getDeltas') {
      // ─── DELTA SYNC ───────────────────────────────────────
      // ส่ง ?since=2024-01-01T00:00:00.000Z เพื่อรับเฉพาะข้อมูลที่เปลี่ยนแปลง
      // ถ้า since = '' หรือไม่ส่ง → full sync (ส่งข้อมูลทั้งหมด)
      const since = e.parameter.since || '';
      result = {
        success:    true,
        serverTime: new Date().toISOString(),
        projects:   sheetToJson(ss, 'Projects'),
        activities: sheetToJson(ss, 'Activities'),
        categories: sheetToJson(ss, 'Categories'),
        expenses:   getActiveExpensesSince(ss, since),
        deletedIds: getDeletedExpenseIdsSince(ss, since),
      };

    } else if (action === 'getExpenses') {
      // Legacy endpoint — คืน active expenses ทั้งหมด
      const projectId = e.parameter.projectId || '';
      let expenses = getActiveExpensesSince(ss, '');
      if (projectId) expenses = expenses.filter(r => r.projectId === projectId);
      result = { success: true, expenses };

    } else {
      result = { success: false, error: 'Unknown action: ' + action };
    }

    return jsonOut(result);
  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    const ss      = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'setupSheets')          return jsonOut(setupSheets(ss));
    if (action === 'addExpense')           return jsonOut(addExpense(ss, payload.expense));
    if (action === 'updateExpense')        return jsonOut(updateExpense(ss, payload.expense));
    if (action === 'deleteExpense')        return jsonOut(deleteExpense(ss, payload.id));
    if (action === 'uploadImage')          return jsonOut(uploadImage(payload));
    if (action === 'deleteImage')          return jsonOut(deleteImage(payload.fileId));
    if (action === 'deleteProjectFolder')  return jsonOut(deleteProjectFolder(payload));
    if (action === 'deleteActivityFolder') return jsonOut(deleteActivityFolder(payload));
    if (action === 'moveFile')             return jsonOut(moveFile(payload));
    if (action === 'cloudBackupSave')      return jsonOut(saveCloudBackup(payload));
    if (action === 'cloudBackupLoad')      return jsonOut(loadCloudBackup());
    if (action === 'login')                return jsonOut(loginUser(ss, payload));
    if (action === 'getUsers')             return jsonOut(getUsers(ss));
    if (action === 'addUser')              return jsonOut(addRow(ss, 'Users', payload.user));
    if (action === 'updateUser')           return jsonOut(updateSheetRow(ss, 'Users', payload.user));
    if (action === 'deleteUser')           return jsonOut(deleteRowById(ss, 'Users', payload.id));
    if (action === 'addProject')           return jsonOut(addRow(ss, 'Projects',   payload.project));
    if (action === 'addActivity')          return jsonOut(addRow(ss, 'Activities', payload.activity));
    if (action === 'updateProject')        return jsonOut(updateSheetRow(ss, 'Projects',   payload.project));
    if (action === 'updateActivity')       return jsonOut(updateSheetRow(ss, 'Activities', payload.activity));
    if (action === 'deleteProject')        return jsonOut(deleteProject(ss, payload.id));
    if (action === 'deleteActivity')       return jsonOut(deleteActivity(ss, payload.id));
    if (action === 'addCategory')          return jsonOut(addRow(ss, 'Categories', payload.category));
    if (action === 'deleteCategory')       return jsonOut(deleteCategoryRow(ss, payload.category));
    if (action === 'deleteCategoryGroup')  return jsonOut(deleteCategoryGroup(ss, payload.l1));
    if (action === 'purgeDeleted')         return jsonOut(purgeDeleted(ss, payload.olderThanDays));

    return jsonOut({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

// ─────────────────────────────────────
//  SETUP SHEETS
// ─────────────────────────────────────
function setupSheets(ss) {
  const defs = [
    {
      name: 'Expenses',
      // col: 1-25 เหมือนเดิม, เพิ่ม col 26=updatedAt, 27=deletedAt
      headers: ['id','date','projectId','projectName','activityId','activityName',
                'catL1','catL2','catL3','docType','description','amount',
                'vatRate','vatAmount','whtRate','whtAmount','netAmount',
                'payment','slipUrls','rcptUrls','note','createdAt','synced',
                'payer','addedBy','updatedAt','deletedAt'],
      widths: [100,90,100,160,100,160,130,160,160,120,220,100,70,100,70,100,100,90,220,220,180,160,60,140,140,160,160],
    },
    {
      name: 'Projects',
      headers: ['id','name','code','budget','status'],
      widths: [120,220,80,100,80],
    },
    {
      name: 'Activities',
      headers: ['id','projectId','projectName','name','budget','status'],
      widths: [120,120,160,220,100,80],
    },
    {
      name: 'Users',
      headers: ['id','username','displayName','password','role','active'],
      widths:  [100, 120, 180, 120, 80, 60],
      sample:  [['USR-1','admin','ผู้ดูแลระบบ','admin1234','admin',true]],
    },
    {
      name: 'Categories',
      headers: ['l1','l2','l3'],
      widths: [150,180,220],
      sample: [
        ['ค่าตอบแทน','ค่าตอบแทนวิทยากร','วิทยากรภายนอก'],
        ['ค่าตอบแทน','ค่าตอบแทนวิทยากร','วิทยากรภายใน'],
        ['ค่าตอบแทน','ค่าตอบแทนผู้เชี่ยวชาญ','ที่ปรึกษาโครงการ'],
        ['ค่าตอบแทน','ค่าตอบแทนเจ้าหน้าที่','เจ้าหน้าที่โครงการ'],
        ['ค่าตอบแทน','ค่าตอบแทนเจ้าหน้าที่','ผู้ช่วยนักวิจัย'],
        ['ค่าใช้สอย','ค่าพาหนะ','ค่าน้ำมันเชื้อเพลิง'],
        ['ค่าใช้สอย','ค่าพาหนะ','ค่าเช่ารถยนต์'],
        ['ค่าใช้สอย','ค่าพาหนะ','ค่าตั๋วเครื่องบิน'],
        ['ค่าใช้สอย','ค่าที่พัก','ค่าโรงแรม/เกสต์เฮาส์'],
        ['ค่าใช้สอย','ค่าอาหาร','อาหารกลางวัน'],
        ['ค่าใช้สอย','ค่าอาหาร','อาหารว่างและเครื่องดื่ม'],
        ['ค่าใช้สอย','ค่าจ้างเหมา','ค่าจ้างเหมาบริการ'],
        ['ค่าใช้สอย','ค่าจัดประชุม/สัมมนา','ค่าเช่าห้องประชุม'],
        ['ค่าวัสดุ','วัสดุสำนักงาน','กระดาษ/เครื่องเขียน'],
        ['ค่าวัสดุ','วัสดุคอมพิวเตอร์','อุปกรณ์ IT'],
        ['ค่าวัสดุ','วัสดุงานสนาม','อุปกรณ์ภาคสนาม'],
        ['ค่าสาธารณูปโภค','ค่าโทรศัพท์/สื่อสาร','ค่าโทรศัพท์มือถือ'],
        ['ค่าสาธารณูปโภค','ค่าโทรศัพท์/สื่อสาร','ค่า Internet/Data'],
      ]
    },
  ];

  defs.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }

    // ── เพิ่ม columns ที่ขาดหายไป (backward compat) ──
    if (sheet.getLastRow() >= 1) {
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      def.headers.forEach((h, idx) => {
        if (!existingHeaders.includes(h)) {
          const newCol = existingHeaders.length + 1;
          sheet.getRange(1, newCol).setValue(h).setFontWeight('bold')
            .setBackground('#2563EB').setFontColor('#FFFFFF');
          existingHeaders.push(h);
        }
      });
    }

    // Write header if completely empty
    if (sheet.getLastRow() === 0) {
      const hRow = sheet.getRange(1, 1, 1, def.headers.length);
      hRow.setValues([def.headers]);
      hRow.setFontWeight('bold');
      hRow.setBackground('#2563EB');
      hRow.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      def.widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
      if (def.sample) {
        def.sample.forEach(row => sheet.appendRow(row));
      }
    }
  });

  // Remove default "Sheet1" if empty
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() <= 1) {
    try { ss.deleteSheet(sheet1); } catch(e) {}
  }

  return { success: true, message: 'Sheets ready (v3 — Delta Sync)' };
}

// ─────────────────────────────────────
//  AUTH — LOGIN & USER MANAGEMENT
// ─────────────────────────────────────
function loginUser(ss, payload) {
  const { username, password } = payload;
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success: false, error: 'ไม่พบตาราง Users — กรุณากด "สร้าง Sheet โครงสร้าง" ก่อน' };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, error: 'ยังไม่มีผู้ใช้ในระบบ' };
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => row[h] = data[i][j]);
    if (String(row.username) === String(username) &&
        String(row.password) === String(password) &&
        String(row.active) !== 'false' && row.active !== false) {
      return { success: true, user: {
        id:          row.id,
        username:    row.username,
        displayName: row.displayName,
        role:        row.role || 'user',
      }};
    }
  }
  return { success: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function getUsers(ss) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success: false, error: 'Users sheet not found' };
  const data = sheetToJson(ss, 'Users');
  return { success: true, users: data.map(u => ({
    id: u.id, username: u.username, displayName: u.displayName,
    role: u.role, active: u.active,
  }))};
}

function deleteRowById(ss, sheetName, id) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: sheetName + ' sheet not found' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Row not found: ' + id };
}

// ─────────────────────────────────────
//  EXPENSES — DELTA SYNC ENGINE
// ─────────────────────────────────────

// คืน column index ของ header name (1-based)
function getColIndex(headers, name) {
  const idx = headers.indexOf(name);
  return idx >= 0 ? idx + 1 : -1;
}

// คืน active expenses ที่ updatedAt >= since (หรือทั้งหมดถ้า since='')
function getActiveExpensesSince(ss, since) {
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const sinceMs = since ? new Date(since).getTime() : 0;

  return data.slice(1)
    .filter(row => {
      if (!row.some(c => c !== '' && c !== null)) return false; // blank row
      const deletedAt = String(row[headers.indexOf('deletedAt')] || '');
      if (deletedAt) return false; // soft-deleted
      if (since) {
        const updatedAt = String(row[headers.indexOf('updatedAt')] || '');
        const createdAt = String(row[headers.indexOf('createdAt')] || '');
        const ts = updatedAt || createdAt;
        if (ts && new Date(ts).getTime() < sinceMs) return false;
      }
      return true;
    })
    .map(row => {
      const obj = {};
      const tz = Session.getScriptTimeZone();
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) {
          // ฟิลด์ 'date' ใช้ format YYYY-MM-DD (ไม่มี timezone shift)
          // ฟิลด์อื่น (createdAt, updatedAt ฯลฯ) ใช้ ISO string
          val = (h === 'date')
            ? Utilities.formatDate(val, tz, 'yyyy-MM-dd')
            : val.toISOString();
        }
        obj[h] = val;
      });
      // แปลง slipUrls / rcptUrls CSV → array
      obj.slipUrls = obj.slipUrls ? String(obj.slipUrls).split(',').map(s=>s.trim()).filter(Boolean) : [];
      obj.rcptUrls = obj.rcptUrls ? String(obj.rcptUrls).split(',').map(s=>s.trim()).filter(Boolean) : [];
      // backfill: expense เก่าที่ไม่มี addedBy → ใช้ payer แทน (ป้องกัน canEdit ล้มเหลว)
      if (!obj.addedBy && obj.payer) obj.addedBy = obj.payer;
      return obj;
    });
}

// คืน IDs ของ expenses ที่ถูก soft-delete ตั้งแต่ since
function getDeletedExpenseIdsSince(ss, since) {
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const sinceMs = since ? new Date(since).getTime() : 0;

  return data.slice(1)
    .filter(row => {
      if (!row.some(c => c !== '' && c !== null)) return false;
      const deletedAt = String(row[headers.indexOf('deletedAt')] || '');
      if (!deletedAt) return false; // not deleted
      if (since) {
        return new Date(deletedAt).getTime() >= sinceMs;
      }
      return true;
    })
    .map(row => String(row[0])); // return IDs
}

function addExpense(ss, exp) {
  const sheet = getOrCreateSheet(ss, 'Expenses');
  const now   = new Date().toISOString();
  const expId = exp.id || 'EXP-' + Date.now();

  // ── ใช้ header-based approach (ป้องกัน column order ไม่ตรง) ──
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  // ถ้ายังไม่มี header แถวแรก → ให้ setupSheets จัดการ (fallback positional)
  if (headers.length === 0 || !headers.includes('id')) {
    const slipStr = Array.isArray(exp.slipUrls) ? exp.slipUrls.filter(Boolean).join(',') : (exp.slipUrls||'');
    const rcptStr = Array.isArray(exp.rcptUrls) ? exp.rcptUrls.filter(Boolean).join(',') : (exp.rcptUrls||'');
    sheet.appendRow([
      expId, exp.date||'', exp.projectId||'', exp.projectName||'',
      exp.activityId||'', exp.activityName||'',
      exp.catL1||'', exp.catL2||'', exp.catL3||'', exp.docType||'',
      exp.description||'', parseFloat(exp.amount)||0,
      parseFloat(exp.vatRate)||0, parseFloat(exp.vatAmount)||0,
      parseFloat(exp.whtRate)||0, parseFloat(exp.whtAmount)||0,
      parseFloat(exp.netAmount)||0, exp.payment||'',
      slipStr, rcptStr, exp.note||'', exp.createdAt||now, true,
      exp.payer||'', exp.addedBy||'', now, ''
    ]);
    return { success: true, id: expId, updatedAt: now };
  }

  // header-based: สร้าง row ตาม header จริงของ Sheet
  const data = {
    id:          expId,
    date:        exp.date         || '',
    projectId:   exp.projectId    || '',
    projectName: exp.projectName  || '',
    activityId:  exp.activityId   || '',
    activityName:exp.activityName || '',
    catL1:       exp.catL1        || '',
    catL2:       exp.catL2        || '',
    catL3:       exp.catL3        || '',
    docType:     exp.docType      || '',
    description: exp.description  || '',
    amount:      parseFloat(exp.amount)    || 0,
    vatRate:     parseFloat(exp.vatRate)   || 0,
    vatAmount:   parseFloat(exp.vatAmount) || 0,
    whtRate:     parseFloat(exp.whtRate)   || 0,
    whtAmount:   parseFloat(exp.whtAmount) || 0,
    netAmount:   parseFloat(exp.netAmount) || 0,
    payment:     exp.payment      || '',
    slipUrls:    Array.isArray(exp.slipUrls) ? exp.slipUrls.filter(Boolean).join(',') : (exp.slipUrls||''),
    rcptUrls:    Array.isArray(exp.rcptUrls) ? exp.rcptUrls.filter(Boolean).join(',') : (exp.rcptUrls||''),
    note:        exp.note         || '',
    createdAt:   exp.createdAt    || now,
    synced:      true,
    payer:       exp.payer        || '',
    addedBy:     exp.addedBy      || '',
    updatedAt:   now,
    deletedAt:   '',
  };
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return { success: true, id: expId, updatedAt: now };
}

function updateExpense(ss, exp) {
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet) return { success: false, error: 'Expenses sheet not found' };
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(exp.id)) {
      const row = i + 1;
      const now = new Date().toISOString();
      sheet.getRange(row, 2).setValue(exp.date||'');
      sheet.getRange(row, 3).setValue(exp.projectId||'');
      sheet.getRange(row, 4).setValue(exp.projectName||'');
      sheet.getRange(row, 5).setValue(exp.activityId||'');
      sheet.getRange(row, 6).setValue(exp.activityName||'');
      sheet.getRange(row, 7).setValue(exp.catL1||'');
      sheet.getRange(row, 8).setValue(exp.catL2||'');
      sheet.getRange(row, 9).setValue(exp.catL3||'');
      sheet.getRange(row,10).setValue(exp.docType||'');
      sheet.getRange(row,11).setValue(exp.description||'');
      sheet.getRange(row,12).setValue(parseFloat(exp.amount)||0);
      sheet.getRange(row,13).setValue(parseFloat(exp.vatRate)||0);
      sheet.getRange(row,14).setValue(parseFloat(exp.vatAmount)||0);
      sheet.getRange(row,15).setValue(parseFloat(exp.whtRate)||0);
      sheet.getRange(row,16).setValue(parseFloat(exp.whtAmount)||0);
      sheet.getRange(row,17).setValue(parseFloat(exp.netAmount)||0);
      sheet.getRange(row,18).setValue(exp.payment||'');
      const slipStr = Array.isArray(exp.slipUrls) ? exp.slipUrls.filter(Boolean).join(',') : (exp.slipUrls||'');
      const rcptStr = Array.isArray(exp.rcptUrls) ? exp.rcptUrls.filter(Boolean).join(',') : (exp.rcptUrls||'');
      sheet.getRange(row,19).setValue(slipStr);
      sheet.getRange(row,20).setValue(rcptStr);
      sheet.getRange(row,21).setValue(exp.note||'');
      // ── header-based สำหรับ payer / updatedAt / deletedAt (ป้องกัน schema ไม่ตรง) ──
      const payerCol   = getColIndex(headers, 'payer');
      const updCol     = getColIndex(headers, 'updatedAt');
      const delCol     = getColIndex(headers, 'deletedAt');
      if (payerCol > 0) sheet.getRange(row, payerCol).setValue(exp.payer||'');
      // addedBy — ล็อคไว้ ไม่อัปเดต (เจ้าของรายการไม่เปลี่ยน)
      if (updCol   > 0) sheet.getRange(row, updCol).setValue(now);
      if (delCol   > 0) sheet.getRange(row, delCol).setValue(''); // clear soft-delete flag
      return { success: true, updatedAt: now };
    }
  }
  return { success: false, error: 'Expense not found' };
}

// ── SOFT DELETE (v3) ── ไม่ลบ row จริง แค่ set deletedAt
function deleteExpense(ss, id) {
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet) return { success: false, error: 'Expenses sheet not found' };
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const row    = i + 1;
      const now    = new Date().toISOString();
      const updCol = getColIndex(headers, 'updatedAt');
      const delCol = getColIndex(headers, 'deletedAt');
      if (updCol > 0) sheet.getRange(row, updCol).setValue(now);
      if (delCol > 0) sheet.getRange(row, delCol).setValue(now);
      // ถ้า column ยังไม่มี (sheet เก่า) — ใช้ hard delete เป็น fallback
      if (updCol < 0 || delCol < 0) {
        sheet.deleteRow(row);
      }
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

// ── อัปเดตแถวใน Sheet โดย match id ──
function updateSheetRow(ss, sheetName, row) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: sheetName + ' sheet not found' };
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(row.id)) {
      headers.forEach((h, col) => {
        if (row[h] !== undefined) {
          sheet.getRange(i + 1, col + 1).setValue(row[h]);
        }
      });
      return { success: true };
    }
  }
  return { success: false, error: 'Row not found: ' + row.id };
}

function deleteProject(ss, id) {
  const sheet = ss.getSheetByName('Projects');
  if (!sheet) return { success: false, error: 'Projects sheet not found' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Project not found: ' + id };
}

function deleteActivity(ss, id) {
  const sheet = ss.getSheetByName('Activities');
  if (!sheet) return { success: false, error: 'Activities sheet not found' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Activity not found: ' + id };
}

// ─────────────────────────────────────
//  IMAGE MANAGEMENT (Google Drive)
// ─────────────────────────────────────
function deleteImage(fileId) {
  try {
    if (!fileId) return { success: false, error: 'No fileId' };
    DriveApp.getFileById(fileId).setTrashed(true);
    return { success: true, fileId: fileId };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteActivityFolder(payload) {
  try {
    const { projectId, projectCode, projectName, activityId, activityCode, activityName } = payload;
    const rootFolder = getRootUploadFolder();
    const projName   = buildProjectFolderName(projectId, projectCode, projectName);
    const projIter   = rootFolder.getFoldersByName(safeFolderName(projName));
    if (!projIter.hasNext()) return { success: false, error: 'Project folder not found: ' + projName };
    const projFolder = projIter.next();
    const actName    = buildActivityFolderName(activityId, activityName, activityCode);
    const actIter    = projFolder.getFoldersByName(safeFolderName(actName));
    if (!actIter.hasNext()) return { success: false, error: 'Activity folder not found: ' + actName };
    actIter.next().setTrashed(true);
    return { success: true, folder: actName };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteProjectFolder(payload) {
  try {
    const { projectId, projectCode, projectName } = payload;
    const rootFolder = getRootUploadFolder();
    const projName   = buildProjectFolderName(projectId, projectCode, projectName);
    const iter       = rootFolder.getFoldersByName(safeFolderName(projName));
    if (!iter.hasNext()) return { success: false, error: 'Project folder not found: ' + projName };
    iter.next().setTrashed(true);
    return { success: true, folder: projName };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function safeFolderName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_').trim().substring(0, 60);
}

function moveFile(payload) {
  try {
    const { fileId, projectId, projectCode, projectName,
            activityId, activityCode, activityName, expenseDate } = payload;
    const file         = DriveApp.getFileById(fileId);
    const rootFolder   = getRootUploadFolder();
    const projFolder   = getOrCreateSubFolder(rootFolder, buildProjectFolderName(projectId, projectCode, projectName));
    let parentFolder   = projFolder;
    if (activityId || activityCode || activityName) {
      parentFolder = getOrCreateSubFolder(projFolder, buildActivityFolderName(activityId, activityName, activityCode));
    }
    const dayFolder  = getOrCreateSubFolder(parentFolder, getDateFolder(expenseDate));
    const oldParents = file.getParents();
    dayFolder.addFile(file);
    while (oldParents.hasNext()) {
      const old = oldParents.next();
      if (old.getId() !== dayFolder.getId()) old.removeFile(file);
    }
    return { success: true, fileId: file.getId() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function uploadImage(payload) {
  try {
    const { imageBase64, mimeType, filename, fileType,
            projectId, projectName, projectCode,
            activityId, activityCode, activityName, expenseDate } = payload;

    const rootFolder     = getRootUploadFolder();
    const projFolder     = getOrCreateSubFolder(rootFolder, buildProjectFolderName(projectId, projectCode, projectName));
    let parentFolder     = projFolder;
    let actFolderName    = '';
    if (activityId || activityCode || activityName) {
      actFolderName  = buildActivityFolderName(activityId, activityName, activityCode);
      parentFolder   = getOrCreateSubFolder(projFolder, actFolderName);
    }
    const dayFolder  = getOrCreateSubFolder(parentFolder, getDateFolder(expenseDate));
    const ext        = getExtension(mimeType, filename);
    const ts         = getTimestamp(expenseDate);
    const shortId    = (projectCode || projectId || 'X').replace(/[^A-Za-z0-9]/g, '').substring(0, 8);
    const prefix     = (fileType === 'rcpt') ? 'rcpt' : 'slip';
    const finalName  = prefix + '_' + ts + '_' + shortId + '.' + ext;
    const b64        = imageBase64.includes(',') ? imageBase64.split(',')[1] : imageBase64;
    const decoded    = Utilities.base64Decode(b64);
    const blob       = Utilities.newBlob(decoded, mimeType || 'image/jpeg', finalName);
    const file       = dayFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success:  true,
      url:      'https://drive.google.com/uc?id=' + file.getId(),
      fileId:   file.getId(),
      fileName: finalName,
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────
//  CLOUD BACKUP
// ─────────────────────────────────────
function saveCloudBackup(payload) {
  try {
    const root   = getRootUploadFolder();
    const folder = getOrCreateSubFolder(root, '_backups');
    const iter   = folder.getFilesByName('backups.json');
    while (iter.hasNext()) iter.next().setTrashed(true);
    folder.createFile('backups.json', payload.data, 'application/json');
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function loadCloudBackup() {
  try {
    const root   = getRootUploadFolder();
    const folder = getOrCreateSubFolder(root, '_backups');
    const iter   = folder.getFilesByName('backups.json');
    if (!iter.hasNext()) return { success: true, backups: [] };
    return { success: true, backups: JSON.parse(iter.next().getBlob().getDataAsString()) };
  } catch (err) {
    return { success: false, error: err.toString(), backups: [] };
  }
}

// ─────────────────────────────────────
//  ADD ROW (generic)
// ─────────────────────────────────────
function addRow(ss, sheetName, rowObj) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: sheetName + ' sheet not found' };
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headers.length === 0) return { success: false, error: 'No headers in ' + sheetName };
  const row = headers.map(h => (rowObj[h] !== undefined ? rowObj[h] : ''));
  sheet.appendRow(row);
  return { success: true };
}

// ─────────────────────────────────────
//  CATEGORIES CRUD
// ─────────────────────────────────────
function deleteCategoryRow(ss, cat) {
  const sheet = ss.getSheetByName('Categories');
  if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'No data' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const iL1 = headers.indexOf('l1');
  const iL2 = headers.indexOf('l2');
  const iL3 = headers.indexOf('l3');
  if (iL1 < 0) return { success: false, error: 'l1 column not found' };
  for (let i = data.length - 1; i >= 1; i--) {
    const rowL1 = String(data[i][iL1] || '').trim();
    const rowL2 = iL2 >= 0 ? String(data[i][iL2] || '').trim() : '';
    const rowL3 = iL3 >= 0 ? String(data[i][iL3] || '').trim() : '';
    if (rowL1 === (cat.l1 || '') && rowL2 === (cat.l2 || '') && rowL3 === (cat.l3 || '')) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Category not found' };
}

function deleteCategoryGroup(ss, l1) {
  const sheet = ss.getSheetByName('Categories');
  if (!sheet || sheet.getLastRow() < 2) return { success: true, deleted: 0 };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const iL1 = headers.indexOf('l1');
  if (iL1 < 0) return { success: false, error: 'l1 column not found' };
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][iL1] || '').trim() === String(l1 || '').trim()) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return { success: true, deleted };
}

// ─────────────────────────────────────
//  PURGE DELETED EXPENSES
// ─────────────────────────────────────
function purgeDeleted(ss, olderThanDays) {
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet || sheet.getLastRow() < 2) return { success: true, purged: 0 };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const iDel = headers.indexOf('deletedAt');
  if (iDel < 0) return { success: false, error: 'deletedAt column not found' };
  const days = parseInt(olderThanDays);
  // days === 0 หรือ NaN = ลบทุก row ที่มี deletedAt (ไม่จำกัดอายุ)
  const noLimit = isNaN(days) || days <= 0;
  const cutoff = noLimit ? null : new Date(Date.now() - days * 86400000);
  let purged = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const delVal = String(data[i][iDel] || '').trim();
    if (!delVal) continue;
    if (noLimit) {
      sheet.deleteRow(i + 1);
      purged++;
    } else {
      const delDate = new Date(delVal);
      if (!isNaN(delDate.getTime()) && delDate.getTime() < cutoff.getTime()) {
        sheet.deleteRow(i + 1);
        purged++;
      }
    }
  }
  return { success: true, purged };
}

// ─────────────────────────────────────
//  REPAIR / DIAGNOSE EXPENSE HEADERS
// ─────────────────────────────────────
function diagnoseExpenseHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet) { Logger.log('Expenses sheet not found'); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Headers: ' + JSON.stringify(headers));
  const required = ['id','date','projectId','projectName','activityId','activityName',
    'catL1','catL2','catL3','description','amount','vat','wht','netAmount',
    'payer','addedBy','updatedAt','deletedAt'];
  const missing = required.filter(h => !headers.includes(h));
  if (missing.length === 0) Logger.log('All required headers present');
  else Logger.log('Missing headers: ' + JSON.stringify(missing));
}

function repairExpenseHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets(ss);
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const iPayer   = headers.indexOf('payer');
  const iAddedBy = headers.indexOf('addedBy');
  if (iPayer < 0 || iAddedBy < 0) { Logger.log('payer/addedBy columns not found'); return; }
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const payer   = String(data[i][iPayer]   || '').trim();
    const addedBy = String(data[i][iAddedBy] || '').trim();
    if (!payer && addedBy) { sheet.getRange(i + 1, iPayer + 1).setValue(addedBy); fixed++; }
    else if (!addedBy && payer) { sheet.getRange(i + 1, iAddedBy + 1).setValue(payer); fixed++; }
  }
  Logger.log('repairExpenseHeaders: fixed ' + fixed + ' rows');
}
