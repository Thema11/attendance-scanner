// ============================================================
//  نظام تسجيل الحضور بـ QR Code — النيابة العامة
//  الملف: Code.gs  |  Backend الرئيسي
// ============================================================

const SHEET_ID       = 'Y1xb4nVpmQBSf_OSBEgdHLLvWqa9ddei2IvhuwRTgITJU';
const LATE_THRESHOLD = 15;

// ── أسماء الأوراق (عربي) ────────────────────────────────────
const SH = {
  WORKSHOPS : 'الورش',
  QR_CODES  : 'رموز QR',
  LOG       : 'سجل الحضور',
  STATS     : 'الإحصائيات',
};

// ── ألوان النيابة ────────────────────────────────────────────
const CLR = {
  TEAL      : '#02807A',
  TEAL_LIGHT: '#309C99',
  GOLD      : '#AA926D',
  GOLD_LIGHT: '#C8B08A',
  WHITE     : '#FFFFFF',
  LIGHT_BG  : '#F0FAF9',
  GREEN_BG  : '#D9F2E6',
  GREEN_TXT : '#0D5C2E',
  AMBER_BG  : '#FFF3CD',
  AMBER_TXT : '#7B4F00',
  PURPLE_BG : '#EDE9FE',
  PURPLE_TXT: '#4C1D95',
  RED_BG    : '#FEE2E2',
  RED_TXT   : '#7F1D1D',
  GRAY_BG   : '#F3F4F6',
  GRAY_TXT  : '#374151',
  HEADER_TXT: '#FFFFFF',
};

// ── نقطة دخول GET ───────────────────────────────────────────
// تخدم صفحات Apps Script العادية
// وتقبل أيضاً طلبات API من GitHub Pages عبر ?data=
function doGet(e) {
  // إذا جاء طلب API من GitHub Pages
  if (e.parameter.data) {
    try {
      const payload = JSON.parse(decodeURIComponent(e.parameter.data));
      const handlers = {
        validateQR      : () => validateQR(payload),
        registerAttendee: () => registerAttendee(payload),
        getDashboard    : () => getDashboardData(payload),
        createWorkshop  : () => createWorkshop(payload),
        createQR        : () => createQRCode(payload),
        toggleQR        : () => toggleQR(payload),
        getWorkshops    : () => getWorkshops(),
        getQRList       : () => getQRList(payload),
        updateWorkshop  : () => updateWorkshop(payload),
        archiveWorkshop : () => archiveWorkshop(payload),
        deleteWorkshop  : () => deleteWorkshop(payload),
        deleteQR        : () => deleteQR(payload),
      };
      const fn = handlers[payload.action];
      if (!fn) return jsonResponse({ ok: false, error: 'إجراء غير معروف' });
      return jsonResponse(fn());
    } catch(err) {
      return jsonResponse({ ok: false, error: err.message });
    }
  }

  // طلب صفحة عادية
  const page     = e.parameter.page || 'scan';
  const template = HtmlService.createTemplateFromFile(page);
  template.SCRIPT_URL = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('نظام الحضور — النيابة العامة')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── نقطة دخول POST ──────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const handlers = {
      validateQR      : () => validateQR(payload),
      registerAttendee: () => registerAttendee(payload),
      getDashboard    : () => getDashboardData(payload),
      createWorkshop  : () => createWorkshop(payload),
      createQR        : () => createQRCode(payload),
      toggleQR        : () => toggleQR(payload),
      getWorkshops    : () => getWorkshops(),
      getQRList       : () => getQRList(payload),
      updateWorkshop  : () => updateWorkshop(payload),
      archiveWorkshop : () => archiveWorkshop(payload),
      deleteWorkshop  : () => deleteWorkshop(payload),
      deleteQR        : () => deleteQR(payload),
    };
    const fn = handlers[payload.action];
    if (!fn) return jsonResponse({ ok: false, error: 'إجراء غير معروف' });
    return jsonResponse(fn());
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
//  1. التحقق من QR
// ============================================================
function validateQR({ token }) {
  if (!token) return { ok: false, reason: 'لا يوجد رمز' };
  const rows = getSheet(SH.QR_CODES).getDataRange().getValues();
  const idx  = makeIdx(rows[0]);
  const now  = new Date();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[idx['الرمز السري']] !== token) continue;
    if (!row[idx['مفعّل']]) return { ok: false, reason: 'الرمز غير مفعَّل' };
    const from  = new Date(row[idx['صالح من']]);
    const until = new Date(row[idx['صالح حتى']]);
    if (now < from)  return { ok: false, reason: 'لم يبدأ وقت التسجيل بعد' };
    if (now > until) return { ok: false, reason: 'انتهت صلاحية الرمز' };
    return {
      ok           : true,
      qr_id        : row[idx['معرف QR']],
      workshop_id  : row[idx['معرف الورشة']],
      day_number   : row[idx['رقم اليوم']],
      event_date   : Utilities.formatDate(new Date(row[idx['تاريخ اليوم']]), 'Asia/Riyadh', 'yyyy-MM-dd'),
      workshop_name: getWorkshopName(row[idx['معرف الورشة']]),
      late_after   : row[idx['متأخر بعد']] || '',
    };
  }
  return { ok: false, reason: 'رمز غير معروف' };
}

// ============================================================
//  2. تسجيل الحضور
// ============================================================
function registerAttendee({ qr_id, workshop_id, day_number, event_date,
                            attendee_name, approved_by, override, override_reason }) {
  if (!attendee_name || !attendee_name.trim())
    return { ok: false, error: 'الاسم مطلوب' };

  const name     = attendee_name.trim();
  const logSheet = getSheet(SH.LOG);
  const rows     = logSheet.getDataRange().getValues();
  const idx      = makeIdx(rows[0]);
  const now      = new Date();
  const arrTime  = Utilities.formatDate(now, 'Asia/Riyadh', 'HH:mm:ss');
  const today    = Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy-MM-dd');

  // فحص التكرار
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[idx['الاسم الكامل']].toString().toLowerCase().trim() === name.toLowerCase()
        && r[idx['معرف QR']] === qr_id
        && (r[idx['الحالة']] === 'حاضر' || r[idx['الحالة']] === 'تجاوز')) {
      if (!override) return {
        ok: false, duplicate: true,
        registered_at: r[idx['وقت الوصول']],
        log_id: r[idx['معرف السجل']],
      };
      break;
    }
  }

  // تحديد التأخر
  let is_late = false;
  const qrRows = getSheet(SH.QR_CODES).getDataRange().getValues();
  const qi = makeIdx(qrRows[0]);
  for (let i = 1; i < qrRows.length; i++) {
    if (qrRows[i][qi['معرف QR']] === qr_id) {
      const la = qrRows[i][qi['متأخر بعد']];
      if (la) is_late = now > new Date(la);
      break;
    }
  }

  const log_id = 'LOG-' + now.getTime();
  const status  = override ? 'تجاوز' : 'حاضر';

  const newRow = new Array(rows[0].length).fill('');
  newRow[idx['معرف السجل']]    = log_id;
  newRow[idx['معرف الورشة']]   = workshop_id;
  newRow[idx['معرف QR']]       = qr_id;
  newRow[idx['الاسم الكامل']]  = name;
  newRow[idx['تاريخ الحضور']]  = event_date || today;
  newRow[idx['وقت الوصول']]    = arrTime;
  newRow[idx['رقم اليوم']]     = day_number;
  newRow[idx['الحالة']]        = status;
  newRow[idx['متأخر']]         = is_late ? 'نعم' : 'لا';
  newRow[idx['اعتمد من']]      = approved_by || '';
  newRow[idx['سبب التجاوز']]   = override_reason || '';

  const newRowIndex = logSheet.getLastRow() + 1;
  logSheet.appendRow(newRow);

  // تلوين الصف حسب الحالة
  colorLogRow(logSheet, newRowIndex, status, is_late);

  return { ok: true, log_id, arrival_time: arrTime, is_late, status };
}

// تلوين صف الحضور
function colorLogRow(sheet, rowIndex, status, is_late) {
  const numCols = sheet.getLastColumn();
  const range   = sheet.getRange(rowIndex, 1, 1, numCols);
  if (status === 'تجاوز') {
    range.setBackground(CLR.PURPLE_BG);
    range.setFontColor(CLR.PURPLE_TXT);
  } else if (is_late) {
    range.setBackground(CLR.AMBER_BG);
    range.setFontColor(CLR.AMBER_TXT);
  } else {
    range.setBackground(CLR.GREEN_BG);
    range.setFontColor(CLR.GREEN_TXT);
  }
}

// ============================================================
//  3. بيانات Dashboard
// ============================================================
function getDashboardData({ workshop_id, qr_id }) {
  const rows    = getSheet(SH.LOG).getDataRange().getValues();
  const idx     = makeIdx(rows[0]);
  const records = rows.slice(1).filter(r =>
    r[idx['معرف الورشة']] === workshop_id &&
    r[idx['معرف QR']]     === qr_id &&
    (r[idx['الحالة']] === 'حاضر' || r[idx['الحالة']] === 'تجاوز')
  );

  const total     = records.length;
  const lateCount = records.filter(r => r[idx['متأخر']] === 'نعم').length;
  const last      = total > 0 ? records[records.length-1][idx['الاسم الكامل']] : '—';

  return {
    ok: true,
    kpis: {
      total, lateCount, lastPerson: last,
      onTime : total - lateCount,
      avgTime: calcAverageTime(records, idx),
    },
    distribution : buildTimeDistribution(records, idx),
    earliest     : total > 0 ? { name: records[0][idx['الاسم الكامل']], time: records[0][idx['وقت الوصول']] } : null,
    table        : records.slice(-50).reverse().map(r => ({
      name        : r[idx['الاسم الكامل']],
      arrival_time: r[idx['وقت الوصول']],
      status      : r[idx['الحالة']],
      is_late     : r[idx['متأخر']] === 'نعم',
      date        : r[idx['تاريخ الحضور']],
    })),
  };
}

function buildTimeDistribution(records, idx) {
  const b = {};
  records.forEach(r => {
    const t = r[idx['وقت الوصول']].toString();
    if (!t) return;
    const [h, m] = t.split(':').map(Number);
    const slot = `${String(h).padStart(2,'0')}:${m<15?'00':m<30?'15':m<45?'30':'45'}`;
    b[slot] = (b[slot] || 0) + 1;
  });
  return Object.entries(b).sort((a,b)=>a[0].localeCompare(b[0])).map(([time,count])=>({time,count}));
}

function calcAverageTime(records, idx) {
  if (!records.length) return '—';
  const total = records.reduce((s, r) => {
    const t = r[idx['وقت الوصول']].toString();
    if (!t) return s;
    const [h,m] = t.split(':').map(Number);
    return s + h*60 + m;
  }, 0);
  const avg = Math.round(total / records.length);
  return `${String(Math.floor(avg/60)).padStart(2,'0')}:${String(avg%60).padStart(2,'0')}`;
}

// ============================================================
//  4. إدارة الورش و QR
// ============================================================
function createWorkshop({ name, start_date, end_date, location }) {
  if (!name) return { ok: false, error: 'اسم الورشة مطلوب' };
  const ws_id = 'WS-' + new Date().getTime();
  const sheet = getSheet(SH.WORKSHOPS);
  const row   = sheet.getLastRow() + 1;
  sheet.appendRow([ws_id, name, start_date, end_date, location || '', 'نشطة']);
  // تلوين الصف
  sheet.getRange(row, 1, 1, 6).setBackground(CLR.LIGHT_BG);
  updateStatsSheet();
  return { ok: true, workshop_id: ws_id };
}

function createQRCode({ workshop_id, day_number, event_date, valid_from, valid_until, late_after }) {
  if (!workshop_id || !day_number || !event_date)
    return { ok: false, error: 'بيانات ناقصة' };
  const qr_id = 'QR-' + workshop_id + '-D' + day_number;
  const token = Utilities.getUuid();
  const url   = ScriptApp.getService().getUrl() + '?page=scan&token=' + token;
  const sheet = getSheet(SH.QR_CODES);
  const row   = sheet.getLastRow() + 1;
  sheet.appendRow([qr_id, workshop_id, day_number, event_date, token, url,
                   valid_from||'', valid_until||'', late_after||'', true]);
  sheet.getRange(row, 1, 1, 10).setBackground(CLR.LIGHT_BG);
  // قائمة منسدلة لعمود مفعّل
  const activeCell = sheet.getRange(row, 10);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE','FALSE'], true).build();
  activeCell.setDataValidation(rule);
  return { ok: true, qr_id, token, url };
}

function toggleQR({ qr_id, is_active }) {
  const sheet = getSheet(SH.QR_CODES);
  const rows  = sheet.getDataRange().getValues();
  const idx   = makeIdx(rows[0]);
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idx['معرف QR']] === qr_id) {
      sheet.getRange(i+1, idx['مفعّل']+1).setValue(is_active);
      const bg = is_active ? CLR.LIGHT_BG : CLR.GRAY_BG;
      sheet.getRange(i+1, 1, 1, rows[0].length).setBackground(bg);
      return { ok: true };
    }
  }
  return { ok: false, error: 'QR غير موجود' };
}

function getWorkshops() {
  const rows = getSheet(SH.WORKSHOPS).getDataRange().getValues();
  const idx  = makeIdx(rows[0]);
  return {
    ok: true,
    workshops: rows.slice(1).map(r => ({
      workshop_id: r[idx['معرف الورشة']], name      : r[idx['اسم الورشة']],
      start_date : r[idx['تاريخ البداية']], end_date: r[idx['تاريخ النهاية']],
      location   : r[idx['الموقع']],        status  : r[idx['الحالة']],
    }))
  };
}

function getQRList({ workshop_id }) {
  const rows = getSheet(SH.QR_CODES).getDataRange().getValues();
  const idx  = makeIdx(rows[0]);
  return {
    ok: true,
    qr_list: rows.slice(1)
      .filter(r => !workshop_id || r[idx['معرف الورشة']] === workshop_id)
      .map(r => ({
        qr_id      : r[idx['معرف QR']],        workshop_id: r[idx['معرف الورشة']],
        day_number : r[idx['رقم اليوم']],       event_date : r[idx['تاريخ اليوم']],
        url        : r[idx['الرابط']],           valid_from : r[idx['صالح من']],
        valid_until: r[idx['صالح حتى']],        is_active  : r[idx['مفعّل']],
      }))
  };
}

// ============================================================
//  5. إعداد Sheets الاحترافية — تُشغَّل مرة واحدة
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── 1. ورقة الورش ─────────────────────────────────────────
  const wsSheet = getOrCreate(ss, SH.WORKSHOPS);
  if (wsSheet.getLastRow() === 0) {
    const wsHeaders = ['معرف الورشة','اسم الورشة','تاريخ البداية','تاريخ النهاية','الموقع','الحالة'];
    wsSheet.appendRow(wsHeaders);
    styleHeader(wsSheet, wsHeaders.length, CLR.TEAL);
    wsSheet.setColumnWidth(1, 160);
    wsSheet.setColumnWidth(2, 220);
    wsSheet.setColumnWidth(3, 130);
    wsSheet.setColumnWidth(4, 130);
    wsSheet.setColumnWidth(5, 180);
    wsSheet.setColumnWidth(6, 100);
    // قائمة منسدلة للحالة
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['نشطة','مغلقة'], true).build();
    wsSheet.getRange('F2:F1000').setDataValidation(statusRule);
    // تنسيق شرطي
    addConditionalFormatting(wsSheet, 'F2:F1000', 'نشطة',  CLR.GREEN_BG,  CLR.GREEN_TXT);
    addConditionalFormatting(wsSheet, 'F2:F1000', 'مغلقة', CLR.GRAY_BG,   CLR.GRAY_TXT);
    wsSheet.setTabColor(CLR.TEAL);
    wsSheet.setFrozenRows(1);
    wsSheet.setRightToLeft(true);
  }

  // ── 2. ورقة QR ────────────────────────────────────────────
  const qrSheet = getOrCreate(ss, SH.QR_CODES);
  if (qrSheet.getLastRow() === 0) {
    const qrHeaders = ['معرف QR','معرف الورشة','رقم اليوم','تاريخ اليوم',
                       'الرمز السري','الرابط','صالح من','صالح حتى','متأخر بعد','مفعّل'];
    qrSheet.appendRow(qrHeaders);
    styleHeader(qrSheet, qrHeaders.length, CLR.GOLD);
    qrSheet.setColumnWidth(1, 160);
    qrSheet.setColumnWidth(2, 160);
    qrSheet.setColumnWidth(3, 90);
    qrSheet.setColumnWidth(4, 120);
    qrSheet.setColumnWidth(5, 280);
    qrSheet.setColumnWidth(6, 300);
    qrSheet.setColumnWidth(7, 140);
    qrSheet.setColumnWidth(8, 140);
    qrSheet.setColumnWidth(9, 140);
    qrSheet.setColumnWidth(10, 80);
    // قائمة مفعّل
    const boolRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE','FALSE'], true).build();
    qrSheet.getRange('J2:J1000').setDataValidation(boolRule);
    addConditionalFormatting(qrSheet, 'J2:J1000', 'TRUE',  CLR.GREEN_BG, CLR.GREEN_TXT);
    addConditionalFormatting(qrSheet, 'J2:J1000', 'FALSE', CLR.RED_BG,   CLR.RED_TXT);
    qrSheet.setTabColor(CLR.GOLD);
    qrSheet.setFrozenRows(1);
    qrSheet.setRightToLeft(true);
  }

  // ── 3. ورقة سجل الحضور ────────────────────────────────────
  const logSheet = getOrCreate(ss, SH.LOG);
  if (logSheet.getLastRow() === 0) {
    const logHeaders = ['معرف السجل','معرف الورشة','معرف QR','الاسم الكامل',
                        'تاريخ الحضور','وقت الوصول','رقم اليوم',
                        'الحالة','متأخر','اعتمد من','سبب التجاوز'];
    logSheet.appendRow(logHeaders);
    styleHeader(logSheet, logHeaders.length, CLR.TEAL);
    logSheet.setColumnWidth(1, 160);
    logSheet.setColumnWidth(2, 140);
    logSheet.setColumnWidth(3, 160);
    logSheet.setColumnWidth(4, 200);
    logSheet.setColumnWidth(5, 120);
    logSheet.setColumnWidth(6, 110);
    logSheet.setColumnWidth(7, 90);
    logSheet.setColumnWidth(8, 100);
    logSheet.setColumnWidth(9, 80);
    logSheet.setColumnWidth(10, 120);
    logSheet.setColumnWidth(11, 180);
    // قوائم منسدلة
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['حاضر','تجاوز'], true).build();
    logSheet.getRange('H2:H5000').setDataValidation(statusRule);
    const lateRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['نعم','لا'], true).build();
    logSheet.getRange('I2:I5000').setDataValidation(lateRule);
    // تنسيق شرطي
    addConditionalFormatting(logSheet, 'H2:H5000', 'حاضر',  CLR.GREEN_BG,  CLR.GREEN_TXT);
    addConditionalFormatting(logSheet, 'H2:H5000', 'تجاوز', CLR.PURPLE_BG, CLR.PURPLE_TXT);
    addConditionalFormatting(logSheet, 'I2:I5000', 'نعم',   CLR.AMBER_BG,  CLR.AMBER_TXT);
    logSheet.setTabColor(CLR.TEAL_LIGHT);
    logSheet.setFrozenRows(1);
    logSheet.setRightToLeft(true);
  }

  // ── 4. ورقة الإحصائيات ────────────────────────────────────
  buildStatsSheet(ss);

  Logger.log('✅ تم إعداد قاعدة البيانات بنجاح!');
}

// ── بناء ورقة الإحصائيات ─────────────────────────────────────
function buildStatsSheet(ss) {
  let st = ss.getSheetByName(SH.STATS);
  if (st) ss.deleteSheet(st);
  st = ss.insertSheet(SH.STATS);
  st.setRightToLeft(true);
  st.setTabColor(CLR.GOLD);

  // ── عنوان رئيسي ─────────────────────────────────────────
  st.getRange('A1:F1').merge()
    .setValue('📊 لوحة الإحصائيات — نظام الحضور — النيابة العامة')
    .setBackground(CLR.TEAL).setFontColor(CLR.WHITE)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  st.setRowHeight(1, 44);

  // ── مؤشرات KPI ──────────────────────────────────────────
  const kpiRow = 3;
  const kpis = [
    ['إجمالي الحضور',        `=COUNTA('سجل الحضور'!D2:D5000)`,              CLR.TEAL],
    ['الحضور في الوقت',      `=COUNTIF('سجل الحضور'!I2:I5000,"لا")`,        CLR.TEAL_LIGHT],
    ['المتأخرون',            `=COUNTIF('سجل الحضور'!I2:I5000,"نعم")`,       CLR.GOLD],
    ['حالات التجاوز',        `=COUNTIF('سجل الحضور'!H2:H5000,"تجاوز")`,    '#8B5CF6'],
    ['إجمالي الورش',         `=COUNTA('الورش'!A2:A1000)`,                   CLR.TEAL],
    ['رموز QR المفعّلة',     `=COUNTIF('رموز QR'!J2:J1000,"TRUE")`,         CLR.GOLD],
  ];

  kpis.forEach(([label, formula, color], i) => {
    const col = i + 1;
    // الرقم
    st.getRange(kpiRow, col)
      .setFormula(formula)
      .setFontSize(28).setFontWeight('bold')
      .setFontColor(color)
      .setHorizontalAlignment('center')
      .setBackground(CLR.LIGHT_BG);
    // التسمية
    st.getRange(kpiRow + 1, col)
      .setValue(label)
      .setFontSize(10).setFontColor(CLR.GRAY_TXT)
      .setHorizontalAlignment('center')
      .setBackground(CLR.LIGHT_BG);
    // تنسيق الخلية
    st.getRange(kpiRow, col, 2, 1)
      .setBorder(true, true, true, true, false, false,
                 CLR.TEAL_LIGHT, SpreadsheetApp.BorderStyle.SOLID);
    st.setRowHeight(kpiRow, 52);
    st.setRowHeight(kpiRow+1, 30);
    st.setColumnWidth(col, 140);
  });

  // ── فراغ ────────────────────────────────────────────────
  st.setRowHeight(6, 20);

  // ── جدول توزيع الحضور حسب الحالة ───────────────────────
  const tblRow = 7;
  const tblHeaders = ['الحالة', 'العدد', 'النسبة'];
  tblHeaders.forEach((h, i) => {
    st.getRange(tblRow, i+1).setValue(h)
      .setBackground(CLR.GOLD).setFontColor(CLR.WHITE)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });
  const tblData = [
    ['حاضر',  `=COUNTIF('سجل الحضور'!H2:H5000,"حاضر")`,  `=IF(B8=0,"—",TEXT(B8/B10,"0%"))`],
    ['تجاوز', `=COUNTIF('سجل الحضور'!H2:H5000,"تجاوز")`, `=IF(B9=0,"—",TEXT(B9/B10,"0%"))`],
    ['المجموع',`=B8+B9`,                                   '"100%"'],
  ];
  tblData.forEach(([label, formula, pct], i) => {
    const r = tblRow + 1 + i;
    st.getRange(r, 1).setValue(label);
    st.getRange(r, 2).setFormula(formula);
    st.getRange(r, 3).setFormula ? st.getRange(r, 3).setFormula(pct) : st.getRange(r, 3).setValue(pct);
    const bg = i === 0 ? CLR.GREEN_BG : i === 1 ? CLR.PURPLE_BG : CLR.LIGHT_BG;
    const fg = i === 0 ? CLR.GREEN_TXT : i === 1 ? CLR.PURPLE_TXT : CLR.TEAL;
    st.getRange(r, 1, 1, 3).setBackground(bg).setFontColor(fg);
  });
  st.getRange(tblRow, 1, 4, 3)
    .setBorder(true, true, true, true, true, true,
               CLR.BORDER || '#C0C0C0', SpreadsheetApp.BorderStyle.SOLID);

  // ── جدول التأخر ──────────────────────────────────────────
  const t2Row = 7;
  ['الوضع','العدد'].forEach((h, i) => {
    st.getRange(t2Row, i+4).setValue(h)
      .setBackground(CLR.TEAL).setFontColor(CLR.WHITE)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });
  [
    ['في الوقت',  `=COUNTIF('سجل الحضور'!I2:I5000,"لا")`],
    ['متأخر',     `=COUNTIF('سجل الحضور'!I2:I5000,"نعم")`],
  ].forEach(([label, formula], i) => {
    const r = t2Row + 1 + i;
    st.getRange(r, 4).setValue(label);
    st.getRange(r, 5).setFormula(formula);
    const bg = i === 0 ? CLR.GREEN_BG : CLR.AMBER_BG;
    const fg = i === 0 ? CLR.GREEN_TXT : CLR.AMBER_TXT;
    st.getRange(r, 4, 1, 2).setBackground(bg).setFontColor(fg);
  });
  st.getRange(t2Row, 4, 3, 2)
    .setBorder(true, true, true, true, true, true,
               '#C0C0C0', SpreadsheetApp.BorderStyle.SOLID);

  // ── فراغ ────────────────────────────────────────────────
  st.setRowHeight(12, 20);

  // ── آخر 10 سجلات ─────────────────────────────────────────
  const lastRow = 13;
  st.getRange(lastRow, 1, 1, 5).merge()
    .setValue('🕐 آخر السجلات المضافة')
    .setBackground(CLR.TEAL).setFontColor(CLR.WHITE)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('right');
  ['الاسم','تاريخ الحضور','وقت الوصول','الحالة','متأخر'].forEach((h, i) => {
    st.getRange(lastRow+1, i+1).setValue(h)
      .setBackground(CLR.TEAL_LIGHT).setFontColor(CLR.WHITE)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });
  // معادلات ARRAY_FORMULA لجلب آخر 10 سجلات
  const cols = ['D','E','F','H','I'];
  cols.forEach((col, i) => {
    st.getRange(lastRow+2, i+1)
      .setFormula(`=IFERROR(INDEX('سجل الحضور'!${col}:${col},COUNTA('سجل الحضور'!${col}:${col})-0),"")`);
  });
  st.getRange(lastRow, 1, 12, 5)
    .setBorder(true, true, true, true, false, true,
               '#C0C0C0', SpreadsheetApp.BorderStyle.SOLID);

  // ── تجميد وإخفاء الشبكة ────────────────────────────────
  st.setFrozenRows(1);
  st.setHiddenGridlines(true);
}

// تحديث ورقة الإحصائيات عند إضافة بيانات
function updateStatsSheet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    buildStatsSheet(ss);
  } catch(e) { Logger.log('updateStats error: ' + e.message); }
}

// ============================================================
//  أدوات مساعدة
// ============================================================
function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader(sheet, numCols, bgColor) {
  const hdr = sheet.getRange(1, 1, 1, numCols);
  hdr.setBackground(bgColor)
     .setFontColor(CLR.WHITE)
     .setFontWeight('bold')
     .setFontSize(11)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
}

function addConditionalFormatting(sheet, range, value, bg, fg) {
  const r    = sheet.getRange(range);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(value)
    .setBackground(bg)
    .setFontColor(fg)
    .setRanges([r])
    .build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}
function makeIdx(headers) {
  const map = {};
  headers.forEach((h, i) => { map[h] = i; });
  return map;
}
function getWorkshopName(workshop_id) {
  const rows = getSheet(SH.WORKSHOPS).getDataRange().getValues();
  const idx  = makeIdx(rows[0]);
  for (let i = 1; i < rows.length; i++)
    if (rows[i][idx['معرف الورشة']] === workshop_id) return rows[i][idx['اسم الورشة']];
  return workshop_id;
}
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
                       .setMimeType(ContentService.MimeType.JSON);
}
