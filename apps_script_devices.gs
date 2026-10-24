/**
 * حد الأجهزة: ٤ أجهزة لكل طالب، وعند الجهاز الخامس يتوقف الحساب لحين إعادة تشغيله من لوحة المعلم.
 *
 * طريقة التركيب (٣ خطوات):
 *  1) ألصق هذا الملف كاملاً كملف جديد داخل مشروع Apps Script (بجانب الكود الحالي).
 *  2) في دالة doGet الحالية أضف هذه الأسطر في أولها (بعد قراءة e.parameter.action):
 *
 *      var act = (e.parameter.action || '');
 *      if (act === 'device')       return devJson_(devCheck_(e.parameter.phone, e.parameter.device));
 *      if (act === 'resetDevices') return devJson_(devReset_(e.parameter.key, e.parameter.phone));
 *      if (act === 'login' || act === 'progress') {
 *        var g = devCheck_(e.parameter.phone, e.parameter.device);
 *        if (!g.ok) return devJson_(g);          // الحساب موقوف → لا يكمل الدخول
 *      }
 *      // ... ثم يكمل الكود الحالي كما هو
 *
 *  3) انشر النسخة من جديد: Deploy ← Manage deployments ← Edit ← New version ← Deploy (نفس الرابط).
 *
 *  رمز المعلم: عدّل الدالة isTeacherKey_ بالأسفل لتطابق طريقتك الحالية في التحقق من رمز المعلم
 *  (نفس الرمز المستخدم في «دخول المعلم» داخل لوحة تقدمي).
 */

var MAX_DEVICES = 4;
var DEV_SHEET = 'الأجهزة';   // أعمدة: الهاتف | الأجهزة | الحالة | آخر تحديث

function devJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function devSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DEV_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DEV_SHEET);
    sh.appendRow(['الهاتف', 'الأجهزة', 'الحالة', 'آخر تحديث']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function devPhone_(p) { return String(p || '').replace(/\D/g, ''); }

/** يرجع {ok:true} أو {ok:false, error:'devices'|'blocked'} */
function devCheck_(phone, device) {
  phone = devPhone_(phone);
  device = String(device || '').slice(0, 64);
  if (!phone || !device) return { ok: true };   // طلب قديم بلا معرّف جهاز: لا نمنعه
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = devSheet_();
    var vals = sh.getDataRange().getValues();
    var row = -1;
    for (var i = 1; i < vals.length; i++) {
      if (devPhone_(vals[i][0]) === phone) { row = i + 1; break; }
    }
    if (row < 0) {
      sh.appendRow(["'" + phone, JSON.stringify([device]), 'نشط', new Date()]);
      return { ok: true };
    }
    var list = [];
    try { list = JSON.parse(vals[row - 1][1] || '[]'); } catch (e) { list = []; }
    var status = String(vals[row - 1][2] || '');
    if (status === 'موقوف') return { ok: false, error: 'blocked' };
    if (list.indexOf(device) >= 0) return { ok: true };
    if (list.length >= MAX_DEVICES) {
      sh.getRange(row, 3).setValue('موقوف');
      sh.getRange(row, 4).setValue(new Date());
      sh.getRange(row, 1, 1, 4).setBackground('#fde2e2');   // يظهر لك أحمر في الشيت
      return { ok: false, error: 'devices' };
    }
    list.push(device);
    sh.getRange(row, 2).setValue(JSON.stringify(list));
    sh.getRange(row, 4).setValue(new Date());
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** إعادة تشغيل الحساب: يمسح الأجهزة المسجلة ويرجع الحالة نشط */
function devReset_(key, phone) {
  if (!isTeacherKey_(key)) return { ok: false, error: 'unauthorized' };
  phone = devPhone_(phone);
  if (!phone) return { ok: false, error: 'phone' };
  var sh = devSheet_();
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (devPhone_(vals[i][0]) === phone) {
      sh.getRange(i + 1, 2, 1, 3).setValues([['[]', 'نشط', new Date()]]);
      sh.getRange(i + 1, 1, 1, 4).setBackground(null);
      return { ok: true };
    }
  }
  return { ok: true };   // لا يوجد سجل أصلاً = الحساب غير موقوف
}

/** ⚠️ عدّلها لتطابق طريقتك الحالية في التحقق من رمز المعلم */
function isTeacherKey_(key) {
  var saved = PropertiesService.getScriptProperties().getProperty('TEACHER_KEY');
  return !!saved && String(key || '') === saved;
}
