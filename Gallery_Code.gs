/**
 * معرض «إبداعي مع النجم» — مشروع Apps Script مستقل (لا يلمس سكربت المنصة ولا شيتها بالكتابة).
 *
 * ماذا يفعل:
 *  - الطالب المشترك (رقم هاتفه مطابق لشيت «الطلاب» والحالة 2) يرفع رسمته من صفحة gallery.html.
 *  - تُحفظ الصورة في مجلد على درايف، ويُضاف صف في شيت «معرض إبداعي مع النجم» (يُنشأ تلقائيًا).
 *  - لا تظهر الرسمة لأحد إلا بعد أن تضع علامة ✔ في عمود «معتمد» (وتكتب العنوان بنفسك).
 *  - الصفحة تقرأ الرسومات المعتمدة فقط من هنا، وبعد التحقق من المشترك (لا أرقام هواتف تُرسل أبدًا).
 *
 * التشغيل: شغّل الدالة setup مرة واحدة (Run) ثم Deploy ← New deployment ← Web app
 *          (Execute as: Me ، Who has access: Anyone).
 */

var PLATFORM_SHEET_ID = '1K8BnadP-aXYVV7MqzVj6jgO6AYt4eRK7e4UBMW1YJYc';   // شيت المنصة: قراءة فقط (للتحقق من المشترك)
var GALLERY_SHEET_ID = '1p3KSH49AaR_YbiHsUHSXkftI-PV8F_Rmk3R-SbbgND4';   // شيت «معرض طلاب نجم الأحياء»: تُكتب فيه الرسومات في تبويب «المعرض»
var TEACHER_PASS = 'Teach@1212';   // كلمة سر المعلم — تُفحص هنا في الخادم فقط ولا تظهر في صفحة الطلاب
var VERSION = '2026-11-06-gallery-v10';
var TAB = 'المعرض';
var MAX_IMG_BYTES = 2500000;     // أقصى حجم للصورة بعد الضغط
var MAX_PER_DAY = 50;            // أقصى عدد مشاركات للطالب في اليوم
var MAX_PENDING = 50;            // أقصى عدد رسومات معلّقة للطالب الواحد
var HEAD = ['معتمد ✔', 'رسمة الأسبوع ⭐', 'العنوان المعتمد (اكتبه أنت)', 'المفهوم / الدرس', 'الرسمة', 'الاسم', 'عنوان من الطالب',
            'التاريخ', 'هاتف الطالب', 'رقم الملف', 'تهنئة واتساب', 'الترتيب', 'الوحدة'];
// أعمدة (من 1): 1 معتمد | 2 أسبوع | 3 عنوان معتمد | 4 مفهوم | 5 الرسمة | 6 الاسم | 7 عنوان الطالب | 8 التاريخ | 9 هاتف | 10 ملف | 11 تهنئة | 12 الترتيب (الرقم الأكبر = أعلى في المعرض) | 13 الوحدة

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function props_() { return PropertiesService.getScriptProperties(); }
function digits_(v) {
  return String(v == null ? '' : v).replace(/[\u0660-\u0669]/g, function (c) { return c.charCodeAt(0) - 0x660; })
    .replace(/[\u06F0-\u06F9]/g, function (c) { return c.charCodeAt(0) - 0x6F0; }).replace(/\D/g, '');
}
function key_(v) { return digits_(v).replace(/^974/, ''); }
function stamp_() { return Utilities.formatDate(new Date(), 'Asia/Qatar', 'yyyy-MM-dd HH:mm'); }
function safeText_(v, max) {          // نص آمن: بلا رموز تحكم/وسوم، وبلا بداية صيغة في الشيت
  var s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f<>]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max || 80);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// ================== التهيئة (مرة واحدة) ==================
function setup() {
  var p = props_(), msg = [];
  var folder;
  if (p.getProperty('FOLDER_ID')) { try { folder = DriveApp.getFolderById(p.getProperty('FOLDER_ID')); } catch (e) { folder = null; } }
  if (!folder) { folder = DriveApp.createFolder('إبداعي مع النجم — الرسومات'); p.setProperty('FOLDER_ID', folder.getId()); }
  var ss;
  if (GALLERY_SHEET_ID) { ss = SpreadsheetApp.openById(GALLERY_SHEET_ID); p.setProperty('SHEET_ID', GALLERY_SHEET_ID); }
  else {
    if (p.getProperty('SHEET_ID')) { try { ss = SpreadsheetApp.openById(p.getProperty('SHEET_ID')); } catch (e) { ss = null; } }
    if (!ss) { ss = SpreadsheetApp.create('معرض إبداعي مع النجم'); p.setProperty('SHEET_ID', ss.getId()); }
  }
  var sh = ss.getSheetByName(TAB);
  if (!sh) {
    if (GALLERY_SHEET_ID) { sh = ss.insertSheet(TAB); }      // يضيف تبويبًا جديدًا فقط، ولا يحذف أي تبويب موجود
    else {
      var first = ss.getSheets()[0];
      sh = first.getName() === TAB ? first : ss.insertSheet(TAB);
      try { if (first.getName() !== TAB && ss.getSheets().length > 1) ss.deleteSheet(first); } catch (e) {}
    }
  }
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#cfe2f3');
    sh.setFrozenRows(1);
    try { sh.setRightToLeft(true); } catch (e) {}
    sh.setColumnWidth(3, 190); sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 150); sh.setColumnWidth(6, 120); sh.setColumnWidth(7, 150);
  }
  if (!String(sh.getRange(1, 12).getValue() || '').trim()) sh.getRange(1, 12).setValue('الترتيب').setFontWeight('bold').setBackground('#cfe2f3');
  if (!String(sh.getRange(1, 13).getValue() || '').trim()) sh.getRange(1, 13).setValue('الوحدة').setFontWeight('bold').setBackground('#cfe2f3');
  // اختبار الوصول لشيت المنصة
  var plat = SpreadsheetApp.openById(PLATFORM_SHEET_ID), st = plat.getSheetByName('الطلاب');
  msg.push('✅ المجلد: ' + folder.getUrl());
  msg.push('✅ شيت الاعتماد (افتحه واعمل له علامة مرجعية): ' + ss.getUrl());
  msg.push(st ? '✅ شيت المنصة مقروء — تبويب الطلاب فيه ' + (st.getLastRow() - 1) + ' صف' : '⚠️ لم أجد تبويب «الطلاب» في شيت المنصة');
  Logger.log(msg.join('\n'));
  return msg.join('\n');
}
function sheet_() {
  var id = GALLERY_SHEET_ID || props_().getProperty('SHEET_ID');
  if (!id) throw new Error('setup_not_run');
  return SpreadsheetApp.openById(id).getSheetByName(TAB);
}

// ================== التحقق من المشترك (قراءة فقط من شيت المنصة) ==================
function readStudents_() {
  var sh = SpreadsheetApp.openById(PLATFORM_SHEET_ID).getSheetByName('الطلاب');
  if (!sh || sh.getLastRow() < 2) return {};
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return (x + '').trim(); });
  function col(names) { for (var i = 0; i < names.length; i++) { var ix = h.indexOf(names[i]); if (ix > -1) return ix; } return -1; }
  var cN = col(['اسم الطالب', 'الاسم', 'name']), cP = col(['هاتف الطالب', 'phone', 'الهاتف']), cQ = col(['هاتف ولي الأمر', 'parent']), cS = col(['الحالة', 'status']);
  var map = {};
  if (cP < 0) return map;
  for (var r = 1; r < v.length; r++) {
    var st = cS > -1 ? String(v[r][cS] == null ? '' : v[r][cS]).trim() : '';
    if (st !== '2') continue;
    var pk = key_(v[r][cP]); if (!pk) continue;
    map[pk] = { n: cN > -1 ? String(v[r][cN] || '').trim() : '' };
  }
  return map;
}
// تخزين مؤقت 5 دقائق للمشتركين المفعّلين؛ أي رقم غير موجود أو غير مطابق يُفحص مباشرة من الشيت
function subscriber_(phone) {
  var pk = key_(phone); if (!pk) return null;
  var cache = null, map = null;
  try { cache = CacheService.getScriptCache(); var raw = cache.get('gal_subs_v2'); if (raw) map = JSON.parse(raw); } catch (e) { map = null; }
  if (map && map[pk]) return { pk: pk, name: map[pk].n };
  var fresh = readStudents_();
  try { var js = JSON.stringify(fresh); if (cache && js.length < 90000) cache.put('gal_subs_v2', js, 300); } catch (e2) {}
  return fresh[pk] ? { pk: pk, name: fresh[pk].n } : null;
}

// ================== القراءة ==================
function listFor_(sub) {
  var sh = sheet_(), n = sh.getLastRow(), items = [], mine = [];
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 10).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i], approved = (r[0] === true || String(r[0]).toUpperCase() === 'TRUE'), id = String(r[9] || '');
      if (!id) continue;
      var title = String(r[2] || '').trim() || String(r[6] || '').trim();
      if (approved) items.push({ id: id, title: title || 'رسمة', concept: String(r[3] || '').trim(), name: String(r[5] || '').trim(),
        weekly: (r[1] === true || String(r[1]).toUpperCase() === 'TRUE'), ts: String(r[7] || '') });
      if (key_(r[8]) === sub.pk) mine.push({ id: id, title: title, name: String(r[5] || '').trim(), weekly: (r[1] === true || String(r[1]).toUpperCase() === 'TRUE'),
        status: approved ? 'published' : 'pending', ts: String(r[7] || '') });
    }
  }
  items.reverse(); mine.reverse();     // الأحدث أولًا
  return { ok: true, name: sub.name, items: items, mine: mine };
}

// رقم الترتيب: الأكبر = أعلى. الصفوف التي لا ترتيب لها تأخذ رقم صفها (فالأحدث أعلى تلقائيًا)
function ordOf_(r, i) {
  var o = r[11];
  return (o === '' || o == null || isNaN(Number(o))) ? i + 2 : Number(o);
}

function waLink_(phone, name, approved) {                // رابط واتساب جاهز بالتهنئة — يُرسَل للمعلم فقط بعد التحقق من كلمة السر
  var ph = digits_(phone); if (!ph) return '';
  if (ph.length === 8) ph = '974' + ph;
  var msg = approved
    ? 'تحياتي يا د/ ' + name + '! رسمتك اتنشرت في معرض «إبداعي مع النجم» ⭐ — أ. عزت نجم الأحياء'
    : 'تحياتي يا د/ ' + name + '! وصلتني رسمتك في معرض «إبداعي مع النجم» وسأراجعها قريبًا ⭐ — أ. عزت نجم الأحياء';
  return 'https://wa.me/' + ph + '?text=' + encodeURIComponent(msg);
}
function unitOf_(r) {
  var u = r[12];
  return (u === '' || u == null) ? '' : String(u);
}

// عرض عام: الرسومات المعتمدة فقط، بلا أي رقم هاتف وبلا تحقق
function listPublic_() {
  var sh = sheet_(), n = sh.getLastRow(), items = [];
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 13).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i], id = String(r[9] || '');
      if (!id || !isTrue_(r[0])) continue;
      var title = String(r[2] || '').trim() || String(r[6] || '').trim();
      items.push({ id: id, title: title || 'رسمة', concept: String(r[3] || '').trim(), name: String(r[5] || '').trim(),
        weekly: isTrue_(r[1]), ts: String(r[7] || ''), ord: ordOf_(r, i), unit: unitOf_(r) });
    }
  }
  items.sort(function (a, b) { return b.ord - a.ord; });
  return { ok: true, items: items };
}

// ================== أدوات المعلم (بكلمة سر، من الخادم فقط) ==================
function isTrue_(x) { return x === true || String(x).toUpperCase() === 'TRUE'; }
function adminOk_(pw) {                       // حماية من التخمين: 10 محاولات خاطئة ثم إيقاف 10 دقائق
  var c = CacheService.getScriptCache(), f = Number(c.get('adm_fail') || 0);
  if (f >= 10) return 'locked';
  if (String(pw || '') === TEACHER_PASS) return true;
  c.put('adm_fail', String(f + 1), 600);
  return 'bad_password';
}
function adminList_() {                       // كل الرسومات (المعتمدة والمعلّقة)، بلا أي أرقام هواتف
  var sh = sheet_(), n = sh.getLastRow(), items = [];
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 13).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i], id = String(r[9] || ''); if (!id) continue;
      items.push({ id: id, approved: isTrue_(r[0]), weekly: isTrue_(r[1]), title: String(r[2] || '').trim(), studentTitle: String(r[6] || '').trim(),
        concept: String(r[3] || '').trim(), name: String(r[5] || '').trim(), ts: String(r[7] || ''), ord: ordOf_(r, i), unit: unitOf_(r),
        wa: waLink_(r[8], String(r[5] || '').trim(), isTrue_(r[0])) });
    }
  }
  items.sort(function (a, b) { return b.ord - a.ord; });
  return { ok: true, items: items };
}
function findRow_(sh, id) {
  var n = sh.getLastRow(); if (n < 2) return 0;
  var ids = sh.getRange(2, 10, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function adminUpdate_(d) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), row = findRow_(sh, d.id); if (!row) return { ok: false, error: 'not_found' };
    if (d.name !== undefined) { var nm = safeText_(d.name, 40); if (nm.length < 2) return { ok: false, error: 'bad_name' }; sh.getRange(row, 6).setValue(nm); }
    if (d.title !== undefined) sh.getRange(row, 3).setValue(safeText_(d.title, 80));
    if (d.concept !== undefined) sh.getRange(row, 4).setValue(safeText_(d.concept, 60));
    if (d.unit !== undefined) {
      var un = String(d.unit).trim();
      if (un !== '' && (!/^\d{1,2}$/.test(un) || Number(un) < 1)) return { ok: false, error: 'bad_unit' };
      sh.getRange(row, 13).setValue(un === '' ? '' : Number(un));
    }
    if (d.approved !== undefined) sh.getRange(row, 1).setValue(!!d.approved);
    if (d.weekly !== undefined) sh.getRange(row, 2).setValue(!!d.weekly);
    var wa = waLink_(sh.getRange(row, 9).getValue(), String(sh.getRange(row, 6).getValue() || ''), isTrue_(sh.getRange(row, 1).getValue()));
    return { ok: true, wa: wa };
  } finally { lock.releaseLock(); }
}
function adminReorder_(d) {                   // ids بالترتيب المطلوب من الأعلى إلى الأسفل
  var ids = d.ids;
  if (!Array.isArray(ids) || !ids.length || ids.length > 2000) return { ok: false, error: 'bad_request' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow(); if (n < 2) return { ok: true };
    if (!String(sh.getRange(1, 12).getValue() || '').trim()) sh.getRange(1, 12).setValue('الترتيب').setFontWeight('bold').setBackground('#cfe2f3');
    var idv = sh.getRange(2, 10, n - 1, 1).getValues(), ov = sh.getRange(2, 12, n - 1, 1).getValues(), pos = {}, N = ids.length;
    for (var i = 0; i < N; i++) pos[String(ids[i])] = N - i;
    for (var r = 0; r < idv.length; r++) { var k = String(idv[r][0]); if (pos[k] !== undefined) ov[r][0] = pos[k]; }
    sh.getRange(2, 12, n - 1, 1).setValues(ov);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function adminDelete_(d) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), row = findRow_(sh, d.id); if (!row) return { ok: false, error: 'not_found' };
    sh.deleteRow(row);
    try { DriveApp.getFileById(String(d.id)).setTrashed(true); } catch (e) {}
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ================== الإرسال ==================
function isImage_(bytes) {
  if (!bytes || bytes.length < 12) return null;
  var b = bytes.map(function (x) { return x & 255; });
  if (b[0] === 0xFF && b[1] === 0xD8) return 'image/jpeg';
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'image/png';
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'image/webp';
  return null;
}
function submit_(d) {
  var sub = subscriber_(d.phone);
  if (!sub) return { ok: false, error: 'not_subscriber' };
  var name = safeText_(d.name, 40), title = safeText_(d.title, 80);
  if (name.length < 2) return { ok: false, error: 'bad_name' };
  var unit = String(d.unit == null ? '' : d.unit).trim();
  if (!/^\d{1,2}$/.test(unit) || Number(unit) < 1) return { ok: false, error: 'bad_unit' };
  var b64 = String(d.image || '').replace(/^data:[^;]+;base64,/, '');
  if (!b64) return { ok: false, error: 'bad_image' };
  var bytes; try { bytes = Utilities.base64Decode(b64); } catch (e) { return { ok: false, error: 'bad_image' }; }
  if (bytes.length > MAX_IMG_BYTES) return { ok: false, error: 'too_big' };
  var mime = isImage_(bytes); if (!mime) return { ok: false, error: 'bad_image' };

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow();
    if (n >= 2) {                       // حدود الاستخدام
      var v = sh.getRange(2, 1, n - 1, 9).getValues(), today = stamp_().slice(0, 10), perDay = 0, pending = 0;
      for (var i = 0; i < v.length; i++) {
        if (key_(v[i][8]) !== sub.pk) continue;
        if (String(v[i][7]).slice(0, 10) === today) perDay++;
        if (!(v[i][0] === true || String(v[i][0]).toUpperCase() === 'TRUE')) pending++;
      }
      if (perDay >= MAX_PER_DAY) return { ok: false, error: 'limit_day' };
      if (pending >= MAX_PENDING) return { ok: false, error: 'limit_pending' };
    }
    var ext = mime === 'image/png' ? 'png' : (mime === 'image/webp' ? 'webp' : 'jpg');
    var folder = DriveApp.getFolderById(props_().getProperty('FOLDER_ID'));
    var file = folder.createFile(Utilities.newBlob(bytes, mime, 'drawing_' + Date.now() + '_' + sub.pk + '.' + ext));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    var fid = file.getId(), row = sh.getLastRow() + 1;
    var phone = (sub.pk.length === 8) ? ('974' + sub.pk) : sub.pk;
    sh.getRange(row, 8, 1, 3).setNumberFormat('@');
    sh.getRange(row, 1, 1, 10).setValues([[false, false, '', '', '', name, title, stamp_(), phone, fid]]);
    sh.getRange(row, 1, 1, 2).insertCheckboxes();
    sh.getRange(row, 5).setFormula('=IMAGE("https://drive.google.com/thumbnail?id="&J' + row + '&"&sz=w300")');
    sh.getRange(row, 11).setFormula('=IF(A' + row + '=TRUE,HYPERLINK("https://wa.me/"&I' + row + '&"?text="&ENCODEURL("تحياتي يا د/ "&F' + row + '&"! رسمتك اتنشرت في معرض «إبداعي مع النجم» ⭐ — أ. عزت نجم الأحياء"),"📲 تهنئة"),"")');
    sh.setRowHeight(row, 100);
    var ordMax = 0;
    if (n >= 2) { var ov = sh.getRange(2, 12, n - 1, 1).getValues(); for (var q = 0; q < ov.length; q++) { var o = Number(ov[q][0]); if (o > ordMax) ordMax = o; } }
    sh.getRange(row, 12).setValue(Math.max(ordMax, row) + 1);     // الأحدث في الأعلى
    sh.getRange(row, 13).setValue(Number(unit));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ================== MY BOARD (رسومات المعلم، منفصلة تمامًا عن رسومات الطلاب) ==================
var BOARD_TAB = 'MY BOARD';
var BOARD_HEAD = ['العنوان', 'الرسمة', 'التاريخ', 'رقم الملف', 'الترتيب'];
function boardSheet_() {
  var ss = SpreadsheetApp.openById(GALLERY_SHEET_ID || props_().getProperty('SHEET_ID'));
  var sh = ss.getSheetByName(BOARD_TAB);
  if (!sh) {
    sh = ss.insertSheet(BOARD_TAB);
    sh.getRange(1, 1, 1, BOARD_HEAD.length).setValues([BOARD_HEAD]).setFontWeight('bold').setBackground('#cfe2f3');
    sh.setFrozenRows(1);
    try { sh.setRightToLeft(true); } catch (e) {}
  }
  return sh;
}
function listBoard_() {
  var sh = boardSheet_(), n = sh.getLastRow(), items = [];
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 5).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i], id = String(r[3] || ''); if (!id) continue;
      var o = Number(r[4]);
      items.push({ id: id, title: String(r[0] || '').trim() || 'رسمة', ts: String(r[2] || ''), ord: (isNaN(o) || r[4] === '') ? i + 2 : o });
    }
  }
  items.sort(function (a, b) { return b.ord - a.ord; });
  return { ok: true, items: items };
}
function boardAdd_(d) {
  var title = safeText_(d.title, 80) || 'رسمة';
  var b64 = String(d.image || '').replace(/^data:[^;]+;base64,/, '');
  if (!b64) return { ok: false, error: 'bad_image' };
  var bytes; try { bytes = Utilities.base64Decode(b64); } catch (e) { return { ok: false, error: 'bad_image' }; }
  if (bytes.length > MAX_IMG_BYTES) return { ok: false, error: 'too_big' };
  var mime = isImage_(bytes); if (!mime) return { ok: false, error: 'bad_image' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = boardSheet_(), n = sh.getLastRow();
    var ext = mime === 'image/png' ? 'png' : (mime === 'image/webp' ? 'webp' : 'jpg');
    var folder = DriveApp.getFolderById(props_().getProperty('FOLDER_ID'));
    var file = folder.createFile(Utilities.newBlob(bytes, mime, 'board_' + Date.now() + '.' + ext));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    var fid = file.getId(), row = Math.max(n, 1) + 1, ordMax = 0;
    if (n >= 2) { var ov = sh.getRange(2, 5, n - 1, 1).getValues(); for (var q = 0; q < ov.length; q++) { var o = Number(ov[q][0]); if (o > ordMax) ordMax = o; } }
    sh.getRange(row, 3, 1, 2).setNumberFormat('@');
    sh.getRange(row, 1, 1, 5).setValues([[title, '', stamp_(), fid, Math.max(ordMax, row) + 1]]);
    sh.getRange(row, 2).setFormula('=IMAGE("https://drive.google.com/thumbnail?id="&D' + row + '&"&sz=w300")');
    sh.setRowHeight(row, 100);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function boardDelete_(d) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = boardSheet_(), n = sh.getLastRow(); if (n < 2) return { ok: false, error: 'not_found' };
    var ids = sh.getRange(2, 4, n - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(d.id)) {
        sh.deleteRow(i + 2);
        try { DriveApp.getFileById(String(d.id)).setTrashed(true); } catch (e) {}
        return { ok: true };
      }
    }
    return { ok: false, error: 'not_found' };
  } finally { lock.releaseLock(); }
}

// ================== المسارات ==================
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.action === 'gallery_ping') return json_({ ok: true, version: VERSION });
    if (p.action === 'gallery_public') return json_(listPublic_());
    if (p.action === 'board_public') return json_(listBoard_());
    if (p.action === 'gallery_list') {
      var sub = subscriber_(p.phone);
      if (!sub) return json_({ ok: false, error: 'not_subscriber' });
      return json_(listFor_(sub));
    }
    return json_({ ok: true, service: 'najm-gallery', version: VERSION });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.action === 'gallery_submit') return json_(submit_(d));
    if (String(d.action || '').indexOf('admin_') === 0) {
      var ok = adminOk_(d.pass);
      if (ok !== true) return json_({ ok: false, error: ok === 'locked' ? 'locked' : 'bad_password' });
      if (d.action === 'admin_list') return json_(adminList_());
      if (d.action === 'admin_update') return json_(adminUpdate_(d));
      if (d.action === 'admin_delete') return json_(adminDelete_(d));
      if (d.action === 'admin_reorder') return json_(adminReorder_(d));
      if (d.action === 'admin_board_add') return json_(boardAdd_(d));
      if (d.action === 'admin_board_delete') return json_(boardDelete_(d));
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}
