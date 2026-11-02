/**
 * معرض «إبداعي مع النجم» — مشروع Apps Script مستقل (لا يلمس سكربت المنصة ولا شيتها بالكتابة).
 *
 * ماذا يفعل:
 *  - الطالب المشترك (رقمه + رقم ولي أمره مطابقان لشيت «الطلاب» والحالة 2) يرفع رسمته من صفحة gallery.html.
 *  - تُحفظ الصورة في مجلد على درايف، ويُضاف صف في شيت «معرض إبداعي مع النجم» (يُنشأ تلقائيًا).
 *  - لا تظهر الرسمة لأحد إلا بعد أن تضع علامة ✔ في عمود «معتمد» (وتكتب العنوان بنفسك).
 *  - الصفحة تقرأ الرسومات المعتمدة فقط من هنا، وبعد التحقق من المشترك (لا أرقام هواتف تُرسل أبدًا).
 *
 * التشغيل: شغّل الدالة setup مرة واحدة (Run) ثم Deploy ← New deployment ← Web app
 *          (Execute as: Me ، Who has access: Anyone).
 */

var PLATFORM_SHEET_ID = '1K8BnadP-aXYVV7MqzVj6jgO6AYt4eRK7e4UBMW1YJYc';   // شيت المنصة: قراءة فقط (للتحقق من المشترك)
var VERSION = '2026-11-01-gallery';
var TAB = 'المعرض';
var MAX_IMG_BYTES = 2500000;     // أقصى حجم للصورة بعد الضغط
var MAX_PER_DAY = 5;             // أقصى عدد مشاركات للطالب في اليوم
var MAX_PENDING = 8;             // أقصى عدد رسومات معلّقة للطالب الواحد
var HEAD = ['معتمد ✔', 'رسمة الأسبوع ⭐', 'العنوان المعتمد (اكتبه أنت)', 'المفهوم / الدرس', 'الرسمة', 'الاسم', 'عنوان من الطالب',
            'التاريخ', 'هاتف الطالب', 'رقم الملف', 'تهنئة واتساب'];
// أعمدة (من 1): 1 معتمد | 2 أسبوع | 3 عنوان معتمد | 4 مفهوم | 5 الرسمة | 6 الاسم | 7 عنوان الطالب | 8 التاريخ | 9 هاتف | 10 ملف | 11 تهنئة

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
  if (p.getProperty('SHEET_ID')) { try { ss = SpreadsheetApp.openById(p.getProperty('SHEET_ID')); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create('معرض إبداعي مع النجم'); p.setProperty('SHEET_ID', ss.getId()); }
  var sh = ss.getSheetByName(TAB);
  if (!sh) {
    var first = ss.getSheets()[0];
    sh = first.getName() === TAB ? first : ss.insertSheet(TAB);
    try { if (first.getName() !== TAB && ss.getSheets().length > 1) ss.deleteSheet(first); } catch (e) {}
  }
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#cfe2f3');
    sh.setFrozenRows(1);
    try { sh.setRightToLeft(true); } catch (e) {}
    sh.setColumnWidth(3, 190); sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 150); sh.setColumnWidth(6, 120); sh.setColumnWidth(7, 150);
  }
  // اختبار الوصول لشيت المنصة
  var plat = SpreadsheetApp.openById(PLATFORM_SHEET_ID), st = plat.getSheetByName('الطلاب');
  msg.push('✅ المجلد: ' + folder.getUrl());
  msg.push('✅ شيت الاعتماد (افتحه واعمل له علامة مرجعية): ' + ss.getUrl());
  msg.push(st ? '✅ شيت المنصة مقروء — تبويب الطلاب فيه ' + (st.getLastRow() - 1) + ' صف' : '⚠️ لم أجد تبويب «الطلاب» في شيت المنصة');
  Logger.log(msg.join('\n'));
  return msg.join('\n');
}
function sheet_() {
  var id = props_().getProperty('SHEET_ID');
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
  if (cP < 0 || cQ < 0) return map;
  for (var r = 1; r < v.length; r++) {
    var st = cS > -1 ? String(v[r][cS] == null ? '' : v[r][cS]).trim() : '';
    if (st !== '2') continue;
    var pk = key_(v[r][cP]); if (!pk) continue;
    map[pk] = { p: key_(v[r][cQ]), n: cN > -1 ? String(v[r][cN] || '').trim() : '' };
  }
  return map;
}
// تخزين مؤقت 5 دقائق للمشتركين المفعّلين؛ أي رقم غير موجود أو غير مطابق يُفحص مباشرة من الشيت
function subscriber_(phone, parent) {
  var pk = key_(phone), qk = key_(parent); if (!pk || !qk) return null;
  var cache = null, map = null;
  try { cache = CacheService.getScriptCache(); var raw = cache.get('gal_subs_v1'); if (raw) map = JSON.parse(raw); } catch (e) { map = null; }
  if (map && map[pk] && map[pk].p === qk) return { pk: pk, name: map[pk].n };
  var fresh = readStudents_();
  try { var js = JSON.stringify(fresh); if (cache && js.length < 90000) cache.put('gal_subs_v1', js, 300); } catch (e2) {}
  return (fresh[pk] && fresh[pk].p === qk) ? { pk: pk, name: fresh[pk].n } : null;
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
  var sub = subscriber_(d.phone, d.parent);
  if (!sub) return { ok: false, error: 'not_subscriber' };
  var name = safeText_(d.name, 40), title = safeText_(d.title, 80);
  if (name.length < 2) return { ok: false, error: 'bad_name' };
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
    sh.getRange(row, 11).setFormula('=IF(A' + row + '=TRUE,HYPERLINK("https://wa.me/"&I' + row + '&"?text="&ENCODEURL("مبروك يا "&F' + row + '&"! رسمتك اتنشرت في معرض «إبداعي مع النجم» ⭐ — أ. عزت نجم الأحياء"),"📲 تهنئة"),"")');
    sh.setRowHeight(row, 100);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ================== المسارات ==================
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.action === 'gallery_ping') return json_({ ok: true, version: VERSION });
    if (p.action === 'gallery_list') {
      var sub = subscriber_(p.phone, p.parent);
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
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}
