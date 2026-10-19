/**
 * كود Google Apps Script — النسخة الكاملة (الكود القديم كما هو + إضافات لوحة التقدم).
 * انسخه بالكامل مكان الكود الحالي، ثم: Deploy > Manage deployments > ✏️ Edit
 * > Version: New version > Deploy  (هكذا يبقى الرابط نفسه دون تغيير).
 *
 * الشيتات: "الطلاب" | "المحتوى" | "المتفوقون" | "لمحة عني"  (كما كانت)
 * جديد: شيت "النتائج" يُنشأ تلقائيًا عند أول نتيجة، لا تعدّل فيه يدويًا.
 * جديد: التسجيل الذاتي — الطالب يملأ نموذجًا فيُضاف صف في شيت "الطلاب" ملوّن بالأصفر.
 *   عمود "الحالة": 2 = مفعّل (يدخل المحتوى المشترك واللوحة). 0 أو 1 أو فراغ = تصفح عام فقط.
 *   عند كتابتك 2 يزول اللون الأصفر تلقائيًا.
 * أمان: لم يعد `?sheet=users` يعرض أي بيانات — التحقق يتم داخل السكربت ويرجع الاسم فقط.
 *
 * أعمدة المحتوى/المتفوقون/لمحة عني: كما في الكود القديم تمامًا (لم تتغير).
 *
 * تحديث 2026-10-17 (أ):
 *  - طلب التسجيل لم يعد يُتجاهل بصمت: كل طلب (جديد/مكرر/مرفوض) يُسجَّل في شيت "سجل الطلبات".
 *  - الرقم المكرر لا يُضاف مرة ثانية، ويُكتب في عمود "ملاحظات" للصف الموجود، والطالب يرى رسالة واضحة.
 *  - شيت "الطلاب": منع تكرار هاتف الطالب (يُلغى الإدخال اليدوي المكرر مع تنبيه، واللصق يُلوَّن بالأحمر).
 *  - قائمة 📊 نجم ← فحص الأرقام المكررة.
 *  - للتأكد أن النشر محدَّث: افتح رابط السكربت + ?action=ping فيظهر رقم الإصدار.
 *
 * تحديث 2026-10-18 (أ): لوحة تقدمي ترجع الآن totalExams = عدد الاختبارات الظاهرة للمشتركين
 *  (صفوف شيت "المحتوى" ظهورها 1 أو 2 ولها رابط فعلي، ورابطها يحتوي MCQ؛ كل نسخة أ/ب اختبار مستقل).
 *  للتحقق: افتح رابط السكربت + ?action=exams ليعرض القائمة التي حُسبت.
 *
 * تحديث 2026-10-18 (ب): رمز الاختبار القديم والجديد يُحسبان اختبارًا واحدًا (U2L1_P1_P2_* = U2L1_P2_*،
 *  U2L1_P1_P3_* = U2L1_P3_*) في لوحة تقدمي وفي إحصاء الأسئلة، دون تعديل أي صف في شيت "النتائج".
 *
 * تحديث 2026-10-19 (أ): الأرقام المكررة في شيت "الطلاب" لم تعد تُحذف تلقائيًا — تُلوَّن بالأحمر كل الصفوف المتشابهة،
 *  ويظهر تبويب "الأرقام المكررة" فيه كل مجموعة جنبًا إلى جنب (الاسم، الحالة، عدد النتائج، تاريخ التسجيل...) لتقرر بنفسك
 *  أي صف تحتفظ به. ويظهر تنبيه عند فتح الملف إن وُجد تكرار.
 *
 * تحديث 2026-10-19 (ب): تبويب "أرقام متكررة": كل رقم وتحته مباشرة صفوفه المتشابهة، وأمام كل صف خانة ✔.
 *  تضع ✔ أمام الصف الذي تريد بقاءه ثم من القائمة 📊 نجم ← "حذف المكرر (احتفظ بالمُحدّد ✔)": يُحذف الباقي من شيت "الطلاب"
 *  بعد رسالة تأكيد، ويُحفظ نسخة من كل صف محذوف في تبويب "المحذوفات" (للرجوع إليه).
 *
 * تحديث 2026-10-19 (ج): صفوف كل رقم في تبويب "أرقام متكررة" مُجمَّعة (تجميع الصفوف الأصلي في جوجل شيت): أزرار [−]/[+]
 *  بجانب عنوان كل رقم لطيّ صفوفه وفتحها، وأزرار 1/2 أعلى اليمين لطيّ كل الأرقام أو فتحها دفعة واحدة.
 *
 * تحديث 2026-10-19 (د): عمود "المجموعة" (العمود C في شيت الطلاب) يظهر في تبويب "أرقام متكررة" بجانب الاسم،
 *  وإن كان الرقم المكرر في مجموعات مختلفة يُكتب ذلك في عنوانه، ويظهر اسم المجموعة في رسالة تأكيد الحذف.
 *
 * تحديث 2026-10-19 (هـ): تبويب "أرقام متكررة" يعرض **كل أعمدة** صف الطالب كما هي في شيت الطلاب (أيًّا كانت)،
 *  مع عمود "اكتمال" (كم خانة معبأة)، والخانات الفارغة بلون وردي، و⭐ تقترح الصف الأكمل معلومات (ثم المفعّل ثم الأقدم).
 *
 * تحديث 2026-10-19 (و): تذكير الأسبوع للمشتركين (الحالة 2):
 *  - الموقع يسجّل متى فتح الطالب كل ملف (شيت "فتح المحتوى")، ويعرض له شريط تذكير داخل التطبيق.
 *  - السكربت يسجّل أول ظهور لكل ملف جديد (شيت "سجل المحتوى"؛ الملفات الموجودة وقت التفعيل تُعتبر قديمة ولا يُذكَّر بها).
 *  - شيت "تذكيرات الأسبوع": كل طالب مرّ أسبوع على ملف لم يفتحه، مع نص الرسالة وزر واتساب جاهز؛ ✔ أمام الاسم = أُرسل.
 *  - للتفعيل: القائمة 📊 نجم ← تفعيل تذكير الأسبوع (يوميًا).
 */

var RESULTS_SHEET = 'النتائج';
var LOG_SHEET = 'سجل الطلبات';
var VERSION = '2026-10-19-f';
var DUP_COLOR = '#ea9999';
var DUP_SHEET = 'أرقام متكررة';
var DUP_SHEET_OLD = 'الأرقام المكررة';
var DEL_SHEET = 'المحذوفات';
var CONTENT_LOG = 'سجل المحتوى', OPENS_SHEET = 'فتح المحتوى', REMINDED_SHEET = 'تذكيرات مرسلة', REMINDERS_SHEET = 'تذكيرات الأسبوع';
var REMIND_DAYS = 7, REMIND_WINDOW_DAYS = 45;

function doGet(e) {
  var p = e.parameter || {};
  if (p.action === 'ping') {
    return json_({ ok: true, version: VERSION, time: new Date().toISOString() });
  }
  if (p.action === 'exams') {
    var ex = examItems_();
    var oldN = ex.items.filter(function (i) { return i.oldName; }).length;
    return json_({ ok: true, rule: ex.rule, total: ex.items.length, oldNameLinks: oldN, note: oldN ? 'تنبيه: يوجد روابط في شيت المحتوى ما زالت بالاسم القديم (oldName=true) — غيّرها إلى الاسم الجديد' : 'كل الروابط بالأسماء الصحيحة', items: ex.items });
  }
  if (p.action === 'pending') {
    return json_(pending_(p.phone, p.parent));
  }
  if (p.action === 'progress') {
    return json_(progress_(p.phone, p.parent));
  }
  if (p.action === 'login') {
    return json_(login_(p.phone, p.parent));
  }
  if (((p.sheet || 'users') + '').toLowerCase() === 'users') {
    return json_([]); // أُغلق عرض قائمة الطلاب للعموم
  }

  // ===== الكود القديم بدون تغيير =====
  var sheetParam = ((e.parameter.sheet || 'users') + '').toLowerCase();
  var sheetNameMap = { users: 'الطلاب', content: 'المحتوى', students: 'المتفوقون', about: 'لمحة عني' };
  var sheetName = sheetNameMap[sheetParam] || 'الطلاب';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = (headers[j] + '').trim();
      obj[key] = values[i][j];
    }
    rows.push(obj);
  }

  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.action === 'saveResult') return json_(saveResult_(d));
    if (d.action === 'register') return json_(register_(d));
    if (d.action === 'open') return json_(open_(d));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    logReq_('خطأ', {}, String(err));
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- أدوات ----------
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function digits_(v) {
  return String(v == null ? '' : v)
    .replace(/[\u0660-\u0669]/g, function (c) { return c.charCodeAt(0) - 0x660; })
    .replace(/[\u06F0-\u06F9]/g, function (c) { return c.charCodeAt(0) - 0x6F0; })
    .replace(/\D/g, '');
}
// مفتاح المطابقة: نفس منطق الدخول القديم (حذف 974 من البداية)
function key_(v) { return digits_(v).replace(/^974/, ''); }
// يوحّد رمز الاختبار القديم والجديد: U2L1_P1_P2_MCQ_B و U2L1_P2_MCQ_B = اختبار واحد
function normExam_(id) { return String(id == null ? '' : id).replace(/^(U\d+L\d+)_P1_(P[23])_/i, '$1_$2_'); }

function stamp_() { return Utilities.formatDate(new Date(), 'Asia/Qatar', 'yyyy-MM-dd HH:mm'); }

// سجل لكل طلب تسجيل/دخول حتى لا يضيع شيء بصمت (أي خطأ هنا لا يوقف العملية الأصلية)
function logReq_(kind, d, outcome) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(LOG_SHEET);
      sh.appendRow(['التاريخ', 'النوع', 'الاسم', 'هاتف الطالب', 'هاتف ولي الأمر', 'النتيجة']);
      sh.getRange('C:E').setNumberFormat('@');
      sh.setFrozenRows(1);
    }
    if (sh.getLastRow() > 2500) sh.deleteRows(2, 500);
    d = d || {};
    sh.appendRow([stamp_(), kind, String(d.name || '').slice(0, 80), String(d.phone || '').slice(0, 20), String(d.parent || '').slice(0, 20), String(outcome || '').slice(0, 300)]);
  } catch (e) {}
}

function students_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('الطلاب');
  if (!sh) return [];
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return (x + '').trim(); }), out = [];
  for (var i = 1; i < v.length; i++) {
    var r = {};
    for (var j = 0; j < h.length; j++) r[h[j]] = v[i][j];
    out.push({
      phone: key_(r.phone || r['هاتف الطالب'] || r['الهاتف'] || ''),
      parent: key_(r.parent || r['هاتف ولي الأمر'] || ''),
      name: String(r.name || r['اسم الطالب'] || r['الاسم'] || r['الطالب'] || '').trim(),
      active: String(r['الحالة'] != null ? r['الحالة'] : (r.status != null ? r.status : '')).trim() === '2'
    });
  }
  return out;
}

function resultsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RESULTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RESULTS_SHEET);
    sh.appendRow(['id', 'التاريخ', 'هاتف الطالب', 'الاسم', 'معرف الاختبار', 'الاختبار', 'الدرجة', 'الدرجة الكلية', 'النسبة', 'الوقت', 'الأخطاء']);
    sh.getRange('A:C').setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  if (!sh.getRange(1, 11).getValue()) sh.getRange(1, 11).setValue('الأخطاء'); // للشيت القديم
  return sh;
}

function parseWrong_(c) {
  try { var a = JSON.parse(c); return Array.isArray(a) ? a : null; } catch (e) { return null; }
}

// ---------- حفظ نتيجة (من ملفات الامتحانات) ----------
function saveResult_(d) {
  var phone = digits_(d.phone), score = Number(d.score), total = Number(d.total);
  if (!d.id || !d.exam || !phone || !isFinite(score) || !isFinite(total) || total <= 0) return { ok: false, error: 'bad data' };
  var st = students_().filter(function (s) { return s.phone === key_(phone); })[0];
  if (!st || !st.active) return { ok: true }; // لا نخزّن إلا المفعّلين
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = resultsSheet_(), n = sh.getLastRow();
    if (n > 1) {
      var ids = sh.getRange(2, 1, n - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (ids[i][0] === String(d.id)) return { ok: true };
    }
    var pct = Math.round(score / total * 10000) / 100;
    var wrong = [];
    (Array.isArray(d.wrong) ? d.wrong : []).slice(0, 60).forEach(function (w) {
      wrong.push({ q: String((w && w.q) || '').slice(0, 200), a: String((w && w.a) || '').slice(0, 100), c: String((w && w.c) || '').slice(0, 100) });
    });
    var wj = JSON.stringify(wrong);
    while (wj.length > 45000 && wrong.length) { wrong.pop(); wj = JSON.stringify(wrong); }
    if (!Array.isArray(d.wrong)) wj = ''; // نسخة قديمة بلا تفاصيل: تبقى "غير معروف" ولا تُعرض كأنها بلا أخطاء
    sh.appendRow([String(d.id), String(d.at || new Date().toISOString()), phone, st.name || String(d.name || ''),
      String(d.exam), String(d.title || d.exam), score, total, pct, String(d.time || ''), wj]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ---------- الاختبارات الظاهرة للمشتركين (لسطر "أنجز X من أصل Y" في لوحة تقدمي) ----------
function examItems_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('المحتوى');
  if (!sh || sh.getLastRow() < 2) return { rule: 'none', items: [] };
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return (x + '').trim(); });
  function col(names) { for (var i = 0; i < names.length; i++) { var ix = h.indexOf(names[i]); if (ix > -1) return ix; } return -1; }
  var cLevel = col(['ظهور', 'level']), cLink = col(['الرابط', 'link']), cTitle = col(['العنوان', 'title']), cUnit = col(['الوحدة', 'unit']);
  if (cLink < 0 || cTitle < 0) return { rule: 'none', items: [] };
  var rows = [];
  for (var i = 1; i < v.length; i++) {
    var raw = cLevel > -1 ? String(v[i][cLevel] == null ? '' : v[i][cLevel]).trim() : '';
    if (raw === '00') continue;                                   // مخفي تمامًا
    var lv = parseInt(raw, 10); if (isNaN(lv)) lv = 1;
    if (lv !== 1 && lv !== 2) continue;                           // 0 = قريبًا، 3 = مميز: لا يُحسبان
    var title = String(v[i][cTitle] == null ? '' : v[i][cTitle]).trim();
    var link = String(v[i][cLink] == null ? '' : v[i][cLink]).trim();
    if (!title || !link || link === '#') continue;                // بلا رابط فعلي = "قريبًا"
    if (cUnit > -1 && !/\d/.test(String(v[i][cUnit] == null ? '' : v[i][cUnit]))) continue; // الموقع يتجاهل الصف بلا رقم وحدة
    rows.push({ title: title, link: link, level: lv });
  }
  var pick = rows.filter(function (r) { return /mcq/i.test(r.link); }), rule = 'link';
  if (!pick.length) { rule = 'title'; pick = rows.filter(function (r) { return /بنك\s*[أا]سئلة/.test(r.title); }); }
  var seen = {}, out = [];
  pick.forEach(function (r) {
    // رابط بالاسم القديم (P1_P2 / P1_P3) ورابط بالاسم الجديد لنفس الاختبار يُحسبان مرة واحدة
    var k = r.link.toLowerCase().split('#')[0].split('?')[0].replace(/_p1_(p[23])_/, '_$1_');
    r.oldName = /_P1_P[23]_/i.test(r.link);
    if (!seen[k]) { seen[k] = 1; out.push(r); }
  });
  return { rule: rule, items: out };
}

// ---------- لوحة التقدم: تتحقق من الرقمين ثم ترجع بيانات هذا الطالب فقط ----------
function progress_(phone, parent) {
  var pk = key_(phone), qk = key_(parent);
  if (!pk || !qk) return { ok: false, error: 'bad input' };
  var sts = students_();
  var me = sts.filter(function (s) { return s.phone === pk && s.parent === qk; })[0];
  if (!me) return { ok: false, error: 'not found' };
  if (!me.active) return { ok: false, error: 'pending' };

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET);
  var best = {}; // phoneKey -> exam -> أفضل محاولة
  var attempts = {}; // exam -> عدد محاولات هذا الطالب
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
    for (var i = 0; i < v.length; i++) {
      var k = key_(v[i][2]), ex = normExam_(v[i][4]), pct = Number(v[i][8]);
      if (!k || !ex) continue;
      best[k] = best[k] || {};
      if (k === pk) attempts[ex] = (attempts[ex] || 0) + 1;
      var cur = best[k][ex];
      if (!cur || pct > cur.percent) best[k][ex] = { exam: ex, title: v[i][5], score: Number(v[i][6]), total: Number(v[i][7]), percent: pct, at: v[i][1], wrong: parseWrong_(v[i][10]) };
    }
  }
  function overall(k) {
    var s = 0, t = 0, c = 0, b = best[k] || {};
    for (var ex in b) { s += b[ex].score; t += b[ex].total; c++; }
    return { pct: t ? s / t * 100 : null, taken: c };
  }
  var mine = overall(pk), exams = [];
  for (var ex in (best[pk] || {})) { var o = best[pk][ex]; o.attempts = attempts[ex] || 1; exams.push(o); }
  exams.sort(function (a, b) { return a.exam < b.exam ? -1 : 1; });

  var topPct = null, n = 0;
  if (mine.pct !== null) {
    var higher = 0;
    var activeKeys = {}; sts.forEach(function (s) { if (s.active) activeKeys[s.phone] = 1; });
    for (var k2 in best) { if (!activeKeys[k2]) continue; var o2 = overall(k2); if (o2.pct === null) continue; n++; if (k2 !== pk && o2.pct > mine.pct + 1e-9) higher++; }
    topPct = Math.ceil((higher + 1) / n * 100);
  }
  var totalExams = 0;
  try { totalExams = examItems_().items.length; } catch (e) { totalExams = 0; }
  return { ok: true, name: me.name, exams: exams, overall: mine.pct === null ? null : Math.round(mine.pct * 100) / 100, taken: mine.taken, totalExams: totalExams, topPct: topPct, n: n };
}

// ---------- تسجيل الدخول: يتحقق من الرقمين والحالة ويرجع الاسم فقط ----------
function login_(phone, parent) {
  var pk = key_(phone), qk = key_(parent);
  var ld = { phone: digits_(phone), parent: digits_(parent) };
  if (!pk || !qk) { logReq_('دخول', ld, 'رُفض: رقم غير صحيح'); return { ok: false, error: 'not_found' }; }
  var me = students_().filter(function (s) { return s.phone === pk && s.parent === qk; })[0];
  if (!me) { logReq_('دخول', ld, 'غير مسجل'); return { ok: false, error: 'not_found' }; }
  if (!me.active) { logReq_('دخول', ld, 'قيد المراجعة (الحالة ليست 2)'); return { ok: false, error: 'pending' }; }
  return { ok: true, name: me.name };
}

// ---------- التسجيل الذاتي: صف جديد أصفر في شيت "الطلاب" بدون حالة ----------
function ensureCol_(sh, names, create) {
  var lc = Math.max(sh.getLastColumn(), 1);
  var h = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return (x + '').trim(); });
  for (var i = 0; i < names.length; i++) { var ix = h.indexOf(names[i]); if (ix > -1) return ix + 1; }
  if (!create) return 0;
  sh.getRange(1, lc + 1).setValue(create);
  return lc + 1;
}

function register_(d) {
  var name = String(d.name || '').trim(), area = String(d.area || '').trim();
  var phone = digits_(d.phone), parent = digits_(d.parent);
  if (/^\d{8}$/.test(phone)) phone = '974' + phone;
  if (/^\d{8}$/.test(parent)) parent = '974' + parent;
  function okPhone(p) { return p.indexOf('974') === 0 ? /^974\d{8}$/.test(p) : /^\d{8,15}$/.test(p); }
  var logd = { name: name, phone: phone, parent: parent };
  if (name.length < 3 || name.length > 80 || area.length < 2 || area.length > 80 || !okPhone(phone) || !okPhone(parent)) {
    logReq_('تسجيل', logd, 'رُفض: بيانات غير صحيحة');
    return { ok: false, error: 'bad data' };
  }
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('الطلاب');
  if (!sh) { logReq_('تسجيل', logd, 'خطأ: شيت الطلاب غير موجود'); return { ok: false, error: 'no sheet' }; }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var cName = ensureCol_(sh, ['اسم الطالب', 'الاسم', 'name'], 'اسم الطالب');
    var cPhone = ensureCol_(sh, ['هاتف الطالب', 'phone', 'الهاتف'], 'هاتف الطالب');
    var cParent = ensureCol_(sh, ['هاتف ولي الأمر', 'parent'], 'هاتف ولي الأمر');
    var cArea = ensureCol_(sh, ['مكان السكن', 'منطقة السكن', 'المنطقة'], 'مكان السكن');
    var cStatus = ensureCol_(sh, ['الحالة', 'status'], 'الحالة');
    var cDate = ensureCol_(sh, ['تاريخ التسجيل'], 'تاريخ التسجيل');
    var n = sh.getLastRow();
    if (n > 1) {
      var ph = sh.getRange(2, cPhone, n - 1, 1).getValues();
      for (var i = 0; i < ph.length; i++) {
        if (key_(ph[i][0]) === key_(phone)) {
          // الرقم موجود: لا نضيف صفًا جديدًا، لكن نترك أثرًا واضحًا
          var r0 = i + 2;
          var samePar = key_(sh.getRange(r0, cParent).getValue()) === key_(parent);
          var st = String(sh.getRange(r0, cStatus).getValue()).trim();
          var cNote = ensureCol_(sh, ['ملاحظات'], 'ملاحظات');
          var cell = sh.getRange(r0, cNote), old = String(cell.getValue() || '');
          cell.setNumberFormat('@');
          cell.setValue((old ? old + ' | ' : '') + 'طلب تسجيل مكرر ' + stamp_());
          logReq_('تسجيل', logd, 'مكرر — الرقم موجود في الصف ' + r0 + (st === '2' ? ' (مفعّل)' : ' (غير مفعّل)'));
          return { ok: false, error: samePar ? (st === '2' ? 'active' : 'pending') : 'exists' };
        }
      }
    }
    var row = n + 1, lc = sh.getLastColumn(), rng = sh.getRange(row, 1, 1, lc);
    rng.setNumberFormat('@');
    sh.getRange(row, cName).setValue(name);
    sh.getRange(row, cPhone).setValue(phone);
    sh.getRange(row, cParent).setValue(parent);
    sh.getRange(row, cArea).setValue(area);
    sh.getRange(row, cDate).setValue(stamp_());
    rng.setBackground('#fff2cc');
    logReq_('تسجيل', logd, 'تمت الإضافة — الصف ' + row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// عند كتابتك 2 في عمود "الحالة" يزول اللون الأصفر تلقائيًا + فحص تكرار هاتف الطالب
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() === 'الطلاب' && e.range.getRow() >= 2) {
      var c = ensureCol_(sh, ['الحالة', 'status'], '');
      if (c && e.range.getColumn() === c && String(e.value).trim() === '2') sh.getRange(e.range.getRow(), 1, 1, sh.getLastColumn()).setBackground(null);
    }
  } catch (err) {}
  try { dupCheckOnEdit_(e); } catch (err2) {}
  try { reminderEdit_(e); } catch (err3) {}
}

// ---------- منع تكرار هاتف الطالب في شيت "الطلاب" ----------
function phoneCol_(sh) { return ensureCol_(sh, ['هاتف الطالب', 'phone', 'الهاتف'], ''); }

// يلوّن الأرقام المكررة بالأحمر ويزيل اللون عن غير المكررة؛ يرجع صفوف المكرر
function markDuplicates_(sh) {
  var cPhone = phoneCol_(sh); if (!cPhone) return [];
  var n = sh.getLastRow(); if (n < 2) return [];
  var rng = sh.getRange(2, cPhone, n - 1, 1), vals = rng.getValues(), bgs = rng.getBackgrounds();
  var first = sh.getRange(2, 1, n - 1, 1).getBackgrounds();
  var count = {};
  vals.forEach(function (v) { var k = key_(v[0]); if (k) count[k] = (count[k] || 0) + 1; });
  var dupRows = [], changed = false;
  for (var i = 0; i < vals.length; i++) {
    var k = key_(vals[i][0]);
    if (k && count[k] > 1) {
      dupRows.push(i + 2);
      if (bgs[i][0] !== DUP_COLOR) { bgs[i][0] = DUP_COLOR; changed = true; }
    } else if (bgs[i][0] === DUP_COLOR) {
      bgs[i][0] = (first[i][0] && first[i][0] !== DUP_COLOR) ? first[i][0] : '#ffffff'; changed = true;
    }
  }
  if (changed) rng.setBackgrounds(bgs);
  return dupRows;
}

// يحدّد أماكن الأعمدة المعروفة في شيت الطلاب من صف العناوين h (يرجع -1 لغير الموجود)
function studentCols_(h) {
  function col(names) { for (var i = 0; i < names.length; i++) { var ix = h.indexOf(names[i]); if (ix > -1) return ix; } return -1; }
  var c = {
    name: col(['اسم الطالب', 'الاسم', 'name']), phone: col(['هاتف الطالب', 'phone', 'الهاتف']), parent: col(['هاتف ولي الأمر', 'parent']),
    area: col(['مكان السكن', 'منطقة السكن', 'المنطقة']), status: col(['الحالة', 'status']), date: col(['تاريخ التسجيل']), note: col(['ملاحظات']),
    group: col(['المجموعة', 'اسم المجموعة', 'المجموعه', 'group'])
  };
  // عمود المجموعة: بالاسم، وإلا العمود C بشرط ألا يكون عنوانه حقلًا معروفًا آخر
  if (c.group < 0 && h.length >= 3 && ['اسم الطالب', 'الاسم', 'name', 'هاتف الطالب', 'phone', 'الهاتف', 'هاتف ولي الأمر', 'parent', 'الحالة', 'status', 'تاريخ التسجيل', 'ملاحظات', 'مكان السكن', 'منطقة السكن', 'المنطقة', ''].indexOf(h[2]) === -1) c.group = 2;
  return c;
}

// يقرأ شيت "الطلاب" ويرجع مجموعات الأرقام المكررة (كل مجموعة = صفوف بنفس هاتف الطالب بعد توحيد الصيغة)
// كل صف يحمل vals = كل أعمدة الصف كما هي. والمصفوفة المرجعة عليها: headers, cols, groupLabel
function dupGroups_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('الطلاب');
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return (x + '').trim(); });
  var cols = studentCols_(h);
  if (cols.phone < 0) return [];
  function s(i, c) { return c > -1 && v[i][c] != null ? String(v[i][c]).trim() : ''; }
  var by = {};
  for (var i = 1; i < v.length; i++) {
    var k = key_(v[i][cols.phone]); if (!k) continue;
    (by[k] = by[k] || []).push({ row: i + 1, vals: v[i].slice(), name: s(i, cols.name), phone: s(i, cols.phone), parent: s(i, cols.parent), area: s(i, cols.area),
      status: s(i, cols.status), date: s(i, cols.date), note: s(i, cols.note), group: s(i, cols.group) });
  }
  var out = [];
  for (var k2 in by) if (by[k2].length > 1) out.push(by[k2]);
  out.sort(function (a, b) { return a[0].row - b[0].row; });
  out.headers = h; out.cols = cols;
  out.groupLabel = cols.group > -1 ? h[cols.group] : '';
  return out;
}

// تخطيط تبويب "أرقام متكررة": A ✔ | B مقترح | C الصف | D مفتاح (مخفي) | E اكتمال | F… كل أعمدة شيت الطلاب | ثم "نتائج مسجلة"
var DUP_FIRST = 5;   // فهرس (من 0) أول عمود من أعمدة الطلاب

// يقرأ التبويب ويرجع المجموعات حسب مفتاح الرقم: [{key, m, rows:[{r, vals, checked}]}]
function parseDupTab_(tab) {
  var v = tab.getDataRange().getValues();
  if (v.length < 3) return [];
  var resIdx = v[1].indexOf('نتائج مسجلة');
  if (resIdx < DUP_FIRST) return [];
  var m = resIdx - DUP_FIRST, by = {}, order = [];
  for (var i = 2; i < v.length; i++) {
    var key = String(v[i][3] == null ? '' : v[i][3]).trim(), rn = v[i][2];
    if (!key || !/^\d+$/.test(String(rn).trim())) continue;
    if (!by[key]) { by[key] = { key: key, m: m, rows: [] }; order.push(key); }
    by[key].rows.push({ r: Number(rn), vals: v[i].slice(DUP_FIRST, DUP_FIRST + m), checked: v[i][0] === true || String(v[i][0]).toUpperCase() === 'TRUE' });
  }
  return order.map(function (k) { return by[k]; });
}

// يزيل تجميع الصفوف السابق من التبويب قبل إعادة بنائه (clear() لا يزيله)
function clearGroups_(sh, upTo) {
  try {
    var start = 0;
    for (var r = 1; r <= upTo + 1; r++) {
      var d = r <= upTo ? sh.getRowGroupDepth(r) : 0;
      if (d > 0 && !start) start = r;
      if (d === 0 && start) { sh.getRange(start, 1, r - start, 1).shiftRowGroupDepth(-1); start = 0; }
    }
  } catch (e) {}
}

function isEmptyCell_(x) { return x === '' || x == null || String(x).trim() === ''; }

// يبني/يحدّث تبويب "أرقام متكررة": كل رقم وتحته مباشرة صفوفه بكل تفاصيلها. لا يُنشأ إلا عند وجود تكرار، ولا يحذف أي صف
function buildDupSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), groups = dupGroups_(), gl = groups.groupLabel || 'المجموعة';
  var oldTab = ss.getSheetByName(DUP_SHEET_OLD);          // تبويب النسخة السابقة (يُنشأ تلقائيًا) — يُستبدل بالجديد
  if (oldTab) { try { ss.deleteSheet(oldTab); } catch (e0) {} }
  var out = ss.getSheetByName(DUP_SHEET);
  if (!groups.length && !out) return 0;
  var keep = {};                                           // نحافظ على علامات ✔ التي وضعتها عند إعادة البناء
  if (out) { try { parseDupTab_(out).forEach(function (g) { g.rows.forEach(function (x) { if (x.checked) keep[g.key + '|' + x.r] = true; }); }); } catch (e1) {} }
  if (!out) out = ss.insertSheet(DUP_SHEET);
  clearGroups_(out, out.getLastRow());
  out.clear();
  if (!groups.length) { out.getRange(1, 1).setValue('✅ لا توجد أرقام مكررة (آخر فحص ' + stamp_() + ')'); return 0; }
  var h = groups.headers, cols = groups.cols, m = h.length, N = DUP_FIRST + m + 1;
  function blank() { var a = []; for (var q = 0; q < N; q++) a.push(''); return a; }
  var counts = {}, rs = ss.getSheetByName(RESULTS_SHEET);
  if (rs && rs.getLastRow() > 1) rs.getRange(2, 3, rs.getLastRow() - 1, 1).getValues().forEach(function (r) { var k = key_(r[0]); if (k) counts[k] = (counts[k] || 0) + 1; });
  var total = m - (cols.note > -1 ? 1 : 0);                // الخانات التي تُحسب في الاكتمال (بدون عمود الملاحظات)
  function filled(x) { var n = 0; for (var j = 0; j < m; j++) if (j !== cols.note && !isEmptyCell_(x.vals[j])) n++; return n; }
  var rows = [], kinds = [], spans = [], colors = [], curStart = 0;
  var t = blank(); t[0] = '① ضع ✔ أمام الصف الذي تريد بقاءه في كل رقم   ② من القائمة 📊 نجم اختر «حذف المكرر (احتفظ بالمُحدّد ✔)»   ③ يُحذف الباقي بعد تأكيدك، وتُحفظ نسخة منه في تبويب «' + DEL_SHEET + '»   (⭐ = المقترح: الأكمل معلومات، ثم المفعّل، ثم الأقدم — والخانات الفارغة بلون وردي)';
  rows.push(t); kinds.push('title'); colors.push(null);
  var hd = ['✔ احتفظ', 'مقترح', 'الصف', 'مفتاح', 'اكتمال'];
  for (var j0 = 0; j0 < m; j0++) hd.push(h[j0] || ('عمود ' + (j0 + 1)));
  hd.push('نتائج مسجلة');
  rows.push(hd); kinds.push('head'); colors.push(null);
  groups.forEach(function (g) {
    var key = key_(g[0].phone), best = g[0];
    g.forEach(function (x) {
      var fx = filled(x), fb = filled(best);
      if (fx > fb || (fx === fb && x.status === '2' && best.status !== '2')) best = x;
    });
    var gset = {}; g.forEach(function (x) { if (x.group) gset[x.group] = 1; });
    var diffGroups = Object.keys(gset).length > 1;
    var b = blank(); b[DUP_FIRST - 1 + 0] = '📞 ' + g[0].phone + '   —   ' + g.length + ' صفوف متشابهة' + (diffGroups ? '   —   ⚠️ في مجموعات مختلفة' : '');
    rows.push(b); kinds.push('block'); colors.push(null);
    curStart = rows.length + 1;                       // أول صف بيانات تحت عنوان هذا الرقم (رقم الصف في الشيت)
    g.forEach(function (x) {
      var fx = filled(x);
      var r = [!!keep[key + '|' + x.row], x === best ? '⭐' : '', x.row, key, fx + ' من ' + total];
      var rc = ['#fff2cc', '#fff2cc', '#fff2cc', '#fff2cc', '#fff2cc'];
      for (var j = 0; j < m; j++) { r.push(x.vals[j] == null ? '' : x.vals[j]); rc.push(j !== cols.note && isEmptyCell_(x.vals[j]) ? '#f4cccc' : '#fff2cc'); }
      r.push(counts[key_(x.phone)] || 0); rc.push('#fff2cc');
      rows.push(r); kinds.push('row'); colors.push(rc);
    });
    spans.push([curStart, g.length]);
    rows.push(blank()); kinds.push('gap'); colors.push(null);
  });
  // تنسيق نصي لأعمدة الهاتف والمفتاح قبل كتابة القيم (حتى لا تضيع الأصفار/الصيغة)
  out.getRange(1, 4, rows.length, 1).setNumberFormat('@');
  for (var jp = 0; jp < m; jp++) if (/هاتف|phone/i.test(String(h[jp]))) out.getRange(1, DUP_FIRST + jp + 1, rows.length, 1).setNumberFormat('@');
  out.getRange(1, 1, rows.length, N).setValues(rows);
  out.getRange(1, 1, 1, N).setFontWeight('bold');
  out.getRange(2, 1, 1, N).setFontWeight('bold');
  for (var i = 0; i < kinds.length; i++) {
    var rr = i + 1;
    if (kinds[i] === 'head') out.getRange(rr, 1, 1, N).setBackground('#cfe2f3');
    else if (kinds[i] === 'block') { out.getRange(rr, 1, 1, N).setBackground('#f5c542'); out.getRange(rr, 1, 1, N).setFontWeight('bold'); }
    else if (kinds[i] === 'row') { out.getRange(rr, 1, 1, N).setBackgrounds([colors[i]]); out.getRange(rr, 1).insertCheckboxes(); }
  }
  try { out.hideColumns(4); } catch (e2) {}
  try { // تجميع صفوف كل رقم: [−]/[+] بجانب عنوانه (العنوان فوق الصفوف)
    spans.forEach(function (s) { out.getRange(s[0], 1, s[1], 1).shiftRowGroupDepth(1); });
    out.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  } catch (e3) {}
  out.setFrozenRows(2);
  return groups.length;
}

// من القائمة: يحذف من شيت "الطلاب" كل الصفوف غير المعلَّمة ✔ في تبويب "أرقام متكررة" (بعد تأكيدك + نسخة احتياطية)
function applyDuplicateChoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var tab = ss.getSheetByName(DUP_SHEET), st = ss.getSheetByName('الطلاب');
  if (!tab || !st) { ui.alert('لا يوجد تبويب "' + DUP_SHEET + '" الآن (أي لا توجد أرقام مكررة أو لم يُفحص بعد). افتح 📊 نجم ← فحص الأرقام المكررة أولًا.'); return; }
  var cPhone = phoneCol_(st);
  if (!cPhone) { ui.alert('لم أجد عمود هاتف الطالب في شيت الطلاب'); return; }
  var hdr = st.getRange(1, 1, 1, st.getLastColumn()).getValues()[0].map(function (x) { return (x + '').trim(); });
  var cols = studentCols_(hdr);
  var groups = parseDupTab_(tab), plan = [], none = 0, multi = 0, changed = 0, kept = 0;
  groups.forEach(function (g) {
    var ch = g.rows.filter(function (x) { return x.checked; });
    if (ch.length === 0) { none++; return; }
    if (ch.length > 1) { multi++; return; }
    var ok = g.rows.every(function (x) {   // تأكد أن الصف في الشيت ما زال هو نفسه تمامًا كما يظهر في التبويب
      var cur = st.getRange(x.r, 1, 1, Math.max(g.m, 1)).getValues()[0];
      if (key_(cur[cols.phone]) !== g.key) return false;
      for (var j = 0; j < g.m; j++) if (String(cur[j] == null ? '' : cur[j]).trim() !== String(x.vals[j] == null ? '' : x.vals[j]).trim()) return false;
      return true;
    });
    if (!ok) { changed++; return; }
    kept++;
    g.rows.forEach(function (x) {
      if (!x.checked) plan.push({ r: x.r, name: cols.name > -1 ? String(x.vals[cols.name] || '').trim() : '', group: cols.group > -1 ? String(x.vals[cols.group] || '').trim() : '', key: g.key });
    });
  });
  var notes = (none ? '\n• ' + none + ' رقم لم تحدد له صفًا (تُرك كما هو).' : '') + (multi ? '\n• ' + multi + ' رقم حددت فيه أكثر من صف (تُرك كما هو).' : '') + (changed ? '\n• ' + changed + ' رقم تغيّرت بياناته بعد بناء التبويب (تُرك — حدّث التبويب بالفحص ثم أعد المحاولة).' : '');
  if (!plan.length) { ui.alert('لا يوجد ما يُحذف.' + notes); return; }
  var lines = plan.slice(0, 15).map(function (p) { return '• الصف ' + p.r + ' — ' + (p.name || 'بلا اسم') + (p.group ? ' — ' + p.group : ''); }).join('\n') + (plan.length > 15 ? '\n… و' + (plan.length - 15) + ' آخرين' : '');
  var resp = ui.alert('تأكيد الحذف', 'سيُحذف ' + plan.length + ' صف من شيت «الطلاب»، ويبقى صف واحد لكل رقم (' + kept + ' رقم):\n' + lines + notes + '\n\nتُحفظ نسخة من المحذوف في تبويب «' + DEL_SHEET + '». هل تريد المتابعة؟', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var lc = st.getLastColumn(), hdr2 = st.getRange(1, 1, 1, lc).getValues()[0];
    plan.sort(function (a, b) { return b.r - a.r; });         // من الأسفل للأعلى حتى لا تتحرك أرقام الصفوف
    var saved = plan.map(function (p) { return [p.r].concat(st.getRange(p.r, 1, 1, lc).getValues()[0]); });
    var bk = ss.getSheetByName(DEL_SHEET) || ss.insertSheet(DEL_SHEET);
    bk.appendRow(['— حذف بتاريخ ' + stamp_() + ' — الأعمدة التالية: الصف الأصلي ثم أعمدة شيت الطلاب —'].concat(hdr2));
    saved.forEach(function (row) { bk.appendRow(['محذوف'].concat(row)); });
    plan.forEach(function (p) { st.deleteRow(p.r); });
  } finally { lock.releaseLock(); }
  markDuplicates_(st);
  buildDupSheet_();
  ss.toast('تم حذف ' + plan.length + ' صف والاحتفاظ بصف واحد لكل رقم (' + kept + '). نسخة المحذوف في تبويب «' + DEL_SHEET + '».', 'نجم', 20);
}

function dupCheckOnEdit_(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== 'الطلاب' || e.range.getRow() < 2) return;
  var cPhone = phoneCol_(sh); if (!cPhone) return;
  if (cPhone < e.range.getColumn() || cPhone > e.range.getLastColumn()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  var r = e.range.getRow(), v = e.value;
  var dups = markDuplicates_(sh);                 // يلوّن كل الصفوف المتشابهة بالأحمر (ويُزيل اللون عن غير المكرر)
  var n = buildDupSheet_();                        // يحدّث تبويب "الأرقام المكررة"
  if (!dups.length) return;
  if (single && dups.indexOf(r) > -1 && String(v == null ? '' : v).trim() !== '') {
    // إدخال يدوي لرقم موجود: لا نحذف شيئًا — نخبرك بالصفوف المتشابهة لتقرر
    var k = key_(v), same = [];
    dupGroups_().forEach(function (g) { if (key_(g[0].phone) === k) same = g; });
    var lines = same.map(function (x) { return '• الصف ' + x.row + ' — ' + (x.name || 'بلا اسم') + ' — ' + (x.status === '2' ? 'مفعّل' : 'غير مفعّل'); }).join('\n');
    var msg = '⚠️ رقم الهاتف (' + String(v).trim() + ') مكرر في الصفوف:\n' + lines + '\n\nلم أحذف شيئًا. راجع الصفوف وقرّر أيها تحتفظ به (التفاصيل في تبويب "' + DUP_SHEET + '").';
    try { SpreadsheetApp.getUi().alert('رقم مكرر', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
    catch (x) { ss.toast(msg, 'نجم — رقم مكرر', 30); }
  } else {
    ss.toast('⚠️ يوجد ' + n + ' رقم/أرقام مكررة (الصفوف: ' + dups.join('، ') + ') — ملوّنة بالأحمر. التفاصيل في تبويب "' + DUP_SHEET + '".', 'نجم — تنبيه', 30);
  }
}

// من القائمة: فحص يدوي — يلوّن، ويحدّث تبويب "الأرقام المكررة"، ويفتحه
function checkDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('الطلاب');
  if (!sh) { SpreadsheetApp.getUi().alert('لا يوجد شيت اسمه الطلاب'); return; }
  markDuplicates_(sh);
  var n = buildDupSheet_();
  if (n) {
    ss.setActiveSheet(ss.getSheetByName(DUP_SHEET));
    SpreadsheetApp.getUi().alert('⚠️ يوجد ' + n + ' رقم/أرقام مكررة.\nفي تبويب "' + DUP_SHEET + '": كل رقم وتحته صفوفه. ضع ✔ أمام الصف الذي تريد بقاءه، ثم من القائمة 📊 نجم اختر «حذف المكرر (احتفظ بالمُحدّد ✔)».');
  } else {
    SpreadsheetApp.getUi().alert('✅ لا توجد أرقام مكررة');
  }
}

// ================== تذكير الأسبوع للمشتركين ==================
function today_() { return Utilities.formatDate(new Date(), 'Asia/Qatar', 'yyyy-MM-dd'); }
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Qatar', 'yyyy-MM-dd') : String(v == null ? '' : v).trim().slice(0, 10); }
function daysBetween_(a, b) {
  var pa = String(a).split('-'), pb = String(b).split('-');
  return Math.round((Date.UTC(+pb[0], +pb[1] - 1, +pb[2]) - Date.UTC(+pa[0], +pa[1] - 1, +pa[2])) / 86400000);
}
// مفتاح الملف: الرابط بدون ?... و #... وبتوحيد الاسم القديم/الجديد
function itemKey_(link) { return String(link == null ? '' : link).trim().toLowerCase().split('#')[0].split('?')[0].replace(/_p1_(p[23])_/, '_$1_'); }
function waPhone_(k) { k = digits_(k); return k.length === 8 ? '974' + k : k; }

// الملفات الفعّالة الظاهرة للمشتركين: ظهور 1 أو 2 + رابط فعلي + رقم وحدة (كما يفعل الموقع)
function contentRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('المحتوى');
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return (x + '').trim(); });
  function col(names) { for (var i = 0; i < names.length; i++) { var ix = h.indexOf(names[i]); if (ix > -1) return ix; } return -1; }
  var cLevel = col(['ظهور', 'level']), cLink = col(['الرابط', 'link']), cTitle = col(['العنوان', 'title']), cUnit = col(['الوحدة', 'unit']);
  if (cLink < 0 || cTitle < 0) return [];
  var out = [], seen = {};
  for (var i = 1; i < v.length; i++) {
    var raw = cLevel > -1 ? String(v[i][cLevel] == null ? '' : v[i][cLevel]).trim() : '';
    if (raw === '00') continue;
    var lv = parseInt(raw, 10); if (isNaN(lv)) lv = 1;
    if (lv !== 1 && lv !== 2) continue;
    var title = String(v[i][cTitle] == null ? '' : v[i][cTitle]).trim(), link = String(v[i][cLink] == null ? '' : v[i][cLink]).trim();
    if (!title || !link || link === '#') continue;
    if (cUnit > -1 && !/\d/.test(String(v[i][cUnit] == null ? '' : v[i][cUnit]))) continue;
    var key = itemKey_(link); if (seen[key]) continue; seen[key] = 1;
    out.push({ title: title, link: link, key: key, level: lv });
  }
  return out;
}

// يسجّل أول ظهور لكل ملف جديد. أول مرة (أو إن كان السجل فارغًا): كل الملفات الموجودة تُعتبر "قديمة" ولا يُذكَّر بها
function syncContentLog_(force) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('lastContentSync') || 0);
  if (!force && ss.getSheetByName(CONTENT_LOG) && Date.now() - last < 10 * 60 * 1000) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = ss.getSheetByName(CONTENT_LOG);
    if (!sh) {
      sh = ss.insertSheet(CONTENT_LOG);
      sh.appendRow(['المفتاح', 'العنوان', 'أول ظهور', 'قديم (لا يُذكَّر به)']);
      sh.getRange('A:C').setNumberFormat('@');
      sh.setFrozenRows(1);
    }
    var fresh = sh.getLastRow() < 2, known = {};
    if (!fresh) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { known[String(r[0])] = 1; });
    var add = [], t = today_();
    contentRows_().forEach(function (r) { if (!known[r.key]) { known[r.key] = 1; add.push([r.key, r.title, t, fresh]); } });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setValues(add);
    props.setProperty('lastContentSync', String(Date.now()));
  } finally { lock.releaseLock(); }
}

function loadReminderData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), log = {}, active = {}, opens = {}, reminded = {};
  var sh = ss.getSheetByName(CONTENT_LOG);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
    log[String(r[0])] = { title: String(r[1]), firstSeen: dstr_(r[2]), baseline: r[3] === true || String(r[3]).toUpperCase() === 'TRUE' };
  });
  contentRows_().forEach(function (r) { active[r.key] = r; });
  var ob = ss.getSheetByName(OPENS_SHEET);
  if (ob && ob.getLastRow() > 1) ob.getRange(2, 1, ob.getLastRow() - 1, 2).getValues().forEach(function (r) { opens[String(r[0]) + '|' + String(r[1])] = 1; });
  var rb = ss.getSheetByName(REMINDED_SHEET);
  if (rb && rb.getLastRow() > 1) rb.getRange(2, 1, rb.getLastRow() - 1, 1).getValues().forEach(function (r) { reminded[String(r[0])] = 1; });
  return { log: log, active: active, opens: opens, reminded: reminded };
}

// ملفات لم يفتحها الطالب (pk = مفتاح هاتفه) مرّ عليها أسبوع أو أكثر (وليس أكثر من 45 يومًا)
function pendingItems_(pk, d, skipReminded) {
  var out = [], t = today_();
  for (var key in d.active) {
    var l = d.log[key];
    if (!l || l.baseline || !l.firstSeen) continue;
    var days = daysBetween_(l.firstSeen, t);
    if (days < REMIND_DAYS || days > REMIND_WINDOW_DAYS) continue;
    if (d.opens[pk + '|' + key]) continue;
    if (skipReminded && d.reminded[pk + '|' + key]) continue;
    out.push({ key: key, title: d.active[key].title, link: d.active[key].link, days: days });
  }
  out.sort(function (a, b) { return b.days - a.days; });
  return out;
}

// للموقع: ما الذي لم يفتحه هذا الطالب؟ (يتحقق من الرقمين والحالة 2)
function pending_(phone, parent) {
  var pk = key_(phone), qk = key_(parent);
  if (!pk || !qk) return { ok: false, error: 'bad input' };
  var me = students_().filter(function (s) { return s.phone === pk && s.parent === qk; })[0];
  if (!me) return { ok: false, error: 'not found' };
  if (!me.active) return { ok: false, error: 'pending' };
  try { syncContentLog_(false); } catch (e) {}
  var items = pendingItems_(pk, loadReminderData_(), false).slice(0, 5).map(function (i) { return { title: i.title, link: i.link, days: i.days }; });
  return { ok: true, items: items };
}

// من الموقع: الطالب فتح ملفًا (أول مرة فقط تُسجَّل)
function open_(d) {
  var pk = key_(d.phone), qk = key_(d.parent), link = String(d.link || '');
  if (!pk || !qk || !link) return { ok: false, error: 'bad input' };
  var me = students_().filter(function (s) { return s.phone === pk && s.parent === qk; })[0];
  if (!me || !me.active) return { ok: false, error: 'not allowed' };
  var key = itemKey_(link), valid = false;
  contentRows_().forEach(function (r) { if (r.key === key) valid = true; });
  if (!valid) return { ok: true };                       // ليس من محتوى الموقع: نتجاهله
  var ss = SpreadsheetApp.getActiveSpreadsheet(), lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = ss.getSheetByName(OPENS_SHEET);
    if (!sh) { sh = ss.insertSheet(OPENS_SHEET); sh.appendRow(['هاتف الطالب', 'المفتاح', 'أول فتح', 'العنوان']); sh.getRange('A:B').setNumberFormat('@'); sh.setFrozenRows(1); }
    var n = sh.getLastRow();
    if (n > 1) { var v = sh.getRange(2, 1, n - 1, 2).getValues(); for (var i = 0; i < v.length; i++) if (String(v[i][0]) === pk && String(v[i][1]) === key) return { ok: true }; }
    sh.appendRow([pk, key, stamp_(), String(d.title || '').slice(0, 100)]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function reminderMsg_(name, items) {
  var first = String(name || '').trim().split(/\s+/)[0] || 'طالبنا';
  var n = items.length, titles = items.slice(0, 3).map(function (i) { return '«' + i.title + '»'; }).join('، ') + (n > 3 ? ' و' + (n - 3) + ' أخرى' : '');
  var body = n === 1 ? 'تم إضافة ملف جديد من أسبوع لم تطلع عليه: ' + titles + '. افتح الملف وابدأ بحله.'
                     : 'تم إضافة ' + n + ' ملفات جديدة من أسبوع لم تطلع عليها: ' + titles + '. افتح الملفات وابدأ بحلها.';
  return 'السلام عليكم ' + first + '،\n' + body + '\n— أ. عزت نجم الأحياء ⭐';
}

// يبني شيت "تذكيرات الأسبوع": كل طالب عليه تذكير + نص الرسالة + زر واتساب
function weeklyReminders_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), d = loadReminderData_(), rows = [];
  students_().forEach(function (s) {
    if (!s.active) return;
    var items = pendingItems_(s.phone, d, true);
    if (items.length) rows.push({ s: s, items: items });
  });
  var sh = ss.getSheetByName(REMINDERS_SHEET);
  if (!rows.length && !sh) return 0;
  if (!sh) sh = ss.insertSheet(REMINDERS_SHEET);
  sh.clear();
  if (!rows.length) { sh.getRange(1, 1).setValue('✅ لا يوجد طلاب عليهم تذكير الآن (آخر تحديث ' + stamp_() + ')'); return 0; }
  var out = [['✔ أُرسل', 'اسم الطالب', 'هاتف الطالب', 'هاتف ولي الأمر', 'عدد الملفات', 'الملفات', 'نص الرسالة', 'واتساب الطالب', 'واتساب ولي الأمر', 'المفاتيح']];
  rows.forEach(function (r) {
    var msg = reminderMsg_(r.s.name, r.items), enc = encodeURIComponent(msg);
    var f1 = '=HYPERLINK("https://wa.me/' + waPhone_(r.s.phone) + '?text=' + enc + '","📲 أرسل للطالب")';
    var f2 = r.s.parent ? '=HYPERLINK("https://wa.me/' + waPhone_(r.s.parent) + '?text=' + enc + '","📲 أرسل لولي الأمر")' : '';
    out.push([false, r.s.name, waPhone_(r.s.phone), r.s.parent ? waPhone_(r.s.parent) : '', r.items.length, r.items.map(function (i) { return i.title + ' (منذ ' + i.days + ' يومًا)'; }).join(' | '), msg, f1, f2,
      r.items.map(function (i) { return r.s.phone + '|' + i.key; }).join(';')]);
  });
  sh.getRange(1, 3, out.length, 2).setNumberFormat('@');
  sh.getRange(1, 10, out.length, 1).setNumberFormat('@');
  sh.getRange(1, 1, out.length, 10).setValues(out);
  sh.getRange(1, 1, 1, 10).setFontWeight('bold');
  sh.getRange(1, 1, 1, 10).setBackground('#cfe2f3');
  for (var i = 2; i <= out.length; i++) sh.getRange(i, 1).insertCheckboxes();
  try { sh.hideColumns(10); } catch (e) {}
  sh.setFrozenRows(1);
  return rows.length;
}

// ✔ أمام الطالب = أُرسلت له الرسالة: تُحفظ مفاتيح ملفاته كي لا يُذكَّر بها مرة أخرى
function reminderEdit_(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== REMINDERS_SHEET || e.range.getColumn() !== 1 || e.range.getRow() < 2) return;
  if (String(e.value).toUpperCase() !== 'TRUE') return;
  var r = e.range.getRow(), keys = String(sh.getRange(r, 10).getValue() || '').split(';').filter(Boolean);
  if (!keys.length) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet(), rb = ss.getSheetByName(REMINDED_SHEET);
  if (!rb) { rb = ss.insertSheet(REMINDED_SHEET); rb.appendRow(['هاتف الطالب|المفتاح', 'تاريخ الإرسال']); rb.getRange('A:A').setNumberFormat('@'); rb.setFrozenRows(1); }
  var t = stamp_();
  keys.forEach(function (k) { rb.appendRow([k, t]); });
  sh.getRange(r, 1, 1, 10).setBackground('#d9ead3');
}

function dailyJob() { try { syncContentLog_(true); } catch (e) {} weeklyReminders_(); }

function installReminderTrigger() {
  var ui = SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(9).create();
  dailyJob();
  ui.alert('✅ تم تفعيل تذكير الأسبوع.\nيعمل تلقائيًا كل يوم قرابة الساعة 9 صباحًا: يسجّل الملفات الجديدة، ويجهّز في تبويب "' + REMINDERS_SHEET + '" الطلاب الذين مرّ أسبوع على ملف لم يفتحوه.\nالملفات الموجودة الآن تُعتبر قديمة ولا يُذكَّر بها؛ التذكير للملفات التي تُضاف من الآن فصاعدًا.');
}

function refreshReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try { syncContentLog_(true); } catch (e) {}
  var n = weeklyReminders_(), sh = ss.getSheetByName(REMINDERS_SHEET);
  if (sh) ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(n ? 'يوجد ' + n + ' طالب/طلاب عليهم تذكير. في تبويب "' + REMINDERS_SHEET + '": اضغط زر واتساب ثم ضع ✔ بعد الإرسال.' : '✅ لا يوجد طلاب عليهم تذكير الآن.');
}

// ---------- للمعلم: قائمة 📊 نجم ← تحديث إحصاء الأسئلة (شيت "إحصاء الأسئلة") ----------
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('📊 نجم').addItem('تحديث إحصاء الأسئلة', 'buildQuestionStats').addItem('فحص الأرقام المكررة', 'checkDuplicates').addItem('حذف المكرر (احتفظ بالمُحدّد ✔)', 'applyDuplicateChoices').addItem('تفعيل تذكير الأسبوع (يوميًا)', 'installReminderTrigger').addItem('تحديث تذكيرات الأسبوع الآن', 'refreshReminders').addToUi(); } catch (e) {}
  try { // تنبيه عند فتح الملف إن وُجد تكرار
    var gs = dupGroups_();
    if (gs.length) SpreadsheetApp.getActiveSpreadsheet().toast('⚠️ يوجد ' + gs.length + ' رقم/أرقام مكررة في شيت الطلاب — افتح 📊 نجم ← فحص الأرقام المكررة', 'نجم — تنبيه', 20);
  } catch (e2) {}
}

function buildQuestionStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(RESULTS_SHEET);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('لا توجد نتائج بعد'); return; }
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues(), best = {};
  v.forEach(function (r) { // أفضل محاولة لكل طالب في كل اختبار
    var k = key_(r[2]) + '|' + normExam_(r[4]), pct = Number(r[8]);
    if (!best[k] || pct > best[k].pct) best[k] = { ex: normExam_(r[4]), title: r[5], pct: pct, wrong: parseWrong_(r[10]) };
  });
  var ex = {};
  for (var k in best) {
    var b = best[k];
    if (b.wrong === null) continue; // محاولات قديمة بلا تفاصيل
    var e = ex[b.ex] = ex[b.ex] || { title: b.title, n: 0, qs: {} };
    e.n++;
    b.wrong.forEach(function (w) {
      var o = e.qs[w.q] = e.qs[w.q] || { miss: 0, c: w.c, ans: {} };
      o.miss++;
      if (w.a) o.ans[w.a] = (o.ans[w.a] || 0) + 1;
    });
  }
  var rows = [['الاختبار', 'السؤال', 'عدد من أخطأ', 'عدد من حلّ الاختبار', 'نسبة الخطأ %', 'الإجابة الصحيحة', 'أكثر إجابة خاطئة تكرارًا']];
  Object.keys(ex).sort().forEach(function (id) {
    var e = ex[id];
    Object.keys(e.qs).map(function (q) { return { q: q, o: e.qs[q] }; })
      .sort(function (a, b) { return b.o.miss - a.o.miss; })
      .forEach(function (x) {
        var top = '', tc = 0;
        for (var a in x.o.ans) if (x.o.ans[a] > tc) { tc = x.o.ans[a]; top = a; }
        rows.push([e.title, x.q, x.o.miss, e.n, Math.round(x.o.miss / e.n * 1000) / 10, x.o.c, top]);
      });
  });
  var out = ss.getSheetByName('إحصاء الأسئلة') || ss.insertSheet('إحصاء الأسئلة');
  out.clear();
  out.getRange(1, 1, rows.length, 7).setValues(rows);
  out.setFrozenRows(1);
  out.getRange(1, 1, 1, 7).setFontWeight('bold');
}
