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
 */

var RESULTS_SHEET = 'النتائج';

function doGet(e) {
  var p = e.parameter || {};
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
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
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
      var k = key_(v[i][2]), ex = v[i][4], pct = Number(v[i][8]);
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
  return { ok: true, name: me.name, exams: exams, overall: mine.pct === null ? null : Math.round(mine.pct * 100) / 100, taken: mine.taken, topPct: topPct, n: n };
}

// ---------- تسجيل الدخول: يتحقق من الرقمين والحالة ويرجع الاسم فقط ----------
function login_(phone, parent) {
  var pk = key_(phone), qk = key_(parent);
  if (!pk || !qk) return { ok: false, error: 'not_found' };
  var me = students_().filter(function (s) { return s.phone === pk && s.parent === qk; })[0];
  if (!me) return { ok: false, error: 'not_found' };
  if (!me.active) return { ok: false, error: 'pending' };
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
  if (name.length < 3 || name.length > 80 || area.length < 2 || area.length > 80 || !okPhone(phone) || !okPhone(parent)) return { ok: false, error: 'bad data' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('الطلاب');
  if (!sh) return { ok: false, error: 'no sheet' };
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
        if (key_(ph[i][0]) === key_(phone)) return { ok: true }; // لا نكشف أن الرقم مسجل مسبقًا
      }
    }
    var row = n + 1, lc = sh.getLastColumn(), rng = sh.getRange(row, 1, 1, lc);
    rng.setNumberFormat('@');
    sh.getRange(row, cName).setValue(name);
    sh.getRange(row, cPhone).setValue(phone);
    sh.getRange(row, cParent).setValue(parent);
    sh.getRange(row, cArea).setValue(area);
    sh.getRange(row, cDate).setValue(Utilities.formatDate(new Date(), 'Asia/Qatar', 'yyyy-MM-dd HH:mm'));
    rng.setBackground('#fff2cc');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// عند كتابتك 2 في عمود "الحالة" يزول اللون الأصفر تلقائيًا
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== 'الطلاب' || e.range.getRow() < 2) return;
    var c = ensureCol_(sh, ['الحالة', 'status'], '');
    if (c && e.range.getColumn() === c && String(e.value).trim() === '2') sh.getRange(e.range.getRow(), 1, 1, sh.getLastColumn()).setBackground(null);
  } catch (err) {}
}

// ---------- للمعلم: قائمة 📊 نجم ← تحديث إحصاء الأسئلة (شيت "إحصاء الأسئلة") ----------
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('📊 نجم').addItem('تحديث إحصاء الأسئلة', 'buildQuestionStats').addToUi(); } catch (e) {}
}

function buildQuestionStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(RESULTS_SHEET);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('لا توجد نتائج بعد'); return; }
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues(), best = {};
  v.forEach(function (r) { // أفضل محاولة لكل طالب في كل اختبار
    var k = key_(r[2]) + '|' + r[4], pct = Number(r[8]);
    if (!best[k] || pct > best[k].pct) best[k] = { ex: r[4], title: r[5], pct: pct, wrong: parseWrong_(r[10]) };
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
