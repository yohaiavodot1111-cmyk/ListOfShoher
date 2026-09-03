/**
 * דו"ח אחד לשוחרים | מערכת דיווח נוכחות יומי
 * שכבת שרת (Google Apps Script) שמחברת את גיליון הגוגל שיטס לאתר.
 *
 * הפעלה ראשונה:
 *   1. פותחים את הגיליון, תפריט הרחבות, Apps Script.
 *   2. מדביקים את הקובץ הזה ואת index.html וגם appsscript.json.
 *   3. מריצים פעם אחת את הפונקציה setupSheets כדי ליצור את כל הגיליונות.
 *   4. פורסים כאפליקציית אינטרנט (Deploy, New deployment, Web app).
 */

/* ======================= קבועים ======================= */

var SH_STUDENTS  = 'תלמידים';
var SH_COMMANDS  = 'מפקדים';
var SH_REPORTS   = 'דיווחים';
var SH_SETTINGS  = 'הגדרות';

var DEFAULT_CUTOFF = '08:00';
var TZ = 'Asia/Jerusalem';

// חלוקת הסטטוסים כפי שסוכם
var STATUS_PRESENT = ['נוכח', 'פעילות חוץ', 'טיול/סיור'];
var STATUS_LATE    = ['איחור'];
var STATUS_ABSENT  = ['נעדר', 'מחלה', 'חופשה משפחתית', 'היעדרות ללא אישור', 'השעיה'];
var STATUS_ALL     = ['נוכח', 'איחור', 'נעדר', 'פעילות חוץ', 'טיול/סיור', 'מחלה', 'חופשה משפחתית', 'היעדרות ללא אישור', 'השעיה'];
var NO_REPORT      = 'לא דיווח';

// תפקידים
var ROLE_STUDENT   = 'תלמיד';
var ROLE_MAKAS     = 'מק"ס';
var ROLE_HELPER    = 'עוזר';
var ROLE_HEAD      = 'ראש צוות';
var ROLE_BRANCH    = 'מפקד שלוחה';

var TOKEN_TTL_MIN  = 720;   // תוקף חיבור בדקות
var LOCK_TRIES     = 6;     // ניסיונות כושלים עד נעילה
var LOCK_MIN       = 15;    // משך נעילה בדקות


/* ======================= תפריט והתקנה בהרצה אחת ======================= */

/**
 * נוסף אוטומטית לגיליון בכל פתיחה.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('דוח אחד')
    .addItem('התקנה מלאה', 'bootstrap')
    .addItem('הצגת כתובת האתר', 'showWebAppUrl')
    .addSeparator()
    .addItem('טעינת נתוני הדגמה', 'seedSampleStudents')
    .addToUi();
}

/**
 * הפונקציה היחידה שצריך להריץ. יוצרת את כל הגיליונות,
 * מזינה את סגל המפקדים, ומכינה את שעות הדיווח.
 */
function bootstrap() {
  setupSheets();
  seedCommanders_();
  syncSettingsClasses_();
  var msg = 'ההתקנה הושלמה. הגיליונות מוכנים וסגל המפקדים הוזן.\n\n' +
            'השלב הבא: הזנת התלמידים בגיליון תלמידים, ואז פריסה כאפליקציית אינטרנט.';
  try { SpreadsheetApp.getUi().alert('דוח אחד', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
  return msg;
}

function showWebAppUrl() {
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  var msg = url ? url : 'האתר טרם נפרס. יש ללחוץ Deploy ולבחור Web app.';
  try { SpreadsheetApp.getUi().alert('כתובת האתר', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
  return msg;
}

/** סגל המפקדים לפי המבנה המאושר. לא דורס נתונים קיימים. */
function seedCommanders_() {
  var sh = sheet_(SH_COMMANDS);
  if (sh.getLastRow() > 1) return;
  var rows = [
    ['יוחאי אליהו',   '', '', 'י2, י5',      ROLE_HEAD],
    ['יותם חכמון',    '', '', 'יב2, יג, יד', ROLE_HEAD],
    ['עילי ליבוביץ',  '', '', 'ט3',          ROLE_MAKAS],
    ['מאיה אשטמקר',   '', '', 'ט5',          ROLE_MAKAS],
    ['שילב חן',       '', '', 'יא2',         ROLE_MAKAS],
    ['נועם מאיר',     '', '', 'י2',          ROLE_HELPER],
    ['אביה אוחיון',   '', '', 'י5',          ROLE_HELPER],
    ['יסמין',         '', '', '',            ROLE_BRANCH]
  ];
  sh.getRange(2, 1, rows.length, 5).setValues(rows);
  sh.getRange(2, 3, rows.length, 1).setNote('יש להשלים מספר אישי. בלעדיו לא ניתן להתחבר.');
  sh.setColumnWidth(1, 150); sh.setColumnWidth(4, 150);
}

/** נתוני הדגמה לבדיקת המערכת. אופציונלי. */
function seedSampleStudents() {
  var sh = sheet_(SH_STUDENTS);
  if (!sh) { setupSheets(); sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_STUDENTS); }
  if (sh.getLastRow() > 1) {
    try { SpreadsheetApp.getUi().alert('גיליון התלמידים אינו ריק. לא בוצע שינוי.'); } catch (e) {}
    return;
  }
  var classes = ['ט3', 'ט5', 'י2', 'י5', 'יא2', 'יב2', 'יג', 'יד'];
  var first = ['איתי','נועה','דניאל','שירה','יונתן','מאיה','עומר','תמר','אריאל','רון'];
  var last  = ['כהן','לוי','מזרחי','פרץ','ביטון','אברהם','דהן','אזולאי','חדד','גבאי'];
  var rows = [], n = 100000;
  classes.forEach(function (cls) {
    for (var i = 0; i < 9; i++) {
      n++;
      rows.push([first[(n) % 10], last[(n * 3) % 10], String(300000000 + n), String(n), cls, '2008/05/14']);
    }
  });
  sh.getRange(2, 1, rows.length, 6).setValues(rows);
  syncSettingsClasses_();
  try { SpreadsheetApp.getUi().alert('נטענו ' + rows.length + ' תלמידי הדגמה.'); } catch (e) {}
}


/* ======================= הגשת האתר ======================= */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('דו"ח אחד לשוחרים')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* ======================= הקמת גיליונות ======================= */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SH_STUDENTS, ['שם פרטי', 'שם משפחה', 'תעודת זהות', 'מספר אישי', 'כיתה', 'תאריך לידה']);
  ensureSheet_(ss, SH_COMMANDS, ['שם מלא', 'דרגה', 'מספר אישי', 'כיתות באחריות', 'תפקיד']);
  ensureSheet_(ss, SH_REPORTS,  ['תאריך', 'מזהה', 'כיתה', 'סטטוס', 'סיבה', 'שעת דיווח', 'סטטוס אישור', 'מאשר', 'שעת אישור']);
  ensureSheet_(ss, SH_SETTINGS, ['כיתה', 'שעת דיווח אחרונה']);

  // מילוי שעת ברירת מחדל לכל כיתה קיימת
  syncSettingsClasses_();
  return 'הגיליונות מוכנים';
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// מוסיף לגיליון ההגדרות כל כיתה שקיימת בתלמידים ואין לה שורה
function syncSettingsClasses_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classes = distinctClasses_();
  var setSheet = sheet_(SH_SETTINGS);
  var existing = {};
  var data = setSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) existing[String(data[i][0]).trim()] = true;
  classes.forEach(function (c) {
    if (!existing[c]) setSheet.appendRow([c, DEFAULT_CUTOFF]);
  });
}


/* ======================= קריאת נתונים ======================= */

var SHEET_HEADERS = {};
SHEET_HEADERS[SH_STUDENTS] = ['שם פרטי', 'שם משפחה', 'תעודת זהות', 'מספר אישי', 'כיתה', 'תאריך לידה'];
SHEET_HEADERS[SH_COMMANDS] = ['שם מלא', 'דרגה', 'מספר אישי', 'כיתות באחריות', 'תפקיד'];
SHEET_HEADERS[SH_REPORTS]  = ['תאריך', 'מזהה', 'כיתה', 'סטטוס', 'סיבה', 'שעת דיווח', 'סטטוס אישור', 'מאשר', 'שעת אישור'];
SHEET_HEADERS[SH_SETTINGS] = ['כיתה', 'שעת דיווח אחרונה'];

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ensureSheet_(ss, name, SHEET_HEADERS[name] || ['עמודה']);
  return sh;
}

function sheetObjects_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { headers: values[0] || [], rows: [] };
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var obj = { _row: r + 1 };
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = values[r][c];
    rows.push(obj);
  }
  return { headers: headers, rows: rows };
}

function distinctClasses_() {
  var students = sheetObjects_(SH_STUDENTS).rows;
  var seen = {}, out = [];
  students.forEach(function (s) {
    var c = String(s['כיתה']).trim();
    if (c && !seen[c]) { seen[c] = true; out.push(c); }
  });
  return out;
}

function findStudent_(id) {
  id = String(id).trim();
  var rows = sheetObjects_(SH_STUDENTS).rows;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['תעודת זהות']).trim() === id || String(rows[i]['מספר אישי']).trim() === id) {
      return rows[i];
    }
  }
  return null;
}

function findCommander_(id) {
  id = String(id).trim();
  var rows = sheetObjects_(SH_COMMANDS).rows;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['מספר אישי']).trim() === id) return rows[i];
  }
  return null;
}

function commanderClasses_(commander) {
  var raw = String(commander['כיתות באחריות'] || '').trim();
  if (!raw) return [];
  return raw.split(/[,\u060C]/).map(function (x) { return x.trim(); }).filter(String);
}

function classRoster_(cls) {
  cls = String(cls).trim();
  return sheetObjects_(SH_STUDENTS).rows.filter(function (s) {
    return String(s['כיתה']).trim() === cls;
  });
}

function classCutoff_(cls) {
  cls = String(cls).trim();
  var rows = sheetObjects_(SH_SETTINGS).rows;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['כיתה']).trim() === cls) {
      return formatCutoff_(rows[i]['שעת דיווח אחרונה']) || DEFAULT_CUTOFF;
    }
  }
  return DEFAULT_CUTOFF;
}

function formatCutoff_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  var s = String(v).trim();
  return s;
}


/* ======================= זיהוי ואבטחה ======================= */

function scriptSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('SIGN_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SIGN_SECRET', s);
  }
  return s;
}

function signToken_(payloadObj) {
  var json = JSON.stringify(payloadObj);
  var body = Utilities.base64EncodeWebSafe(json);
  var sig = Utilities.computeHmacSha256Signature(body, scriptSecret_());
  var sigB = Utilities.base64EncodeWebSafe(sig);
  return body + '.' + sigB;
}

function verifyToken_(token) {
  if (!token || token.indexOf('.') < 0) throw new Error('נדרשת התחברות מחדש');
  var parts = token.split('.');
  var body = parts[0], sigB = parts[1];
  var expected = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(body, scriptSecret_()));
  if (sigB !== expected) throw new Error('נדרשת התחברות מחדש');
  var payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(body)).getDataAsString());
  if (payload.exp && Date.now() > payload.exp) throw new Error('פג תוקף החיבור, יש להתחבר שוב');
  return payload;
}

function lockKey_(id) { return 'lock_' + id; }
function triesKey_(id) { return 'tries_' + id; }

function login(idNumber) {
  var id = String(idNumber || '').trim();
  if (!id) return { ok: false, message: 'יש להזין מספר מזהה' };

  var cache = CacheService.getScriptCache();
  if (cache.get(lockKey_(id))) {
    return { ok: false, message: 'החשבון ננעל זמנית לאחר ניסיונות רבים. יש לנסות שוב בעוד מספר דקות' };
  }

  var commander = findCommander_(id);
  var student = commander ? null : findStudent_(id);

  if (!commander && !student) {
    var tries = parseInt(cache.get(triesKey_(id)) || '0', 10) + 1;
    if (tries >= LOCK_TRIES) {
      cache.put(lockKey_(id), '1', LOCK_MIN * 60);
      cache.remove(triesKey_(id));
    } else {
      cache.put(triesKey_(id), String(tries), LOCK_MIN * 60);
    }
    return { ok: false, message: 'מספר מזהה לא נמצא במערכת' };
  }
  cache.remove(triesKey_(id));

  var profile, payload;
  if (commander) {
    var role = String(commander['תפקיד'] || ROLE_MAKAS).trim();
    profile = {
      kind: 'commander',
      id: id,
      name: String(commander['שם מלא'] || '').trim(),
      rank: String(commander['דרגה'] || '').trim(),
      role: role,
      classes: commanderClasses_(commander),
      canGenerate: (role === ROLE_HEAD || role === ROLE_BRANCH),
      canSeeAll:   (role === ROLE_HEAD || role === ROLE_BRANCH),
      canEditAny:  (role === ROLE_BRANCH)
    };
    payload = { id: id, kind: 'commander', role: role, exp: Date.now() + TOKEN_TTL_MIN * 60000 };
  } else {
    profile = {
      kind: 'student',
      id: String(student['מספר אישי']).trim() || id,
      name: (String(student['שם פרטי']).trim() + ' ' + String(student['שם משפחה']).trim()).trim(),
      cls: String(student['כיתה']).trim()
    };
    payload = { id: profile.id, kind: 'student', exp: Date.now() + TOKEN_TTL_MIN * 60000 };
  }

  return { ok: true, token: signToken_(payload), profile: profile };
}


/* ======================= עזרי תאריך ======================= */

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd'); }

function weekStart_(dateStr) {
  // תחילת השבוע = יום ראשון
  var d = parseDate_(dateStr);
  var day = d.getDay(); // 0 ראשון
  d.setDate(d.getDate() - day);
  return Utilities.formatDate(d, TZ, 'yyyy/MM/dd');
}

function parseDate_(dateStr) {
  var p = String(dateStr).split('/');
  return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
}

function nowHM_() { return Utilities.formatDate(new Date(), TZ, 'HH:mm'); }

function isBirthday_(birthValue, dateStr) {
  if (!birthValue) return false;
  var b;
  if (birthValue instanceof Date) b = birthValue;
  else {
    var s = String(birthValue).trim().replace(/-/g, '/').replace(/\./g, '/');
    var p = s.split('/');
    if (p.length < 3) return false;
    // תמיכה גם בפורמט יום/חודש/שנה וגם שנה/חודש/יום
    if (p[0].length === 4) b = new Date(+p[0], +p[1] - 1, +p[2]);
    else b = new Date(+p[2], +p[1] - 1, +p[0]);
  }
  var t = parseDate_(dateStr);
  return b.getDate() === t.getDate() && b.getMonth() === t.getMonth();
}


/* ======================= דיווחים ======================= */

function reportRowIndex_(dateStr, id) {
  var sh = sheet_(SH_REPORTS);
  var values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    var d = values[r][0] instanceof Date ? Utilities.formatDate(values[r][0], TZ, 'yyyy/MM/dd') : String(values[r][0]).trim();
    if (d === dateStr && String(values[r][1]).trim() === String(id).trim()) return r + 1;
  }
  return -1;
}

function getReport_(dateStr, id) {
  var idx = reportRowIndex_(dateStr, id);
  if (idx < 0) return null;
  var sh = sheet_(SH_REPORTS);
  var v = sh.getRange(idx, 1, 1, 9).getValues()[0];
  return {
    _row: idx,
    date: dateStr,
    id: String(v[1]).trim(),
    cls: String(v[2]).trim(),
    status: String(v[3]).trim(),
    reason: String(v[4]).trim(),
    reportedAt: String(v[5]).trim(),
    approval: String(v[6]).trim(),
    approver: String(v[7]).trim(),
    approvedAt: String(v[8]).trim()
  };
}

function upsertReport_(rowObj) {
  var sh = sheet_(SH_REPORTS);
  var idx = reportRowIndex_(rowObj.date, rowObj.id);
  var row = [rowObj.date, rowObj.id, rowObj.cls, rowObj.status, rowObj.reason,
            rowObj.reportedAt, rowObj.approval, rowObj.approver, rowObj.approvedAt];
  if (idx < 0) sh.appendRow(row);
  else sh.getRange(idx, 1, 1, 9).setValues([row]);
}


/* ======================= נקודות קצה: תלמיד ======================= */

function studentGetDay(token, dateStr) {
  var s = verifyToken_(token);
  if (s.kind !== 'student') throw new Error('אין הרשאה');
  dateStr = dateStr || today_();
  var student = findStudent_(s.id);
  if (!student) throw new Error('התלמיד לא נמצא');
  var cls = String(student['כיתה']).trim();
  var rep = getReport_(dateStr, s.id);
  var cutoff = classCutoff_(cls);
  var locked = rep && (rep.approval === 'אושר' || rep.approval === 'שונה');

  // ימי הולדת בכיתה
  var mates = classRoster_(cls).filter(function (m) {
    return isBirthday_(m['תאריך לידה'], dateStr) &&
           String(m['מספר אישי']).trim() !== s.id;
  }).map(function (m) {
    return (String(m['שם פרטי']).trim() + ' ' + String(m['שם משפחה']).trim()).trim();
  });

  return {
    date: dateStr,
    cls: cls,
    statuses: STATUS_ALL,
    presentStatuses: STATUS_PRESENT,
    cutoff: cutoff,
    now: nowHM_(),
    afterCutoff: nowHM_() > cutoff,
    report: rep,
    locked: !!locked,
    birthdays: mates
  };
}

function studentSubmit(token, dateStr, status, reason) {
  var s = verifyToken_(token);
  if (s.kind !== 'student') throw new Error('אין הרשאה');
  dateStr = dateStr || today_();
  var student = findStudent_(s.id);
  if (!student) throw new Error('התלמיד לא נמצא');
  var cls = String(student['כיתה']).trim();

  if (STATUS_ALL.indexOf(status) < 0) throw new Error('סטטוס לא חוקי');
  var existing = getReport_(dateStr, s.id);
  if (existing && (existing.approval === 'אושר' || existing.approval === 'שונה')) {
    throw new Error('הדיווח כבר אושר על ידי המפקד ולא ניתן לשינוי');
  }

  var isPresent = STATUS_PRESENT.indexOf(status) >= 0;
  if (status !== 'נוכח' && !String(reason || '').trim()) {
    throw new Error('חובה לפרט סיבה עבור הסטטוס שנבחר');
  }
  // חסימת דיווח נוכחות אחרי שעת הגבול
  if (status === 'נוכח' && nowHM_() > classCutoff_(cls)) {
    throw new Error('חלף מועד דיווח הנוכחות לכיתה. יש לפנות למפקד');
  }

  upsertReport_({
    date: dateStr, id: s.id, cls: cls, status: status,
    reason: String(reason || '').trim(), reportedAt: nowHM_(),
    approval: 'ממתין', approver: '', approvedAt: ''
  });
  return { ok: true };
}


/* ======================= נקודות קצה: מפקד ======================= */

function commanderScope_(profile) {
  // אילו כיתות המפקד רשאי לצפות בהן
  if (profile.canSeeAll) return distinctClasses_();
  return profile.classes.slice();
}

function commanderApproveScope_(profile) {
  // אילו כיתות המפקד רשאי לאשר או לשנות בהן
  if (profile.role === ROLE_BRANCH) return distinctClasses_();
  return profile.classes.slice();
}

function commanderGetClasses(token, dateStr) {
  var s = verifyToken_(token);
  if (s.kind !== 'commander') throw new Error('אין הרשאה');
  dateStr = dateStr || today_();
  var commander = findCommander_(s.id);
  if (!commander) throw new Error('המפקד לא נמצא');

  var profile = login(s.id).profile;
  var viewClasses = commanderScope_(profile);
  var approveClasses = commanderApproveScope_(profile);

  var out = { date: dateStr, statuses: STATUS_ALL, classes: [], profile: profile };

  viewClasses.forEach(function (cls) {
    var roster = classRoster_(cls);
    var students = roster.map(function (st) {
      var id = String(st['מספר אישי']).trim();
      var rep = getReport_(dateStr, id);
      var effStatus = rep ? rep.status : NO_REPORT;
      return {
        id: id,
        name: (String(st['שם פרטי']).trim() + ' ' + String(st['שם משפחה']).trim()).trim(),
        status: effStatus,
        reason: rep ? rep.reason : '',
        approval: rep ? rep.approval : NO_REPORT,
        reportedAt: rep ? rep.reportedAt : '',
        birthday: isBirthday_(st['תאריך לידה'], dateStr),
        noReport: !rep
      };
    });
    var pending = students.filter(function (x) { return x.approval === 'ממתין' || x.noReport; }).length;
    out.classes.push({
      cls: cls,
      canApprove: approveClasses.indexOf(cls) >= 0,
      cutoff: classCutoff_(cls),
      students: students,
      pending: pending,
      complete: pending === 0
    });
  });
  return out;
}

function commanderSetStatus(token, dateStr, studentId, status, reason, mode) {
  var s = verifyToken_(token);
  if (s.kind !== 'commander') throw new Error('אין הרשאה');
  dateStr = dateStr || today_();
  var profile = login(s.id).profile;

  var student = findStudent_(studentId);
  if (!student) throw new Error('התלמיד לא נמצא');
  var cls = String(student['כיתה']).trim();
  if (commanderApproveScope_(profile).indexOf(cls) < 0) throw new Error('אין הרשאה לכיתה זו');

  var rep = getReport_(dateStr, studentId);

  if (mode === 'approve') {
    // אישור הדיווח כפי שהתלמיד מסר
    if (!rep) throw new Error('אין דיווח לאשר. יש לקבוע סטטוס');
    if (rep.status !== 'נוכח' && !String(rep.reason || '').trim()) {
      throw new Error('חסרה סיבה בדיווח. יש לקבוע סטטוס עם סיבה');
    }
    upsertReport_({
      date: dateStr, id: studentId, cls: cls, status: rep.status, reason: rep.reason,
      reportedAt: rep.reportedAt || nowHM_(), approval: 'אושר',
      approver: profile.name, approvedAt: nowHM_()
    });
    return { ok: true };
  }

  // קביעה או שינוי של סטטוס על ידי המפקד
  if (STATUS_ALL.indexOf(status) < 0) throw new Error('סטטוס לא חוקי');
  if (status !== 'נוכח' && !String(reason || '').trim()) {
    throw new Error('חובה לפרט סיבה עבור הסטטוס שנבחר');
  }
  upsertReport_({
    date: dateStr, id: studentId, cls: cls, status: status, reason: String(reason || '').trim(),
    reportedAt: rep ? (rep.reportedAt || nowHM_()) : nowHM_(),
    approval: 'שונה', approver: profile.name, approvedAt: nowHM_()
  });
  return { ok: true };
}

function commanderApproveClass(token, dateStr, cls) {
  var s = verifyToken_(token);
  if (s.kind !== 'commander') throw new Error('אין הרשאה');
  dateStr = dateStr || today_();
  var profile = login(s.id).profile;
  cls = String(cls).trim();
  if (commanderApproveScope_(profile).indexOf(cls) < 0) throw new Error('אין הרשאה לכיתה זו');

  var roster = classRoster_(cls);
  var missing = [];
  roster.forEach(function (st) {
    var id = String(st['מספר אישי']).trim();
    var rep = getReport_(dateStr, id);
    if (!rep) { missing.push((String(st['שם פרטי']).trim() + ' ' + String(st['שם משפחה']).trim()).trim()); return; }
    if (rep.approval === 'ממתין') {
      upsertReport_({
        date: dateStr, id: id, cls: cls, status: rep.status, reason: rep.reason,
        reportedAt: rep.reportedAt || nowHM_(), approval: 'אושר',
        approver: profile.name, approvedAt: nowHM_()
      });
    }
  });
  return { ok: true, missing: missing };
}

function setCutoff(token, cls, time) {
  var s = verifyToken_(token);
  if (s.kind !== 'commander') throw new Error('אין הרשאה');
  var profile = login(s.id).profile;
  cls = String(cls).trim();
  // עוזר אינו משנה שעת גבול. מק"ס אחראי, ראש צוות ומפקד שלוחה כן.
  if (profile.role === ROLE_HELPER) throw new Error('שינוי שעת הגבול מותר למק"ס האחראי בלבד');
  if (commanderApproveScope_(profile).indexOf(cls) < 0 && !profile.canSeeAll) throw new Error('אין הרשאה לכיתה זו');
  if (!/^\d{2}:\d{2}$/.test(String(time))) throw new Error('שעה לא תקינה');

  var sh = sheet_(SH_SETTINGS);
  var values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === cls) {
      sh.getRange(r + 1, 2).setValue(time);
      return { ok: true };
    }
  }
  sh.appendRow([cls, time]);
  return { ok: true };
}


/* ======================= דשבורד ודוח שלוחה ======================= */

function branchOverview(token, dateStr) {
  var s = verifyToken_(token);
  if (s.kind !== 'commander') throw new Error('אין הרשאה');
  var profile = login(s.id).profile;
  if (!profile.canSeeAll) throw new Error('אין הרשאה לצפייה בכלל השלוחה');
  dateStr = dateStr || today_();

  var classes = distinctClasses_();
  var rows = classes.map(function (cls) {
    var roster = classRoster_(cls);
    var matzevet = roster.length, present = 0, late = 0, absent = 0, pending = 0, noRep = 0;
    roster.forEach(function (st) {
      var rep = getReport_(dateStr, String(st['מספר אישי']).trim());
      var eff = rep ? rep.status : NO_REPORT;
      if (eff === NO_REPORT) { noRep++; absent++; }
      else if (STATUS_LATE.indexOf(eff) >= 0) { late++; present++; }
      else if (STATUS_ABSENT.indexOf(eff) >= 0) { absent++; }
      else { present++; }
      if (rep && rep.approval === 'ממתין') pending++;
      if (!rep) pending++;
    });
    return { cls: cls, matzevet: matzevet, present: present, late: late,
             absent: absent, noReport: noRep, pending: pending, complete: pending === 0 };
  });
  var allComplete = rows.every(function (x) { return x.complete; });
  return { date: dateStr, classes: rows, allComplete: allComplete };
}


/* ======================= בניית הדוח והפקה ======================= */

function weeklyCount_(id, dateStr, category) {
  // סופר כמה ימים בשבוע הנוכחי (מיום ראשון עד התאריך ועד בכלל) התלמיד היה בקטגוריה
  var start = parseDate_(weekStart_(dateStr));
  var end = parseDate_(dateStr);
  var count = 0;
  var d = new Date(start.getTime());
  while (d.getTime() <= end.getTime()) {
    var ds = Utilities.formatDate(d, TZ, 'yyyy/MM/dd');
    var rep = getReport_(ds, id);
    var eff = rep ? rep.status : NO_REPORT;
    if (category === 'late' && STATUS_LATE.indexOf(eff) >= 0) count++;
    if (category === 'absent' && (STATUS_ABSENT.indexOf(eff) >= 0 || eff === NO_REPORT)) count++;
    d.setDate(d.getDate() + 1);
  }
  return count;
}

function buildReportData_(dateStr) {
  var classes = distinctClasses_();
  var blocks = classes.map(function (cls) {
    var roster = classRoster_(cls);
    var matzevet = roster.length;
    var absentList = [], lateList = [], absentCount = 0;

    roster.forEach(function (st) {
      var id = String(st['מספר אישי']).trim();
      var name = (String(st['שם פרטי']).trim() + ' ' + String(st['שם משפחה']).trim()).trim();
      var rep = getReport_(dateStr, id);
      var eff = rep ? rep.status : NO_REPORT;
      var reason = rep ? rep.reason : '';

      if (STATUS_LATE.indexOf(eff) >= 0) {
        lateList.push({ name: name, reason: reason, days: weeklyCount_(id, dateStr, 'late') });
      } else if (STATUS_ABSENT.indexOf(eff) >= 0) {
        absentCount++;
        absentList.push({ name: name, reason: reason, days: weeklyCount_(id, dateStr, 'absent') });
      } else if (eff === NO_REPORT) {
        absentCount++;
        absentList.push({ name: name, reason: NO_REPORT, days: weeklyCount_(id, dateStr, 'absent') });
      }
    });

    return {
      cls: cls,
      matzevet: matzevet,
      presentCount: matzevet - absentCount,
      absentCount: absentCount,
      lateCount: lateList.length,
      absentList: absentList,
      lateList: lateList
    };
  });
  return { date: dateStr, blocks: blocks };
}

function generateReport(token, dateStr, force) {
  var s = verifyToken_(token);
  if (s.kind !== 'commander') throw new Error('אין הרשאה');
  var profile = login(s.id).profile;
  if (!profile.canGenerate) throw new Error('הפקת דוח מותרת לראש צוות ולמפקד שלוחה בלבד');
  dateStr = dateStr || today_();

  var overview = branchOverview(token, dateStr);
  if (!overview.allComplete && !force) {
    return { needConfirm: true,
             message: 'היום לא כל המפקדים אישרו דיווח נוכחות. האם להפיק דוח עם הנתונים הנוכחיים?' };
  }

  var data = buildReportData_(dateStr);
  var files = exportReportFiles_(data);
  return { ok: true, preview: data, pdf: files.pdf, xlsx: files.xlsx, fileName: files.fileName };
}

function exportReportFiles_(data) {
  var dateLabel = data.date.replace(/\//g, '.');
  var tmp = SpreadsheetApp.create('דוח אחד ' + dateLabel);
  var id = tmp.getId();
  try {
    var absSheet = tmp.getActiveSheet();
    absSheet.setName('נעדרים');
    writeReportSheet_(absSheet, data, 'absent', dateLabel);
    var lateSheet = tmp.insertSheet('מאחרים');
    writeReportSheet_(lateSheet, data, 'late', dateLabel);
    SpreadsheetApp.flush();

    var pdf = exportBlob_(id, 'pdf');
    var xlsx = exportBlob_(id, 'xlsx');
    return {
      fileName: 'דוח אחד ' + dateLabel,
      pdf:  { name: 'דוח אחד ' + dateLabel + '.pdf',  data: Utilities.base64Encode(pdf.getBytes()) },
      xlsx: { name: 'דוח אחד ' + dateLabel + '.xlsx', data: Utilities.base64Encode(xlsx.getBytes()) }
    };
  } finally {
    DriveApp.getFileById(id).setTrashed(true);
  }
}

function writeReportSheet_(sheet, data, kind, dateLabel) {
  sheet.setRightToLeft(true);
  var title = (kind === 'absent') ? 'דוח נעדרים יומי' : 'דוח מאחרים יומי';
  var namesHdr = (kind === 'absent') ? 'שמות נעדרים' : 'שמות מאחרים';
  var reasonHdr = (kind === 'absent') ? 'סיבת ההיעדרות' : 'סיבת האיחור';
  var countHdr = (kind === 'absent') ? 'נעדרים' : 'מאחרים';

  // סדר עמודות מימין לשמאל: כיתה, מצבת חניכים, חניכים נוכחים, נעדרים/מאחרים, שמות, סיבה, סה"כ ימים
  var headers = ['כיתה', 'מצבת חניכים', 'חניכים נוכחים', countHdr, namesHdr, reasonHdr, 'סה"כ ימים'];
  var rows = [[title, '', '', '', '', '', ''], headers];

  data.blocks.forEach(function (b) {
    var list = (kind === 'absent') ? b.absentList : b.lateList;
    var cnt = (kind === 'absent') ? b.absentCount : b.lateCount;
    var maxRows = Math.max(6, list.length);
    for (var i = 0; i < maxRows; i++) {
      var item = list[i];
      rows.push([
        i === 0 ? b.cls : '',
        i === 0 ? b.matzevet : '',
        i === 0 ? b.presentCount : '',
        i === 0 ? cnt : '',
        item ? (i + 1) + '. ' + item.name : (i + 1) + '.',
        item ? item.reason : '',
        item ? item.days : ''
      ]);
    }
  });

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).merge()
    .setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold');
  sheet.getRange(2, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#0B2240').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  sheet.getRange(1, 1, rows.length, headers.length)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 260);
  for (var c = 1; c <= headers.length; c++) if (c !== 5 && c !== 6) sheet.setColumnWidth(c, 90);

  // מיזוג תאי הכיתה על פני שש שורות בכל בלוק
  var startRow = 3;
  data.blocks.forEach(function (b) {
    var list = (kind === 'absent') ? b.absentList : b.lateList;
    var maxRows = Math.max(6, list.length);
    for (var col = 1; col <= 4; col++) {
      sheet.getRange(startRow, col, maxRows, 1).merge()
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
    }
    startRow += maxRows;
  });
}

function exportBlob_(spreadsheetId, format) {
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId +
            '/export?format=' + format +
            '&size=A4&portrait=true&fitw=true&gridlines=false&sheetnames=false&printtitle=false';
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  return resp.getBlob();
}
