/** @OnlyCurrentDoc */

/**
 * הזמנת חולצות צוותיות - Google Apps Script
 * הסקריפט צמוד לגיליון גוגל ומשמש כ"שרת" של האתר שב-GitHub Pages:
 *   GET  -> מחזיר לאתר את הפרינטים, החולצות, המחירים וההגדרות מהגיליון
 *   POST -> בודק הזמנה שנשלחה מהאתר וכותב שורה לכל חולצה בלשונית "הזמנות"
 * הפרינטים, סוגי החולצות, המחירים והקישור לפייבוקס נערכים בלשוניות הגיליון, בלי לגעת בקוד.
 * פריסה: Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 */

const TABS = {
  orders: 'הזמנות',
  summary: 'סיכום',
  prints: 'פרינטים',
  images: 'תמונות',
  shirts: 'חולצות',
  settings: 'הגדרות'
};

const ORDER_HEADERS = ['תאריך', 'מספר הזמנה', 'שם מלא', 'טלפון', 'מייל', 'פרינט', 'סוג חולצה', 'מידה', 'כמות', 'מחיר ליחידה', 'סהכ לפריט', 'אישור תשלום', 'מזהה שליחה'];
const DEFAULT_PAYBOX = 'https://links.payboxapp.com/oFRZxusD67b';
const DEFAULT_TITLE = 'חנות המרצ׳ של המסייעת';
const DEFAULT_INTRO = 'בוחרים פרינט, סוג חולצה ומידה, משלמים בפייבוקס ושולחים. אפשר להזמין כמה חולצות בהזמנה אחת.';
const DEFAULT_CLOSED = 'ההזמנות נסגרו. תודה לכל מי שהזמין.';
const IMMODEST_NOTE = 'שימו לב: הפרינט הזה אינו צנוע ולכן מוצג מטושטש. לחיצה על התמונה תציג אותו בתצוגה מקדימה.';
const SAMPLE_SIZES = 'S, M, L, XL, XXL';
const MAX_ITEMS = 30;
const MAX_QTY = 20;

/** תפריט בגיליון */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('חולצות')
    .addItem('הקמת הלשוניות', 'setup')
    .addToUi();
}

/** יוצר את הלשוניות החסרות. לא נוגע בלשוניות קיימות. */
function setup() {
  ensureTabs_(SpreadsheetApp.getActiveSpreadsheet());
}

/** האתר מבקש את ההגדרות */
function doGet() {
  return json_(safe_(function () {
    const cfg = getConfig();
    cfg.ok = true;
    return cfg;
  }));
}

/** האתר שולח הזמנה (JSON בגוף הבקשה) */
function doPost(e) {
  return json_(safe_(function () {
    const res = submitOrder(JSON.parse(e.postData.contents));
    res.ok = true;
    return res;
  }));
}

function safe_(fn) {
  try {
    return fn();
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** מחזיר פרינטים, חולצות והגדרות מהגיליון */
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureTabs_(ss);

  const settings = {};
  rows_(ss, TABS.settings).forEach(function (r) {
    const key = String(r[0] || '').trim();
    if (key) settings[key] = r[1];
  });

  // תמונות לפי פרינט וסוג חולצה (מלפנים ומאחור), ומחיר מיוחד לצירוף אם הוגדר
  const variants = {};
  rows_(ss, TABS.images).forEach(function (r) {
    const print = String(r[0] || '').trim();
    const images = [imageUrl_(r[2]), imageUrl_(r[3])].filter(String);
    const price = Number(r[4]);
    const v = { shirt: String(r[1] || '').trim(), images: images };
    if (price > 0) v.price = price;
    if (!print || !(images.length || v.price)) return;
    (variants[print] = variants[print] || []).push(v);
  });

  const prints = rows_(ss, TABS.prints)
    .filter(function (r) { return String(r[0] || '').trim() && yes_(r[4], true); })
    .map(function (r) {
      const name = String(r[0]).trim();
      return {
        name: name,
        image: imageUrl_(r[1]),
        blurred: yes_(r[2], false),
        note: String(r[3] || '').trim(),
        cover: /אחור|back/i.test(String(r[5] || '')) ? 'back' : 'front',
        variants: variants[name] || []
      };
    });

  const shirts = rows_(ss, TABS.shirts)
    .filter(function (r) { return String(r[0] || '').trim() && yes_(r[3], true); })
    .map(function (r) {
      return {
        name: String(r[0]).trim(),
        price: Number(r[1]) || 0,
        sizes: String(r[2] || '').split(/[,\n]+/).map(function (s) { return s.trim(); }).filter(String)
      };
    })
    .filter(function (s) { return s.sizes.length; });

  return {
    title: String(settings['כותרת'] || DEFAULT_TITLE),
    intro: String(settings['טקסט פתיחה'] || DEFAULT_INTRO),
    payboxUrl: payboxUrl_(settings['קישור פייבוקס']),
    open: yes_(settings['פתוח להזמנות'], true),
    closedText: String(settings['הודעת סגירה'] || DEFAULT_CLOSED),
    prints: prints,
    shirts: shirts
  };
}

/** בודק את ההזמנה וכותב שורה לכל פריט */
function submitOrder(order) {
  const cfg = getConfig();
  if (!cfg.open) throw new Error('ההזמנות סגורות כרגע');
  if (!order || typeof order !== 'object') throw new Error('ההזמנה ריקה');

  const name = clean_(order.name, 80);
  const phone = clean_(order.phone, 20).replace(/[^\d+]/g, '');
  const email = clean_(order.email, 120);
  if (name.length < 2) throw new Error('חסר שם מלא');
  if (!/^0\d{8,9}$/.test(phone) && !/^\+?972\d{8,9}$/.test(phone)) throw new Error('מספר הטלפון לא תקין');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('כתובת המייל לא תקינה');
  if (order.paid !== true) throw new Error('צריך לאשר שבוצע תשלום');
  const submitId = clean_(order.submitId, 40).replace(/[^\w-]/g, '');

  const items = Array.isArray(order.items) ? order.items : [];
  if (!items.length) throw new Error('לא נבחרו חולצות');
  if (items.length > MAX_ITEMS) throw new Error('יותר מדי פריטים בהזמנה אחת');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const orderId = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'MMdd-HHmmss');
  let total = 0;

  const rows = items.map(function (it, i) {
    const print = cfg.prints.filter(function (p) { return p.name === it.print; })[0];
    const shirt = cfg.shirts.filter(function (s) { return s.name === it.shirt; })[0];
    const n = i + 1;
    if (!print) throw new Error('בחולצה ' + n + ' הפרינט כבר לא זמין, רעננו את הדף');
    if (!shirt) throw new Error('בחולצה ' + n + ' סוג החולצה כבר לא זמין, רעננו את הדף');
    if (shirt.sizes.indexOf(it.size) < 0) throw new Error('בחולצה ' + n + ' המידה לא זמינה');
    const qty = Math.floor(Number(it.qty));
    if (!(qty >= 1 && qty <= MAX_QTY)) throw new Error('בחולצה ' + n + ' הכמות לא תקינה');
    const price = priceOf_(print, shirt);
    const line = price * qty;
    total += line;
    return [now, orderId, cellText_(name), cellText_(phone), cellText_(email), print.name, shirt.name, it.size, qty, price, line, 'כן', submitId];
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(TABS.orders);
    const last = sh.getLastRow();

    // אותה שליחה הגיעה פעמיים (למשל ניסיון חוזר אחרי תקלת רשת): לא כותבים שוב
    if (submitId && last > 1) {
      const from = Math.max(2, last - 499);
      const recent = sh.getRange(from, 1, last - from + 1, ORDER_HEADERS.length).getValues()
        .filter(function (r) { return String(r[12]) === submitId; });
      if (recent.length) {
        return {
          orderId: String(recent[0][1]),
          total: recent.reduce(function (sum, r) { return sum + (Number(r[10]) || 0); }, 0)
        };
      }
    }

    sh.getRange(last + 1, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return { orderId: orderId, total: total };
}

/* ---------------- עזר ---------------- */

/** טקסט שהמזמין הקליד לא ייכתב לגיליון כנוסחה (למשל שם שמתחיל ב-=) */
function cellText_(s) {
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/** רק קישור של פייבוקס יוצג באתר ככפתור תשלום */
function payboxUrl_(v) {
  const s = String(v || '').trim();
  return /^https:\/\/([\w-]+\.)*payboxapp\.com\/[^\s"'<>]*$/i.test(s) ? s : DEFAULT_PAYBOX;
}

/** מחיר לצירוף של פרינט וסוג חולצה: מחיר מיוחד מלשונית "תמונות" אם הוגדר, אחרת המחיר של סוג החולצה */
function priceOf_(print, shirt) {
  const v = (print.variants || []).filter(function (x) { return x.shirt === shirt.name && x.price > 0; })[0];
  return v ? v.price : shirt.price;
}

function ensureTabs_(ss) {
  tab_(ss, TABS.orders, ORDER_HEADERS, [], function (sh) {
    sh.getRange('C:E').setNumberFormat('@');          // שם, טלפון ומייל כטקסט (שומר על 0 בתחילת הטלפון)
    sh.getRange('M:M').setNumberFormat('@');
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  });

  tab_(ss, TABS.summary, [], [], function (sh) {
    // מספרים נכפים כמספר (*1) כדי ש-QUERY לא יזהה את העמודה כטקסט כשאין עדיין הזמנות
    sh.getRange('A1').setFormula(
      `=IFERROR(QUERY({'${TABS.orders}'!F2:H, ARRAYFORMULA('${TABS.orders}'!I2:I*1)}, "select Col1, Col2, Col3, sum(Col4) where Col1 <> '' group by Col1, Col2, Col3 label Col1 'פרינט', Col2 'סוג חולצה', Col3 'מידה', sum(Col4) 'כמות'", 0), "אין הזמנות עדיין")`);
    sh.getRange('F1').setFormula(
      `=IFERROR(QUERY({'${TABS.orders}'!B2:D, ARRAYFORMULA('${TABS.orders}'!K2:K*1)}, "select Col1, Col2, Col3, sum(Col4) where Col2 <> '' group by Col1, Col2, Col3 label Col1 'מספר הזמנה', Col2 'שם מלא', Col3 'טלפון', sum(Col4) 'לתשלום'", 0), "")`);
  });

  tab_(ss, TABS.prints, ['שם הפרינט', 'קישור לתמונה', 'מטושטש', 'הערה', 'פעיל', 'תמונה ראשית'], [
    ['לוגו מסייעת', '', 'לא', '', 'כן', 'מלפנים'],
    ['שלחי ניוד', '', 'כן', IMMODEST_NOTE, 'כן', 'מאחור'],
    ['777 בום', '', 'לא', '', 'כן', 'מאחור']
  ], function (sh) {
    sh.getRange('F1').setNote('מלפנים או מאחור: איזה צד של החולצה מוצג בכרטיס ובתמונה הקטנה.');
    sh.getRange('B1').setNote('תמונה כללית לפרינט. לא חובה אם יש תמונות בלשונית "תמונות".');
    sh.getRange('C1').setNote('כן = התמונה מוצגת מטושטשת ונפתחת רק בלחיצה. הערה מומלצת: ' + IMMODEST_NOTE);
  });

  tab_(ss, TABS.images, ['פרינט', 'סוג חולצה', 'תמונה מלפנים', 'תמונה מאחור', 'מחיר מיוחד'], [
    ['לוגו מסייעת', 'אוברסייז', 'mesayaat-oversize-front', 'mesayaat-oversize-back', ''],
    ['לוגו מסייעת', 'טי רגילה', 'mesayaat-tee-front', 'mesayaat-tee-back', ''],
    ['שלחי ניוד', 'אוברסייז', 'mesayaat-oversize-front', 'shalhi-oversize-back', 100],
    ['שלחי ניוד', 'טי רגילה', 'mesayaat-tee-front', 'shalhi-tee-back', 60],
    ['777 בום', 'אוברסייז', 'mesayaat-oversize-front', 'boom-oversize-back', 100],
    ['777 בום', 'טי רגילה', 'mesayaat-tee-front', 'boom-tee-back', 60]
  ], function (sh) {
    sh.getRange('E1').setNote('מחיר לצירוף הזה של פרינט וסוג חולצה. ריק = המחיר מלשונית "חולצות".');
    sh.getRange('C1').setNote('שם של תמונה שנמצאת בתיקיית img באתר (בלי הסיומת), קישור שיתוף מגוגל דרייב עם גישה לכל מי שיש לו את הקישור, או כתובת תמונה.');
  });

  tab_(ss, TABS.shirts, ['סוג חולצה', 'מחיר', 'מידות', 'פעיל'], [
    ['אוברסייז', 70, SAMPLE_SIZES, 'כן'],
    ['טי רגילה', 50, SAMPLE_SIZES, 'כן']
  ]);

  tab_(ss, TABS.settings, ['הגדרה', 'ערך'], [
    ['כותרת', DEFAULT_TITLE],
    ['טקסט פתיחה', DEFAULT_INTRO],
    ['קישור פייבוקס', DEFAULT_PAYBOX],
    ['פתוח להזמנות', 'כן'],
    ['הודעת סגירה', DEFAULT_CLOSED]
  ]);
}

function tab_(ss, name, headers, rows, after) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.setRightToLeft(true);
  if (headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  if (after) after(sh);
  return sh;
}

function rows_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 1)).getValues();
}

function yes_(v, dflt) {
  if (v === true || v === false) return v;
  const s = String(v == null ? '' : v).trim().toLowerCase();
  if (!s) return dflt;
  return ['כן', 'yes', 'y', 'true', '1', 'v', 'x'].indexOf(s) >= 0;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
}

/** מקבל קישור שיתוף של גוגל דרייב, מזהה קובץ או כתובת תמונה רגילה, ומחזיר כתובת שאפשר להציג באתר */
function imageUrl_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/\/d\/([\w-]{20,})/) || s.match(/[?&]id=([\w-]{20,})/);
  if (m) return 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w1200';
  if (/^[\w-]{25,}$/.test(s)) return 'https://drive.google.com/thumbnail?id=' + s + '&sz=w1200';
  if (/^https:\/\//i.test(s)) return s;
  if (/^img\/[\w.\/-]{1,80}$/.test(s) && s.indexOf('..') < 0) return s;   // קובץ בתיקיית img של האתר
  return /^[\w-]{1,24}$/.test(s) ? s : '';   // שם של תמונה שמוגדרת ב-index.html
}
