'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const ExcelJS = require('exceljs');

const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const XLSX_PATH = path.join(DATA_DIR, 'updates.xlsx');   // ملف Excel المشترك
const TZ = process.env.TZ_NAME || 'Asia/Baghdad';         // المنطقة الزمنية للتاريخ والوقت
const AUTH_USER = process.env.APP_USER, AUTH_PASS = process.env.APP_PASS; // اختياري: حماية بكلمة مرور

const TYPES = ['يدوي', 'ثابت', 'محمول بعجلة', 'معيدة'];
const TYPE_FILL = { 'يدوي': 'FFFFE2C2', 'ثابت': 'FFCDF0DB', 'محمول بعجلة': 'FFDDD6FF', 'معيدة': 'FFFFD3E0' };
const HEAD = ['م', 'الرقم التسلسلي للجهاز', 'اسم المبرمج', 'نوع الجهاز', 'اسم الوحدة', 'موقع الوحدة', 'التاريخ', 'الوقت', 'المعرف'];

let records = [];
let version = Date.now();
let saveChain = Promise.resolve();

const httpErr = (status, message) => Object.assign(new Error(message), { status });

/* ---------- Excel: قراءة ---------- */
async function loadFromFile() {
  if (!fs.existsSync(XLSX_PATH)) return;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  const ws = wb.worksheets[0];
  if (!ws) return;
  const off = ws.getRow(1).getCell(1).text.trim() === 'م' ? 1 : 0;
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const g = i => String(row.getCell(i + off).text || '').trim();
    if (!g(1) && !g(2)) return;
    const id = off ? String(row.getCell(9).text || '').trim() : '';
    records.push({ id: id || crypto.randomUUID(), serial: g(1), programmer: g(2), type: g(3), unit: g(4), location: g(5), date: g(6), time: g(7) });
  });
}

/* ---------- Excel: كتابة (بالتتابع وبشكل آمن) ---------- */
async function doSave() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('التحديثات', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }] });
  ws.columns = [6, 24, 22, 16, 24, 28, 14, 14, 10].map(w => ({ width: w }));
  ws.getColumn(9).hidden = true; // عمود المعرف مخفي
  const line = { style: 'thin', color: { argb: 'FFB8C4DC' } };
  const border = { top: line, left: line, bottom: line, right: line };
  const center = { horizontal: 'center', vertical: 'middle', readingOrder: 'rtl' };

  const h = ws.addRow(HEAD); h.height = 30;
  h.eachCell(c => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A1530' } };
    c.font = { bold: true, size: 13, color: { argb: 'FFD4AF37' }, name: 'Arial' };
    c.alignment = center;
    c.border = { top: line, left: line, right: line, bottom: { style: 'medium', color: { argb: 'FFD4AF37' } } };
  });
  records.forEach((r, i) => {
    const row = ws.addRow([i + 1, r.serial, r.programmer, r.type, r.unit, r.location, r.date, r.time, r.id]);
    row.height = 24;
    const zebra = i % 2 ? 'FFE6EDFA' : 'FFFFFFFF';
    row.eachCell((c, col) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: col === 4 ? (TYPE_FILL[r.type] || zebra) : (col === 1 ? 'FFF3E3A6' : zebra) } };
      c.font = { size: 12, name: 'Arial', bold: col === 1 || col === 4, color: { argb: 'FF0A1530' } };
      c.alignment = center; c.border = border;
    });
  });
  ws.autoFilter = { from: 'A1', to: 'H1' };
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const tmp = XLSX_PATH + '.tmp';
  await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, XLSX_PATH);
}
function saveToFile() {
  const p = saveChain.catch(() => {}).then(doSave);
  saveChain = p;
  return p;
}

/* ---------- أدوات ---------- */
function stamp() {
  const p = new Intl.DateTimeFormat('en-US', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: 'numeric', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(new Date()).reduce((a, x) => (a[x.type] = x.value, a), {});
  const h = +p.hour % 24;
  return { date: `${p.year}-${p.month}-${p.day}`, time: `${h % 12 || 12}:${p.minute}:${p.second} ${h >= 12 ? 'م' : 'ص'}` };
}
function clean(b) {
  const o = {};
  for (const k of ['serial', 'programmer', 'unit', 'location']) {
    o[k] = String(b[k] ?? '').trim().slice(0, 200);
    if (!o[k]) throw httpErr(400, 'أكمل جميع الحقول');
  }
  o.type = String(b.type ?? '');
  if (!TYPES.includes(o.type)) throw httpErr(400, 'نوع الجهاز غير صحيح');
  return o;
}
function readBody(req) {
  return new Promise((ok, no) => {
    let d = '';
    req.on('data', c => { d += c; if (d.length > 1e5) { req.destroy(); no(httpErr(413, 'الطلب كبير جداً')); } });
    req.on('end', () => { try { ok(JSON.parse(d || '{}')); } catch { no(httpErr(400, 'بيانات غير صالحة')); } });
    req.on('error', no);
  });
}
function send(res, status, obj) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(obj));
}
function authorized(req) {
  const h = req.headers.authorization || '';
  if (!h.startsWith('Basic ')) return false;
  const a = Buffer.from(Buffer.from(h.slice(6), 'base64').toString());
  const b = Buffer.from(`${AUTH_USER}:${AUTH_PASS}`);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
const snapshot = () => ({ version, records });

/* ---------- الخادم ---------- */
const server = http.createServer(async (req, res) => {
  try {
    if (AUTH_USER && !authorized(req)) {
      res.writeHead(401, { 'WWW-Authenticate': 'Basic realm="Signals", charset="UTF-8"' });
      return res.end('Unauthorized');
    }
    const p = new URL(req.url, 'http://x').pathname;

    if (req.method === 'GET' && (p === '/' || p === '/index.html')) {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' });
      return fs.createReadStream(path.join(__dirname, 'public', 'index.html')).pipe(res);
    }
    if (req.method === 'GET' && p === '/api/download') {
      await saveChain.catch(() => {});
      if (!fs.existsSync(XLSX_PATH)) throw httpErr(404, 'لا يوجد ملف بعد');
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="updates.xlsx"; filename*=UTF-8''${encodeURIComponent('سجل_تحديث_الأجهزة.xlsx')}`,
        'Cache-Control': 'no-store'
      });
      return fs.createReadStream(XLSX_PATH).pipe(res);
    }
    if (p === '/api/records' && req.method === 'GET') return send(res, 200, snapshot());
    if (p === '/api/records' && req.method === 'POST') {
      const rec = { id: crypto.randomUUID(), ...clean(await readBody(req)), ...stamp() };
      records.push(rec); version++;
      await saveToFile();
      return send(res, 201, { ...snapshot(), id: rec.id });
    }
    const m = p.match(/^\/api\/records\/([\w-]+)$/);
    if (m && (req.method === 'PUT' || req.method === 'DELETE')) {
      const i = records.findIndex(r => r.id === m[1]);
      if (i < 0) throw httpErr(404, 'السجل غير موجود (ربما حُذف من جهاز آخر)');
      if (req.method === 'PUT') records[i] = { ...records[i], ...clean(await readBody(req)) };
      else records.splice(i, 1);
      version++;
      await saveToFile();
      return send(res, 200, snapshot());
    }
    send(res, 404, { error: 'غير موجود' });
  } catch (e) {
    if (!e.status) console.error(e);
    send(res, e.status || 500, { error: e.status ? e.message : 'خطأ في الخادم' });
  }
});

loadFromFile()
  .catch(e => console.error('تعذرت قراءة ملف Excel الحالي:', e.message))
  .then(() => server.listen(PORT, () => console.log(`الخادم يعمل على المنفذ ${PORT} — عدد السجلات: ${records.length}`)));
