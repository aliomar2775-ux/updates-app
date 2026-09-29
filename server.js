'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

/* ---------- تحميل ملف .env (بدون أي مكتبة إضافية) ---------- */
(function loadEnv() {
  try {
    const f = path.join(__dirname, '.env');
    if (!fs.existsSync(f)) return;
    for (const line of fs.readFileSync(f, 'utf8').split(/\r?\n/)) {
      if (line.trim().startsWith('#')) continue;
      const m = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
      if (!m) continue;
      let v = m[2];
      if (/^(['"]).*\1$/.test(v)) v = v.slice(1, -1);
      if (!(m[1] in process.env)) process.env[m[1]] = v; // متغيرات السيرفر لها الأولوية على .env
    }
  } catch (e) { console.error('تعذرت قراءة .env:', e.message); }
})();

const ExcelJS = require('exceljs');

const PORT = process.env.PORT || 3000;
const DATA_DIR = path.resolve(__dirname, process.env.DATA_DIR || 'data');
const XLSX_PATH = path.join(DATA_DIR, 'updates.xlsx');   // ملف Excel المشترك
const DB_PATH = path.join(DATA_DIR, 'db.json');          // قاعدة البيانات الداخلية (المصدر الأساسي)
const TZ = process.env.TZ_NAME || 'Asia/Baghdad';
const AUTH_USER = process.env.APP_USER, AUTH_PASS = process.env.APP_PASS;
const AUTH_ON = !!(AUTH_USER && AUTH_PASS);

/* مزامنة اختيارية مع GitHub لحفظ البيانات على السيرفرات المجانية */
const GH = {
  token: process.env.GITHUB_TOKEN,
  repo: process.env.GITHUB_REPO,                          // owner/repo
  branch: process.env.GITHUB_BRANCH || 'data',
  file: process.env.GITHUB_PATH || 'data/db.json'
};
GH.enabled = !!(GH.token && GH.repo);

const TYPES = ['يدوي', 'ثابت', 'محمول بعجلة', 'معيدة'];
const TYPE_FILL = { 'يدوي': 'FFFFE2C2', 'ثابت': 'FFCDF0DB', 'محمول بعجلة': 'FFDDD6FF', 'معيدة': 'FFFFD3E0' };
const HEAD = ['م', 'الرقم التسلسلي للجهاز', 'اسم المبرمج', 'نوع الجهاز', 'اسم الوحدة', 'موقع الوحدة', 'التاريخ', 'الوقت', 'المعرف'];

let records = [];
let version = Date.now();
let saveChain = Promise.resolve();

const httpErr = (status, message) => Object.assign(new Error(message), { status });
const serialize = () => JSON.stringify({ format: 1, savedAt: new Date().toISOString(), records }, null, 2);

/* ---------- شارة الإعداد (تُحقن في الصفحة تلقائياً) ---------- */
const CREDIT = `
<style>
.sig-credit{position:fixed;left:18px;bottom:18px;z-index:9999;direction:rtl;pointer-events:none;
  font-family:'Cairo','Tajawal','Segoe UI',Tahoma,sans-serif;display:flex;align-items:center;gap:13px;
  padding:11px 20px 11px 16px;border-radius:18px;overflow:hidden;color:#f6efd6;
  background:linear-gradient(135deg,rgba(8,17,40,.97),rgba(22,42,90,.95));
  border:1px solid rgba(212,175,55,.7);
  box-shadow:0 14px 34px rgba(0,0,0,.38),0 0 0 4px rgba(212,175,55,.08),inset 0 1px 0 rgba(255,255,255,.08);
  backdrop-filter:blur(8px);animation:sigIn .9s cubic-bezier(.2,.8,.2,1) both}
.sig-credit::before{content:'';position:absolute;inset:0;
  background:linear-gradient(110deg,transparent 30%,rgba(212,175,55,.25) 50%,transparent 70%);
  transform:translateX(-130%);animation:sigShine 6s ease-in-out 1.5s infinite}
.sig-credit::after{content:'';position:absolute;right:0;top:12%;bottom:12%;width:3px;border-radius:3px;
  background:linear-gradient(#f3e3a6,#d4af37,#8a6d17)}
.sig-emblem{flex:none;width:44px;height:44px;border-radius:50%;display:grid;place-items:center;
  background:radial-gradient(circle at 30% 25%,#f8ecb8,#d4af37 55%,#8a6d17);
  box-shadow:0 0 0 3px rgba(212,175,55,.25),0 4px 12px rgba(0,0,0,.4)}
.sig-emblem svg{width:24px;height:24px}
.sig-txt{display:flex;flex-direction:column;line-height:1.35;padding-right:6px}
.sig-lbl{font-size:11px;letter-spacing:.14em;color:#d4af37;font-weight:700}
.sig-rank{font-size:13px;color:#e9e2c8;font-weight:600}
.sig-name{font-size:17px;font-weight:800;
  background:linear-gradient(90deg,#fff4c7,#d4af37 55%,#f3e3a6);-webkit-background-clip:text;background-clip:text;
  -webkit-text-fill-color:transparent;color:#d4af37}
@keyframes sigShine{60%,100%{transform:translateX(130%)}}
@keyframes sigIn{from{opacity:0;transform:translateY(24px) scale(.96)}to{opacity:1;transform:none}}
@media (max-width:560px){.sig-credit{left:10px;bottom:10px;padding:8px 14px 8px 12px;gap:9px}
  .sig-emblem{width:34px;height:34px}.sig-emblem svg{width:18px;height:18px}
  .sig-name{font-size:14px}.sig-rank{font-size:11px}.sig-lbl{font-size:9px}}
@media print{.sig-credit{display:none}}
</style>
<div class="sig-credit" aria-label="إعداد">
  <div class="sig-emblem"><svg viewBox="0 0 24 24" fill="none" stroke="#0a1530" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5a10 10 0 0 1 14 0"/><path d="M8.2 15.7a5.5 5.5 0 0 1 7.6 0"/><circle cx="12" cy="19" r="1.4" fill="#0a1530"/><path d="M12 3l1.1 2.3 2.5.3-1.8 1.7.5 2.5L12 8.6 9.7 9.8l.5-2.5L8.4 5.6l2.5-.3z" fill="#0a1530" stroke="none"/></svg></div>
  <div class="sig-txt">
    <span class="sig-lbl">✦ إعداد ✦</span>
    <span class="sig-rank">الملازم دج المهندس</span>
    <span class="sig-name">نورالدين امبارك</span>
  </div>
</div>`;

/* ---------- Excel: قراءة (لمتابعة ملف قديم فقط) ---------- */
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

/* ---------- GitHub: نسخة احتياطية تلقائية ---------- */
const ghHeaders = () => ({ Authorization: `Bearer ${GH.token}`, Accept: 'application/vnd.github+json', 'User-Agent': 'signals-log', 'X-GitHub-Api-Version': '2022-11-28' });
const ghApi = p => `https://api.github.com/repos/${GH.repo}${p}`;
let ghBranchReady = false;

async function ghEnsureBranch() {
  if (ghBranchReady) return;
  const r = await fetch(ghApi(`/git/ref/heads/${encodeURIComponent(GH.branch)}`), { headers: ghHeaders() });
  if (r.ok) { ghBranchReady = true; return; }
  if (r.status !== 404) throw new Error('GitHub ref ' + r.status);
  const repo = await (await fetch(ghApi(''), { headers: ghHeaders() })).json();
  const base = await (await fetch(ghApi(`/git/ref/heads/${encodeURIComponent(repo.default_branch)}`), { headers: ghHeaders() })).json();
  const c = await fetch(ghApi('/git/refs'), { method: 'POST', headers: { ...ghHeaders(), 'Content-Type': 'application/json' }, body: JSON.stringify({ ref: `refs/heads/${GH.branch}`, sha: base.object.sha }) });
  if (!c.ok) throw new Error('GitHub create branch ' + c.status + ' ' + await c.text());
  ghBranchReady = true;
}
async function ghGet() {
  const r = await fetch(ghApi(`/contents/${GH.file}?ref=${encodeURIComponent(GH.branch)}`), { headers: ghHeaders() });
  if (r.status === 404) return null;
  if (!r.ok) throw new Error('GitHub GET ' + r.status);
  const j = await r.json();
  return { sha: j.sha, text: Buffer.from(j.content, 'base64').toString('utf8') };
}
async function ghPush() {
  await ghEnsureBranch();
  const cur = await ghGet();
  const body = { message: `تحديث البيانات ${new Date().toISOString()}`, content: Buffer.from(serialize()).toString('base64'), branch: GH.branch };
  if (cur) body.sha = cur.sha;
  const r = await fetch(ghApi(`/contents/${GH.file}`), { method: 'PUT', headers: { ...ghHeaders(), 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error('GitHub PUT ' + r.status + ' ' + await r.text());
}
let ghTimer = null, ghBusy = false, ghDirty = false;
function scheduleGh() {
  if (!GH.enabled) return;
  ghDirty = true;
  clearTimeout(ghTimer);
  ghTimer = setTimeout(runGh, 2500); // تجميع الحفظات المتقاربة في رفع واحد
}
async function runGh() {
  if (ghBusy) { ghTimer = setTimeout(runGh, 2500); return; }
  ghBusy = true; ghDirty = false;
  try { await ghPush(); console.log('تمت مزامنة البيانات مع GitHub'); }
  catch (e) { console.error('فشلت مزامنة GitHub:', e.message); ghDirty = true; ghTimer = setTimeout(runGh, 30000); }
  finally { ghBusy = false; }
}

/* ---------- الحفظ (db.json ثم Excel، بالتتابع وبشكل آمن) ---------- */
async function doSave() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const tj = DB_PATH + '.tmp';
  fs.writeFileSync(tj, serialize());
  fs.renameSync(tj, DB_PATH);
  scheduleGh();

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('التحديثات', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }] });
  ws.columns = [6, 24, 22, 16, 24, 28, 14, 14, 10].map(w => ({ width: w }));
  ws.getColumn(9).hidden = true;
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
  const tmp = XLSX_PATH + '.tmp';
  await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, XLSX_PATH);
}
function saveToFile() {
  const p = saveChain.catch(() => {}).then(doSave);
  saveChain = p;
  return p;
}

/* ---------- تحميل البيانات عند التشغيل ---------- */
async function loadRecords() {
  if (fs.existsSync(DB_PATH)) {
    try { records = JSON.parse(fs.readFileSync(DB_PATH, 'utf8')).records || []; console.log('تم التحميل من db.json'); return; }
    catch (e) { console.error('db.json تالف:', e.message); }
  }
  if (GH.enabled) {
    try {
      const g = await ghGet();
      if (g) { records = JSON.parse(g.text).records || []; console.log('تم استرجاع البيانات من GitHub'); return; }
    } catch (e) { console.error('تعذر الاسترجاع من GitHub:', e.message); }
  }
  await loadFromFile();
  if (records.length) console.log('تم التحميل من ملف Excel');
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
function sanitizeRestored(arr) {
  const s = v => String(v ?? '').trim().slice(0, 200);
  return arr.filter(r => r && typeof r === 'object').map(r => ({
    id: String(r.id || '').replace(/[^\w-]/g, '').slice(0, 64) || crypto.randomUUID(),
    serial: s(r.serial), programmer: s(r.programmer), type: TYPES.includes(r.type) ? r.type : '',
    unit: s(r.unit), location: s(r.location), date: s(r.date), time: s(r.time)
  }));
}
function readBody(req, limit = 1e5) {
  return new Promise((ok, no) => {
    let d = '';
    req.on('data', c => { d += c; if (d.length > limit) { req.destroy(); no(httpErr(413, 'الطلب كبير جداً')); } });
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
/* حماية من تخمين كلمة المرور: 10 محاولات خاطئة = حظر 15 دقيقة */
const fails = new Map();
const ipOf = req => String(req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').split(',')[0].trim();
function locked(ip) {
  const f = fails.get(ip);
  if (!f || f.n < 10) return false;
  if (Date.now() - f.t < 15 * 60e3) return true;
  fails.delete(ip); return false;
}
const snapshot = () => ({ version, records });

const RESTORE_PAGE = `<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>استرجاع نسخة احتياطية</title>
<body style="font-family:Tahoma,sans-serif;background:#0a1530;color:#f6efd6;display:grid;place-items:center;min-height:100vh;margin:0">
<div style="background:#12224a;border:1px solid #d4af37;border-radius:16px;padding:28px;max-width:420px;width:90%;text-align:center">
<h2 style="color:#d4af37;margin-top:0">استرجاع نسخة احتياطية</h2>
<p>اختر ملف db.json / backup.json. <b>سيستبدل البيانات الحالية.</b></p>
<input type="file" id="f" accept=".json"><br><br>
<button id="b" style="background:#d4af37;border:0;border-radius:10px;padding:10px 26px;font-weight:700;cursor:pointer">استرجاع</button>
<p id="m"></p><p><a href="/" style="color:#d4af37">عودة للسجل</a> · <a href="/api/backup" style="color:#d4af37">تنزيل نسخة احتياطية</a></p></div>
<script>b.onclick=async()=>{const x=f.files[0];if(!x)return m.textContent='اختر ملفاً أولاً';
try{const r=await fetch('/api/restore',{method:'POST',headers:{'Content-Type':'application/json'},body:await x.text()});
const j=await r.json();m.textContent=r.ok?'تم استرجاع '+j.count+' سجل ✔':(j.error||'فشل')}catch(e){m.textContent='فشل الاسترجاع'}}</script></body></html>`;

/* ---------- الخادم ---------- */
const server = http.createServer(async (req, res) => {
  try {
    const p = new URL(req.url, 'http://x').pathname;
    if (p === '/healthz') { res.writeHead(200, { 'Content-Type': 'text/plain' }); return res.end('ok'); }

    if (AUTH_ON) {
      const ip = ipOf(req);
      if (locked(ip)) { res.writeHead(429, { 'Content-Type': 'text/plain; charset=utf-8' }); return res.end('محاولات كثيرة، حاول بعد 15 دقيقة'); }
      if (!authorized(req)) {
        if (req.headers.authorization) { const f = fails.get(ip) || { n: 0, t: 0 }; fails.set(ip, { n: f.n + 1, t: Date.now() }); }
        res.writeHead(401, { 'WWW-Authenticate': 'Basic realm="Signals", charset="UTF-8"' });
        return res.end('Unauthorized');
      }
      fails.delete(ip);
    }

    if (req.method === 'GET' && (p === '/' || p === '/index.html')) {
      const file = path.join(__dirname, 'public', 'index.html');
      if (!fs.existsSync(file)) throw httpErr(404, 'الملف public/index.html غير موجود');
      let html = fs.readFileSync(file, 'utf8');
      html = html.includes('</body>') ? html.replace('</body>', () => CREDIT + '</body>') : html + CREDIT;
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' });
      return res.end(html);
    }
    if (req.method === 'GET' && p === '/restore') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
      return res.end(RESTORE_PAGE);
    }
    if (req.method === 'GET' && p === '/api/backup') {
      res.writeHead(200, {
        'Content-Type': 'application/json; charset=utf-8',
        'Content-Disposition': `attachment; filename="backup-${stamp().date}.json"`,
        'Cache-Control': 'no-store'
      });
      return res.end(serialize());
    }
    if (req.method === 'POST' && p === '/api/restore') {
      const b = await readBody(req, 5e6);
      const arr = Array.isArray(b) ? b : b.records;
      if (!Array.isArray(arr)) throw httpErr(400, 'ملف غير صالح');
      records = sanitizeRestored(arr); version++;
      await saveToFile();
      return send(res, 200, { ok: true, count: records.length });
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

/* رفع آخر تغييرات إلى GitHub عند إيقاف السيرفر (السيرفرات المجانية تنام وتُعاد بانتظام) */
for (const sig of ['SIGTERM', 'SIGINT']) {
  process.on(sig, async () => {
    try { if (GH.enabled && ghDirty) await ghPush(); } catch (e) { console.error(e.message); }
    process.exit(0);
  });
}

loadRecords()
  .catch(e => console.error('تعذرت قراءة البيانات:', e.message))
  .then(async () => { if (records.length && !fs.existsSync(DB_PATH)) await saveToFile().catch(e => console.error(e.message)); })
  .then(() => server.listen(PORT, () => {
    console.log(`الخادم يعمل على المنفذ ${PORT} — عدد السجلات: ${records.length}`);
    console.log(AUTH_ON ? 'الحماية بكلمة مرور: مفعّلة' : '⚠ الحماية بكلمة مرور: معطّلة (عيّن APP_USER و APP_PASS)');
    console.log(GH.enabled ? `مزامنة GitHub: مفعّلة (${GH.repo} / ${GH.branch})` : 'مزامنة GitHub: غير مفعّلة');
  }));