// 시험용 가짜 구글 서버: backend/Code.gs 를 그대로 node 에서 돌려요.
// node backend/mock/server.js [포트=8787]  → http://127.0.0.1:8787/index.html?ep=http://127.0.0.1:8787/api
// 진짜 Apps Script 대신 표·캐시·잠금을 메모리로 흉내 내요. flush 트리거는 5초마다 돌아요.
'use strict';
const fs = require('fs'), path = require('path'), http = require('http'), crypto = require('crypto'), vm = require('vm');
const ROOT = path.resolve(__dirname, '..', '..');
const PORT = +(process.argv[2] || 8787);

function makeSheet(name) {
  const sh = { name, data: [], hidden: [],
    getLastRow() { let n = this.data.length; while (n && this.data[n - 1].every(v => v === '' || v == null)) n--; return n; },
    getMaxRows() { return Math.max(1000, this.data.length); },
    getRange(r, c, nr = 1, nc = 1) { const s = this; return {
      setValues(v) { for (let i = 0; i < nr; i++) { while (s.data.length < r + i) s.data.push([]); const row = s.data[r - 1 + i]; for (let j = 0; j < nc; j++) row[c - 1 + j] = v[i][j]; } return this; },
      setValue(v) { return this.setValues([[v]]); },
      getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = s.data[r - 1 + i] || []; const o = []; for (let j = 0; j < nc; j++) o.push(row[c - 1 + j] ?? ''); out.push(o); } return out; },
      clearContent() { for (let i = 0; i < nr; i++) { const row = s.data[r - 1 + i]; if (row) for (let j = 0; j < nc; j++) row[c - 1 + j] = ''; } return this; },
      setFontWeight() { return this; }, setBackground() { return this; }, setFontSize() { return this; } }; },
    appendRow(v) { this.data.length = this.getLastRow(); this.data.push(v.slice()); },
    setFrozenRows() {}, hideColumns() {}, setColumnWidth() {}, deleteRows() {} };
  return sh;
}
const ss = { title: '', sheets: {}, rename(t) { this.title = t; }, getId() { return 'mock'; },
  getSheetByName(n) { return this.sheets[n] || null; }, insertSheet(n) { return (this.sheets[n] = makeSheet(n)); },
  getSheets() { return Object.values(this.sheets); }, deleteSheet(s) { delete this.sheets[s.name]; } };
const store = new Map();
const cache = { get(k) { const e = store.get(k); if (!e) return null; if (e.t < Date.now()) { store.delete(k); return null; } return e.v; },
  put(k, v, ttl = 600) { if (String(v).length > 100000) throw new Error('too big'); store.set(k, { v: String(v), t: Date.now() + ttl * 1000 }); },
  getAll(keys) { const o = {}; keys.forEach(k => { const v = cache.get(k); if (v != null) o[k] = v; }); return o; },
  remove(k) { store.delete(k); }, removeAll(ks) { ks.forEach(k => store.delete(k)); } };
const lock = { tryLock() { return true; }, releaseLock() {} };
const props = {};
const ctx = {
  console, Date, Math, JSON, String, Number, Object, Array, Error, isFinite,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, create: () => ss },
  CacheService: { getScriptCache: () => cache },
  LockService: { getDocumentLock: () => lock, getScriptLock: () => lock },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => { const b = { timeBased: () => b, everyMinutes: () => b, everyDays: () => b, atHour: () => b, create: () => b }; return b; } },
  Utilities: { getUuid: () => crypto.randomUUID(), DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (a, s) => Array.from(crypto.createHash('sha256').update(s, 'utf8').digest()).map(b => b > 127 ? b - 256 : b) },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ text: t, setMimeType() { return this; } }) }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);
if (process.env.MOCK_SMALL) vm.runInContext('MAX_STUDENT_ROWS=50;MAX_CLASSES=5;', ctx); // 한도 시험용
console.log(ctx.setup());
setInterval(() => { try { ctx.flush(); } catch (e) { console.error('flush', e); } }, 5000);
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.png': 'image/png', '.json': 'application/json' };
http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*'); res.setHeader('Cache-Control', 'no-store');
  const url = decodeURIComponent(req.url.split('?')[0]);
  if (url === '/api' && req.method === 'POST') {
    let b = ''; req.on('data', c => b += c); req.on('end', () => {
      const out = ctx.doPost({ postData: { contents: b } }); res.setHeader('Content-Type', 'application/json'); res.end(out.text); });
    return;
  }
  if (url === '/__sheet') { res.setHeader('Content-Type', 'application/json; charset=utf-8'); res.end(JSON.stringify({ title: ss.title, sheets: Object.fromEntries(Object.entries(ss.sheets).map(([k, v]) => [k, v.data])) })); return; }
  if (url === '/__flush') { ctx.flush(); res.end('ok'); return; }
  const f = path.join(ROOT, url === '/' ? 'index.html' : url);
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.statusCode = 404; res.end('404'); return; }
  res.setHeader('Content-Type', TYPES[path.extname(f)] || 'application/octet-stream'); fs.createReadStream(f).pipe(res);
}).listen(PORT, '127.0.0.1', () => console.log('mock on http://127.0.0.1:' + PORT));
