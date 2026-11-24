/**
 * 현장체험 안전 길잡이 3D · 우리 반 기록 서버 (Google Apps Script)
 *
 * 설치 (처음 한 번)
 *  1) 새 구글 시트를 만들고 → 확장 프로그램 → Apps Script → 이 파일 내용을 모두 붙여 넣고 저장
 *  2) 위쪽 함수 목록에서 setup 을 골라 ▶ 실행 → 권한 허용
 *     · 시트 제목이 '[26] 현장체험 안전 3D · 우리 반 기록'으로 바뀌고, 탭 3개(안내·반 목록·학생 기록)와
 *       1분마다 기록을 옮겨 적는 트리거, 매일 새벽 오래된 기록을 정리하는 트리거가 만들어져요.
 *  3) 배포 → 새 배포 → 유형: 웹 앱 → 실행: 나 / 액세스 권한: 모든 사용자 → 배포
 *  4) 나온 주소(…/exec)를 게임 폴더의 net-config.js endpoint 에 넣어요.
 *
 * 시트가 터지지 않게 막는 장치
 *  · 학생 기록은 한 사람당 한 줄만 써요(덮어쓰기). 행동마다 줄을 늘리지 않아요.
 *  · 게임이 보내는 진행 기록은 먼저 캐시에만 담고, 1분마다 트리거가 한 번에 모아서 시트에 써요.
 *    → 30명이 동시에 보내도 시트 쓰기는 1분에 한 번이에요.
 *  · 같은 학생이 2초 안에 다시 보내면 무시하고, 서버 전체로도 1분에 1200건이 넘으면 잠깐 쉬어요.
 *  · 반은 최대 300개, 1시간에 새로 만들 수 있는 반은 30개, 한 반 번호는 1~40번이에요.
 *  · 글자 길이를 자르고, = + - @ 로 시작하는 글은 수식이 되지 않게 막아요.
 *  · 가득 차면 멈추지 않고 알아서 지워요: 학생 기록이 12000줄을 넘으면 그 자리에서 가장 오래 안 쓴 기록부터,
 *    반이 300개가 차면 가장 오래 안 쓴 반부터 지우고 새 반을 만들어요. (지금 쓰는 반 기록은 남아요)
 *  · 매일 새벽: 60일 넘게 쓰지 않은 반과 그 반 학생 기록을 지워요.
 *  · 선생님 PIN을 5번 틀리면 5분 동안 막아요.
 *
 * 요청: POST, 본문은 JSON 문자열 (Content-Type: text/plain)
 *   학생: join, progress      선생님: createClass, dashboard, setOpen, deleteClass (PIN 필요)
 */

var TITLE = '[26] 현장체험 안전 3D · 우리 반 기록';
var SH_INFO = '안내', SH_CLASSES = '반 목록', SH_STUDENTS = '학생 기록';
var HEAD_CLASSES = ['반 코드', '반 이름', 'pinHash', 'salt', '참여 열림', '만든 때', '마지막 사용'];
var HEAD_STUDENTS = ['key', '반 코드', '번호', '별명', '장소', '단계', '지금 할 일', '안전 점수', '퀴즈 맞힘', '퀴즈 수', '실수', '하이파이브', '도착', '처음', '마지막'];
var C = { code: 0, name: 1, pinHash: 2, salt: 3, open: 4, created: 5, used: 6 };
var S = { key: 0, code: 1, num: 2, nick: 3, world: 4, step: 5, stepText: 6, score: 7, quiz: 8, quizN: 9, miss: 10, hf: 11, done: 12, first: 13, last: 14 };
var CODE_ALPHABET = 'ACDEFGHJKMNPQRTUVWXY34679';
var MAX_CLASSES = 300, MAX_NUM = 40, MAX_STUDENT_ROWS = 12000, CLASS_IDLE_DAYS = 60;
var BUF_TTL = 21600;          // 캐시에 담아 두는 시간(6시간)
var GLOBAL_PER_MIN = 1200;    // 서버 전체 1분 상한
var WORLDS = { A: '학교→모래내시장역', ride: '열차 안', C: '시민공원역→CGV' };
var LESSON_KEYS = { pri: 1, red: 1, jay: 1, car: 1, run: 1, esc: 1, psd: 1, line: 1, first: 1, door: 1, miss: 1, bike: 1, drive: 1 };

// ---------------------------------------------------------------- 설치

function setup() {
  var ss = ss_();
  ss.rename(TITLE);
  var info = sheetOf_(SH_INFO, ['현장체험 안전 길잡이 3D · 우리 반 기록']);
  info.getRange(2, 1, 6, 1).setValues([
    ['이 시트는 게임 서버가 자동으로 써요. 표를 손으로 고치거나 지우지 마세요.'],
    ['반 목록: 선생님 화면에서 만든 반(반 코드·이름). PIN은 암호로 바꿔 저장해요.'],
    ['학생 기록: 학생 한 명당 한 줄(별명·장소·단계·안전 점수·퀴즈·실수). 1분마다 새로 고쳐져요.'],
    ['기록 보기는 선생님 화면(teacher.html)에서 반 코드와 PIN으로 해요.'],
    ['60일 넘게 쓰지 않은 반과 학생 기록은 매일 새벽 자동으로 지워져요.'],
    ['학생 진짜 이름은 받지 않아요. 별명과 번호만 써요.']
  ]);
  info.setColumnWidth(1, 640);
  info.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  var cl = sheetOf_(SH_CLASSES, HEAD_CLASSES);
  cl.hideColumns(C.pinHash + 1, 2);
  sheetOf_(SH_STUDENTS, HEAD_STUDENTS).hideColumns(S.key + 1);
  var extra = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (extra && ss.getSheets().length > 1 && extra.getLastRow() === 0) ss.deleteSheet(extra);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'flush' || f === 'cleanup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('flush').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('cleanup').timeBased().everyDays(1).atHour(3).create();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  return '설치 끝: ' + TITLE;
}

function ss_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create(TITLE); // 시트에 붙이지 않고 만든 스크립트면 새 시트를 만들어요
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  return ss;
}
function sheetOf_(name, head) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (head && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#fde7c8');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(sh, width) {
  var n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, width).getValues();
}

// ---------------------------------------------------------------- 진입점

function doGet() { return json_({ ok: true, app: 'safety3d', msg: '연결 잘 됐어요!' }); }

function doPost(e) {
  var req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad', msg: '요청 형식이 잘못됐어요.' }); }
  try { return json_(route_(req || {})); }
  catch (err) {
    var code = (err && err.userCode) || 'server';
    if (!err || !err.userCode) console.error(err && err.stack || err);
    return json_({ ok: false, error: code, msg: (err && err.userMsg) || '서버에서 문제가 생겼어요. 잠시 뒤 다시 해 보세요.' });
  }
}

function route_(req) {
  if (!globalBudget_()) fail_('busy', '지금 사람이 많아요. 잠시 뒤 다시 해요.');
  switch (String(req.action || '')) {
    case 'ping': return { ok: true, now: Date.now() };
    case 'join': return join_(req);
    case 'progress': return progress_(req);
    case 'createClass': return createClass_(req);
    case 'dashboard': return dashboard_(req);
    case 'setOpen': return setOpen_(req);
    case 'deleteClass': return deleteClass_(req);
  }
  fail_('bad', '알 수 없는 요청이에요.');
}

// ---------------------------------------------------------------- 학생

function join_(req) {
  var who = student_(req);
  var cls = getClass_(who.code);
  if (!cls) fail_('noclass', '반 코드를 찾을 수 없어요. 선생님께 다시 물어보세요.');
  if (!cls.open) fail_('closed', '지금은 들어갈 수 없는 반이에요.');
  var cache = CacheService.getScriptCache();
  cache.put('used_' + who.code, '1', BUF_TTL);
  var k = 'b_' + who.code + '-' + who.num, old = parse_(cache.get(k));
  var rec = old && old.nick === who.nick ? old : { nick: who.nick, world: 'A', step: 0, stepText: '', score: 100, quiz: 0, quizN: 0, miss: {}, hf: 0, done: false };
  rec.nick = who.nick; rec.t = Date.now(); rec.first = (old && old.first) || rec.t;
  cache.put(k, JSON.stringify(rec), BUF_TTL);
  return { ok: true, code: who.code, name: cls.name, num: who.num, nick: who.nick };
}

function progress_(req) {
  var who = student_(req);
  var cache = CacheService.getScriptCache();
  var k = 'b_' + who.code + '-' + who.num;
  if (cache.get('rl_' + k) && !req.done) return { ok: true, skipped: true }; // 2초 안에 또 오면 무시
  cache.put('rl_' + k, '1', 2);
  if (!cache.get('used_' + who.code)) {           // 캐시가 비었으면 반이 맞는지 한 번 더 봐요
    var cls = getClass_(who.code);
    if (!cls) fail_('noclass', '반 코드를 찾을 수 없어요.');
    cache.put('used_' + who.code, '1', BUF_TTL);
  }
  var old = parse_(cache.get(k)) || {};
  var miss = {};
  if (req.miss && typeof req.miss === 'object') Object.keys(req.miss).slice(0, 20).forEach(function (m) {
    if (LESSON_KEYS[m]) miss[m] = int_(req.miss[m], 0, 99);
  });
  var rec = {
    nick: who.nick, world: WORLDS[req.world] ? String(req.world) : 'A', step: int_(req.step, 0, 30),
    stepText: cleanText_(req.stepText, 60), score: int_(req.score, 0, 100), quiz: int_(req.quiz, 0, 30), quizN: int_(req.quizN, 0, 30),
    miss: miss, hf: int_(req.hf, 0, 999), done: !!req.done || !!old.done, t: Date.now(), first: old.first || Date.now()
  };
  if (rec.done && !old.done) rec.doneAt = Date.now(); else if (old.doneAt) rec.doneAt = old.doneAt;
  cache.put(k, JSON.stringify(rec), BUF_TTL);
  return { ok: true };
}

// ---------------------------------------------------------------- 캐시 → 시트 (1분마다)

function flush() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return;
  try {
    var cache = CacheService.getScriptCache();
    var csh = sheetOf_(SH_CLASSES, HEAD_CLASSES), classes = rows_(csh, HEAD_CLASSES.length);
    if (!classes.length) return;
    var usedKeys = classes.map(function (r) { return 'used_' + r[C.code]; });
    var used = cache.getAll(usedKeys);
    var active = classes.filter(function (r) { return used['used_' + r[C.code]]; }).map(function (r) { return String(r[C.code]); });
    if (!active.length) return;
    var keys = [];
    active.forEach(function (code) { for (var n = 1; n <= MAX_NUM; n++) keys.push('b_' + code + '-' + n); });
    var buf = {};
    for (var i = 0; i < keys.length; i += 100) {
      var part = cache.getAll(keys.slice(i, i + 100));
      Object.keys(part).forEach(function (k) { buf[k] = part[k]; });
    }
    var ssh = sheetOf_(SH_STUDENTS, HEAD_STUDENTS), data = rows_(ssh, HEAD_STUDENTS.length), idx = {};
    data.forEach(function (r, j) { idx[String(r[S.key])] = j; });
    var changed = false, appended = [];
    Object.keys(buf).forEach(function (k) {
      var rec = parse_(buf[k]); if (!rec) return;
      var key = k.slice(2), parts = key.split('-'), row = idx[key] != null ? data[idx[key]] : null;
      var last = row ? Number(row[S.last]) || 0 : 0;
      if (row && last >= rec.t) return; // 이미 쓴 것
      var out = [key, parts[0], Number(parts[1]), safe_(rec.nick), WORLDS[rec.world] || '', rec.step, safe_(rec.stepText), rec.score, rec.quiz, rec.quizN,
        JSON.stringify(rec.miss || {}), rec.hf || 0, rec.done ? (rec.doneAt ? new Date(rec.doneAt) : '도착') : '', new Date(rec.first || rec.t), rec.t];
      if (row) { data[idx[key]] = out; changed = true; } else appended.push(out);
    });
    if (changed) ssh.getRange(2, 1, data.length, HEAD_STUDENTS.length).setValues(data);
    if (appended.length) ssh.getRange(ssh.getLastRow() + 1, 1, appended.length, HEAD_STUDENTS.length).setValues(appended);
    var total = data.length + appended.length;
    if (total > MAX_STUDENT_ROWS) { // 가득 참 → 가장 오래 안 쓴 기록부터 지워요
      var sorted = data.concat(appended).sort(function (a, b) { return (Number(b[S.last]) || 0) - (Number(a[S.last]) || 0); });
      var keepN = Math.floor(MAX_STUDENT_ROWS * 0.9);
      cache.removeAll(sorted.slice(keepN).map(function (r) { return 'b_' + r[S.key]; })); // 지운 기록이 캐시에서 되살아나지 않게
      rewrite_(ssh, total, sorted.slice(0, keepN), HEAD_STUDENTS.length);
    }
    // 반 마지막 사용 시각
    var now = new Date(), touched = false;
    classes.forEach(function (r) { if (used['used_' + r[C.code]]) { r[C.used] = now; touched = true; } });
    if (touched) csh.getRange(2, 1, classes.length, HEAD_CLASSES.length).setValues(classes);
  } finally { lock.releaseLock(); }
}

// 매일 새벽: 오래된 반·학생 기록 정리, 줄 수 상한
function cleanup() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;
  try {
    var csh = sheetOf_(SH_CLASSES, HEAD_CLASSES), classes = rows_(csh, HEAD_CLASSES.length);
    var cut = Date.now() - CLASS_IDLE_DAYS * 86400000, keep = {}, keptClasses = [];
    classes.forEach(function (r) {
      var t = new Date(r[C.used] || r[C.created]).getTime() || 0;
      if (r[C.pinHash] === 'deleted' || t < cut) return;
      keep[String(r[C.code])] = 1; keptClasses.push(r);
    });
    rewrite_(csh, classes.length, keptClasses, HEAD_CLASSES.length);
    var ssh = sheetOf_(SH_STUDENTS, HEAD_STUDENTS), data = rows_(ssh, HEAD_STUDENTS.length);
    var kept = data.filter(function (r) { return keep[String(r[S.code])]; });
    if (kept.length > MAX_STUDENT_ROWS) {
      kept.sort(function (a, b) { return (Number(b[S.last]) || 0) - (Number(a[S.last]) || 0); });
      kept = kept.slice(0, MAX_STUDENT_ROWS);
    }
    rewrite_(ssh, data.length, kept, HEAD_STUDENTS.length);
    CacheService.getScriptCache().remove('classes');
  } finally { lock.releaseLock(); }
}
function rewrite_(sh, oldN, rows, width) {
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
  if (oldN > rows.length) sh.getRange(2 + rows.length, 1, oldN - rows.length, width).clearContent();
  var extra = sh.getMaxRows() - Math.max(rows.length + 1, 200);
  if (extra > 500) sh.deleteRows(Math.max(rows.length + 2, 201), extra); // 빈 줄이 너무 쌓이지 않게
}

// ---------------------------------------------------------------- 선생님

function createClass_(req) {
  var name = cleanText_(req.name, 20);
  if (!name) fail_('bad', '반 이름을 적어 주세요. (예: 정각초 5학년 3반)');
  var pin = String(req.pin == null ? '' : req.pin);
  if (!/^\d{6,8}$/.test(pin)) fail_('bad', 'PIN은 숫자 6~8자리로 정해 주세요.');
  var cache = CacheService.getScriptCache(), cnt = Number(cache.get('mk_count') || 0);
  if (cnt >= 30) fail_('busy', '반을 너무 많이 만들었어요. 1시간 뒤 다시 해 보세요.');
  cache.put('mk_count', String(cnt + 1), 3600);
  return withLock_(function () {
    var sh = sheetOf_(SH_CLASSES, HEAD_CLASSES), rows = rows_(sh, HEAD_CLASSES.length);
    if (rows.length >= MAX_CLASSES) { // 가득 참 → 가장 오래 안 쓴 반 10%를 지우고 계속해요
      var t = function (r) { return new Date(r[C.used] || r[C.created]).getTime() || 0; };
      var sorted = rows.slice().sort(function (a, b) { return t(b) - t(a); });
      var keepN = Math.floor(MAX_CLASSES * 0.9), gone = {};
      sorted.slice(keepN).forEach(function (r) { gone[String(r[C.code])] = 1; });
      var oldN = rows.length; rows = sorted.slice(0, keepN);
      rewrite_(sh, oldN, rows, HEAD_CLASSES.length);
      var ssh = sheetOf_(SH_STUDENTS, HEAD_STUDENTS), data = rows_(ssh, HEAD_STUDENTS.length);
      rewrite_(ssh, data.length, data.filter(function (r) { return !gone[String(r[S.code])]; }), HEAD_STUDENTS.length);
    }
    var used = {}; rows.forEach(function (r) { used[String(r[C.code])] = 1; });
    var code = '';
    for (var tries = 0; tries < 50 && !code; tries++) { var c = randomCode_(5); if (!used[c] && /[A-Z]/.test(c)) code = c; }
    if (!code) fail_('server', '반 코드를 만들지 못했어요. 다시 해 보세요.');
    var salt = Utilities.getUuid();
    sh.appendRow([code, safe_(name), hash_(pin, salt), salt, true, new Date(), new Date()]);
    cache.remove('classes');
    return { ok: true, code: code, name: name };
  });
}

function dashboard_(req) {
  var cls = teacher_(req);
  var rows = rows_(sheetOf_(SH_STUDENTS, HEAD_STUDENTS), HEAD_STUDENTS.length).filter(function (r) { return String(r[S.code]) === cls.code; });
  var list = {};
  rows.forEach(function (r) {
    var miss = parse_(r[S.miss]) || {};
    list[r[S.num]] = { num: Number(r[S.num]), nick: String(r[S.nick]), where: String(r[S.world]), step: Number(r[S.step]), stepText: String(r[S.stepText]),
      score: Number(r[S.score]), quiz: Number(r[S.quiz]), quizN: Number(r[S.quizN]), miss: miss, hf: Number(r[S.hf]) || 0, done: !!r[S.done], last: Number(r[S.last]) || 0 };
  });
  // 아직 시트로 옮기지 않은 최신 기록(캐시)도 함께 보여 줘요
  var keys = []; for (var n = 1; n <= MAX_NUM; n++) keys.push('b_' + cls.code + '-' + n);
  var buf = CacheService.getScriptCache().getAll(keys);
  Object.keys(buf).forEach(function (k) {
    var rec = parse_(buf[k]); if (!rec) return; var num = Number(k.split('-')[1]);
    if (list[num] && list[num].last >= rec.t) return;
    list[num] = { num: num, nick: rec.nick, where: WORLDS[rec.world] || '', step: rec.step, stepText: rec.stepText, score: rec.score, quiz: rec.quiz, quizN: rec.quizN,
      miss: rec.miss || {}, hf: rec.hf || 0, done: !!rec.done, last: rec.t };
  });
  var students = Object.keys(list).map(function (k) { return list[k]; }).sort(function (a, b) { return a.num - b.num; });
  return { ok: true, now: Date.now(), cls: { code: cls.code, name: cls.name, open: cls.open }, students: students };
}

function setOpen_(req) {
  var cls = teacher_(req), open = !!req.open;
  return withLock_(function () {
    var sh = sheetOf_(SH_CLASSES, HEAD_CLASSES), rows = rows_(sh, HEAD_CLASSES.length);
    for (var i = 0; i < rows.length; i++) if (String(rows[i][C.code]) === cls.code) { sh.getRange(i + 2, C.open + 1).setValue(open); break; }
    CacheService.getScriptCache().remove('classes');
    return { ok: true, open: open };
  });
}

function deleteClass_(req) {
  var cls = teacher_(req);
  return withLock_(function () {
    var sh = sheetOf_(SH_CLASSES, HEAD_CLASSES), rows = rows_(sh, HEAD_CLASSES.length);
    for (var i = 0; i < rows.length; i++) if (String(rows[i][C.code]) === cls.code) { sh.getRange(i + 2, C.pinHash + 1, 1, 3).setValues([['deleted', '', false]]); break; }
    var ssh = sheetOf_(SH_STUDENTS, HEAD_STUDENTS), data = rows_(ssh, HEAD_STUDENTS.length);
    rewrite_(ssh, data.length, data.filter(function (r) { return String(r[S.code]) !== cls.code; }), HEAD_STUDENTS.length);
    var cache = CacheService.getScriptCache(), keys = ['classes', 'used_' + cls.code];
    for (var n = 1; n <= MAX_NUM; n++) keys.push('b_' + cls.code + '-' + n);
    cache.removeAll(keys);
    return { ok: true };
  });
}

// ---------------------------------------------------------------- 도우미

function student_(req) {
  var code = String(req.code || '').toUpperCase();
  if (!/^[A-Z0-9]{5}$/.test(code)) fail_('bad', '반 코드는 영어·숫자 5자예요.');
  var num = Number(req.num);
  if (!(num >= 1 && num <= MAX_NUM && num === Math.floor(num))) fail_('bad', '번호는 1~' + MAX_NUM + '이에요.');
  var nick = cleanText_(req.nick, 8);
  if (!nick) fail_('bad', '별명을 적어 주세요.');
  return { code: code, num: num, nick: nick };
}
function classes_() {
  var cache = CacheService.getScriptCache(), c = parse_(cache.get('classes'));
  if (c) return c;
  var out = {};
  rows_(sheetOf_(SH_CLASSES, HEAD_CLASSES), HEAD_CLASSES.length).forEach(function (r) {
    if (r[C.pinHash] === 'deleted') return;
    out[String(r[C.code])] = { code: String(r[C.code]), name: String(r[C.name]), pinHash: String(r[C.pinHash]), salt: String(r[C.salt]), open: r[C.open] === true || r[C.open] === 'TRUE' };
  });
  try { cache.put('classes', JSON.stringify(out), 600); } catch (e) {}
  return out;
}
function getClass_(code) { return classes_()[code] || null; }
function teacher_(req) {
  var code = String(req.code || '').toUpperCase(), cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pf_' + code) || 0);
  if (fails >= 5) fail_('locked', 'PIN을 여러 번 틀렸어요. 5분 뒤 다시 해 보세요.');
  var cls = getClass_(code);
  if (!cls) fail_('noclass', '반 코드를 찾을 수 없어요.');
  if (hash_(String(req.pin || ''), cls.salt) !== cls.pinHash) { cache.put('pf_' + code, String(fails + 1), 300); fail_('pin', 'PIN이 맞지 않아요.'); }
  cache.remove('pf_' + code);
  return cls;
}
function globalBudget_() {
  var cache = CacheService.getScriptCache(), k = 'g_' + Math.floor(Date.now() / 60000), n = Number(cache.get(k) || 0);
  if (n >= GLOBAL_PER_MIN) return false;
  cache.put(k, String(n + 1), 120); // 대략 세기(정확하지 않아도 괜찮아요)
  return true;
}
function withLock_(fn) {
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) fail_('busy', '서버가 바빠요. 잠시 뒤 다시 해 보세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}
function hash_(pin, salt) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function randomCode_(n) { var s = ''; for (var i = 0; i < n; i++) s += CODE_ALPHABET.charAt(Math.floor(Math.random() * CODE_ALPHABET.length)); return s; }
function cleanText_(v, n) { return String(v == null ? '' : v).replace(/[<>&"'`\\]/g, '').replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, n); }
function safe_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function int_(v, lo, hi) { var n = Math.round(Number(v)); if (!isFinite(n)) n = lo; return Math.max(lo, Math.min(hi, n)); }
function parse_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function fail_(code, msg) { var e = new Error(msg); e.userCode = code; e.userMsg = msg; throw e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
