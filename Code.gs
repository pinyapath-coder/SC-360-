/**
 * ระบบแบบประเมิน 360 องศา ฝ่ายขายรถเล็ก (LCV) – RIS 2569
 * Backend API (Google Apps Script) สำหรับแอป PWA – อ่าน/เขียนข้อมูลใน Google Sheet ที่ผูกอยู่
 * ชีตที่ใช้: ตั้งค่า, พนักงาน, คู่ประเมิน, หัวข้อประเมิน, เกณฑ์คะแนน, คำถามปลายเปิด, คำตอบ
 */

var TZ = 'Asia/Bangkok';
var SH = {
  SET: 'ตั้งค่า', EMP: 'พนักงาน', PAIR: 'คู่ประเมิน', ITEM: 'หัวข้อประเมิน',
  SCALE: 'เกณฑ์คะแนน', OPEN: 'คำถามปลายเปิด', RESP: 'คำตอบ'
};
var REL_ORDER = ['ตนเอง', 'หัวหน้า', 'เพื่อน', 'ลูกน้อง'];

/* ---------------- API entry (เรียกจากแอป PWA บน GitHub Pages) ---------------- */

var ACTIONS = {
  info:     function (b) { return apiInfo(); },
  login:    function (b) { return apiLogin(b.id, b.pin); },
  myAnswer: function (b) { return apiMyAnswer(b.id, b.pin, b.rateeId); },
  submit:   function (b) { return apiSubmit(b.id, b.pin, b.data || {}); },
  admin:    function (b) { return apiAdmin(b.pw); },
  report:   function (b) { return apiReport(b.pw, b.rateeId); }
};

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** แอปส่ง POST แบบ text/plain (เลี่ยง CORS preflight) เนื้อหาเป็น JSON {action, ...} */
function doPost(e) {
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = ACTIONS[b.action];
    if (!fn) throw new Error('ไม่รู้จักคำสั่ง ' + b.action);
    return json_({ ok: true, data: fn(b) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** เปิด URL ใน browser เพื่อตรวจว่า API ทำงาน */
function doGet() {
  return json_({ ok: true, service: 'RIS 360 API', title: settings_().title });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('360 องศา')
    .addItem('ดู URL ของ API (หลัง Deploy)', 'showUrl_')
    .addToUi();
}

function showUrl_() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? 'URL ของ API (นำไปใส่ใน config.js):\n' + url : 'ยังไม่ได้ Deploy เป็น Web app');
}

/* ---------------- Sheet helpers ---------------- */

function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต "' + name + '"');
  return sh;
}

function rows_(name) {
  var v = sheet_(name).getDataRange().getValues();
  var h = v.shift().map(function (x) { return String(x).trim(); });
  return v.filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    var o = {};
    h.forEach(function (k, i) { o[k] = r[i]; });
    return o;
  });
}

function nowTh_() {
  var d = new Date();
  return Utilities.formatDate(d, TZ, 'dd/MM/') + (Number(Utilities.formatDate(d, TZ, 'yyyy')) + 543) +
    Utilities.formatDate(d, TZ, ' HH:mm');
}

function id_(v) { return String(v == null ? '' : v).trim(); }

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return id_(v);
}

function settings_() {
  var m = {};
  rows_(SH.SET).forEach(function (r) { m[id_(r['หัวข้อ'])] = r['ค่า']; });
  var yes = function (v) { return /^y/i.test(id_(v)); };
  return {
    title: id_(m['ชื่อแบบประเมิน']) || 'แบบประเมิน 360 องศา',
    openDate: dateStr_(m['วันเปิด']),
    closeDate: dateStr_(m['วันปิด']),
    testMode: yes(m['โหมดทดสอบ']),
    hrPassword: id_(m['รหัสผ่าน HR']),
    usePin: yes(m['ใช้ PIN']),
    w: { 'หัวหน้า': Number(m['น้ำหนักหัวหน้า']) || 0,
         'เพื่อน': Number(m['น้ำหนักเพื่อน']) || 0,
         'ลูกน้อง': Number(m['น้ำหนักลูกน้อง']) || 0 },
    minSub: Number(m['ลูกน้องขั้นต่ำ']) || 3
  };
}

function isOpen_(s) {
  if (s.testMode) return true;
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return today >= s.openDate && today <= s.closeDate;
}

function employees_() {
  var map = {};
  rows_(SH.EMP).forEach(function (r) {
    var e = {
      id: id_(r['รหัส']), name: id_(r['ชื่อ-สกุล']), nick: id_(r['ชื่อเล่น']),
      branch: id_(r['สาขา']), branchName: id_(r['ชื่อสาขา']), team: id_(r['ทีม']),
      position: id_(r['ตำแหน่ง']), isRatee: /^y/i.test(id_(r['ถูกประเมิน'])),
      active: id_(r['สถานะ']) !== 'ลาออก',
      pin: id_(r['PIN']) ? ('0000' + id_(r['PIN'])).slice(-4) : ''
    };
    if (e.id) map[e.id] = e;
  });
  return map;
}

function pairs_(emps) {
  return rows_(SH.PAIR).map(function (r) {
    return { rater: id_(r['รหัสผู้ประเมิน']), ratee: id_(r['รหัสผู้ถูกประเมิน']), rel: id_(r['ความสัมพันธ์']) };
  }).filter(function (p) {
    return emps[p.rater] && emps[p.ratee] && emps[p.rater].active && emps[p.ratee].active &&
      REL_ORDER.indexOf(p.rel) >= 0;
  });
}

function items_() {
  return rows_(SH.ITEM).map(function (r) {
    return { code: id_(r['รหัส']), cat: id_(r['หมวด']), text: id_(r['พฤติกรรมที่สังเกตได้']) };
  });
}

function categories_(items) {
  var cats = [];
  items.forEach(function (it) { if (cats.indexOf(it.cat) < 0) cats.push(it.cat); });
  return cats;
}

function responses_(items) {
  var sh = sheet_(SH.RESP);
  var v = sh.getDataRange().getValues();
  var h = v.shift().map(function (x) { return String(x).trim(); });
  return v.filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    var o = { rater: id_(r[h.indexOf('รหัสผู้ประเมิน')]), ratee: id_(r[h.indexOf('รหัสผู้ถูกประเมิน')]),
              rel: id_(r[h.indexOf('ความสัมพันธ์')]), time: r[h.indexOf('เวลาส่ง')], scores: {} };
    items.forEach(function (it) {
      var x = r[h.indexOf(it.code)];
      o.scores[it.code] = (x === '' || id_(x) === 'N/A') ? null : Number(x);
    });
    ['Q1', 'Q2', 'Q3'].forEach(function (q) { o[q] = id_(r[h.indexOf(q)]); });
    return o;
  });
}

function auth_(id, pin) {
  var s = settings_();
  var emps = employees_();
  var me = emps[id_(id)];
  if (!me || !me.active) throw new Error('ไม่พบรหัสพนักงานนี้ในรายชื่อผู้ประเมิน');
  if (s.usePin && me.pin !== id_(pin)) throw new Error('PIN ไม่ถูกต้อง');
  return { s: s, emps: emps, me: me };
}

function pub_(e) {
  return { id: e.id, name: e.name, nick: e.nick, position: e.position, branchName: e.branchName, team: e.team };
}

/* ---------------- Public API: ผู้ประเมิน ---------------- */

function apiInfo() {
  var s = settings_();
  var scale = rows_(SH.SCALE).map(function (r) {
    return { score: id_(r['คะแนน']), level: id_(r['ระดับ']), desc: id_(r['ความหมาย']) };
  });
  var open = rows_(SH.OPEN).map(function (r) { return { code: id_(r['รหัส']), text: id_(r['คำถาม']) }; });
  return { title: s.title, openDate: s.openDate, closeDate: s.closeDate, isOpen: isOpen_(s),
           testMode: s.testMode, usePin: s.usePin, items: items_(), scale: scale, openQ: open };
}

function apiLogin(id, pin) {
  var a = auth_(id, pin);
  var items = items_();
  var done = {};
  responses_(items).forEach(function (r) { if (r.rater === a.me.id) done[r.ratee] = true; });
  var tasks = pairs_(a.emps).filter(function (p) { return p.rater === a.me.id; }).map(function (p) {
    var o = pub_(a.emps[p.ratee]);
    o.rel = p.rel; o.done = !!done[p.ratee];
    return o;
  });
  tasks.sort(function (x, y) { return REL_ORDER.indexOf(x.rel) - REL_ORDER.indexOf(y.rel); });
  return { me: pub_(a.me), tasks: tasks, isOpen: isOpen_(a.s) };
}

function apiMyAnswer(id, pin, rateeId) {
  var a = auth_(id, pin);
  var r = responses_(items_()).filter(function (x) { return x.rater === a.me.id && x.ratee === id_(rateeId); })[0];
  if (!r) return null;
  return { scores: r.scores, Q1: r.Q1, Q2: r.Q2, Q3: r.Q3 };
}

function apiSubmit(id, pin, data) {
  var a = auth_(id, pin);
  if (!isOpen_(a.s)) throw new Error('ขณะนี้ไม่อยู่ในช่วงเวลาการประเมิน');
  var rateeId = id_(data.rateeId);
  var pair = pairs_(a.emps).filter(function (p) { return p.rater === a.me.id && p.ratee === rateeId; })[0];
  if (!pair) throw new Error('คุณไม่มีรายชื่อผู้ถูกประเมินคนนี้');
  var items = items_();
  var row = [new Date(), a.me.id, rateeId, pair.rel];
  items.forEach(function (it) {
    var v = data.scores ? data.scores[it.code] : null;
    if (v === 'NA') { row.push('N/A'); return; }
    v = Number(v);
    if (!(v >= 1 && v <= 5)) throw new Error('กรุณาให้คะแนนข้อ ' + it.code);
    row.push(v);
  });
  ['Q1', 'Q2', 'Q3'].forEach(function (q) { row.push(String(data[q] || '').slice(0, 2000)); });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(SH.RESP);
    var v = sh.getDataRange().getValues();
    var h = v[0].map(function (x) { return String(x).trim(); });
    var ci = h.indexOf('รหัสผู้ประเมิน'), ti = h.indexOf('รหัสผู้ถูกประเมิน');
    var at = -1;
    for (var i = 1; i < v.length; i++) {
      if (id_(v[i][ci]) === a.me.id && id_(v[i][ti]) === rateeId) { at = i + 1; break; }
    }
    // เก็บรหัสเป็นข้อความ ป้องกัน Sheets แปลงเป็นตัวเลข
    row[1] = "'" + row[1]; row[2] = "'" + row[2];
    if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return apiLogin(id, pin);
}

/* ---------------- Scoring ---------------- */

function mean_(arr) {
  var a = arr.filter(function (x) { return x != null && !isNaN(x); });
  if (!a.length) return null;
  return a.reduce(function (s, x) { return s + x; }, 0) / a.length;
}

function groupStats_(list, items, cats) {
  if (!list.length) return null;
  var item = {}, cat = {};
  items.forEach(function (it) { item[it.code] = mean_(list.map(function (r) { return r.scores[it.code]; })); });
  cats.forEach(function (c) {
    cat[c] = mean_(items.filter(function (it) { return it.cat === c; }).map(function (it) { return item[it.code]; }));
  });
  return { n: list.length, item: item, cat: cat, overall: mean_(cats.map(function (c) { return cat[c]; })) };
}

function verdict_(overall, minCat) {
  if (overall == null) return 'ยังไม่มีข้อมูล';
  if (overall >= 4 && (minCat == null || minCat >= 3)) return 'พร้อมเติบโต';
  if (overall >= 3.5) return 'พัฒนาต่อ';
  return 'ยังไม่พร้อม';
}

/** คำนวณผลของผู้ถูกประเมิน 1 คน */
function scoreRatee_(rateeId, resp, s, items, cats) {
  var mine = resp.filter(function (r) { return r.ratee === rateeId; });
  var by = { 'ตนเอง': [], 'หัวหน้า': [], 'เพื่อน': [], 'ลูกน้อง': [] };
  mine.forEach(function (r) { if (by[r.rel]) by[r.rel].push(r); });

  // กลุ่มที่ใช้คิดคะแนน
  var groups = [];
  var subFew = by['ลูกน้อง'].length > 0 && by['ลูกน้อง'].length < s.minSub;
  groups.push({ key: 'หัวหน้า', label: 'หัวหน้า', w: s.w['หัวหน้า'], list: by['หัวหน้า'] });
  if (subFew) {
    groups.push({ key: 'เพื่อน', label: 'เพื่อน + ลูกน้อง', w: s.w['เพื่อน'] + s.w['ลูกน้อง'],
                  list: by['เพื่อน'].concat(by['ลูกน้อง']), merged: true });
  } else {
    groups.push({ key: 'เพื่อน', label: 'เพื่อนร่วมงาน', w: s.w['เพื่อน'], list: by['เพื่อน'] });
    groups.push({ key: 'ลูกน้อง', label: 'ลูกน้อง', w: s.w['ลูกน้อง'], list: by['ลูกน้อง'] });
  }
  groups.forEach(function (g) { g.stats = groupStats_(g.list, items, cats); g.n = g.list.length; });
  var used = groups.filter(function (g) { return g.stats && g.w > 0; });

  function weighted(pick) {
    var sw = 0, sx = 0;
    used.forEach(function (g) { var x = pick(g.stats); if (x != null) { sw += g.w; sx += g.w * x; } });
    return sw ? sx / sw : null;
  }
  var wItem = {}, wCat = {};
  items.forEach(function (it) { wItem[it.code] = weighted(function (st) { return st.item[it.code]; }); });
  cats.forEach(function (c) { wCat[c] = weighted(function (st) { return st.cat[c]; }); });
  var overall = mean_(cats.map(function (c) { return wCat[c]; }));
  var catVals = cats.map(function (c) { return wCat[c]; }).filter(function (x) { return x != null; });
  var minCat = catVals.length ? Math.min.apply(null, catVals) : null;
  var sw = used.reduce(function (t, g) { return t + g.w; }, 0);

  return {
    self: groupStats_(by['ตนเอง'], items, cats),
    groups: groups.map(function (g) {
      return { key: g.key, label: g.label, n: g.n, weight: g.w, merged: !!g.merged,
               effWeight: (g.stats && sw) ? g.w / sw : 0, stats: g.stats };
    }),
    item: wItem, cat: wCat, overall: overall, minCat: minCat,
    verdict: verdict_(overall, minCat),
    capped: overall != null && overall >= 4 && minCat != null && minCat < 3,
    responses: mine,
    subFew: subFew
  };
}

/* ---------------- Public API: HR ---------------- */

function checkHr_(pw) {
  var s = settings_();
  if (!s.hrPassword || id_(pw) !== s.hrPassword) throw new Error('รหัสผ่าน HR ไม่ถูกต้อง');
  return s;
}

function apiAdmin(pw) {
  var s = checkHr_(pw);
  var emps = employees_(), pairs = pairs_(emps), items = items_(), cats = categories_(items);
  var resp = responses_(items);
  var key = function (a, b) { return a + '|' + b; };
  var pairSet = {};
  pairs.forEach(function (p) { pairSet[key(p.rater, p.ratee)] = p; });
  resp = resp.filter(function (r) { return pairSet[key(r.rater, r.ratee)]; });
  var doneSet = {};
  resp.forEach(function (r) { doneSet[key(r.rater, r.ratee)] = true; });

  var ratees = Object.keys(emps).map(function (k) { return emps[k]; })
    .filter(function (e) { return e.isRatee && e.active; })
    .map(function (e) {
      var sc = scoreRatee_(e.id, resp, s, items, cats);
      var assigned = pairs.filter(function (p) { return p.ratee === e.id; });
      var byRel = {};
      REL_ORDER.forEach(function (r) { byRel[r] = { total: 0, done: 0 }; });
      assigned.forEach(function (p) { byRel[p.rel].total++; if (doneSet[key(p.rater, p.ratee)]) byRel[p.rel].done++; });
      var o = pub_(e);
      o.total = assigned.length;
      o.done = assigned.filter(function (p) { return doneSet[key(p.rater, p.ratee)]; }).length;
      o.byRel = byRel; o.overall = sc.overall; o.minCat = sc.minCat; o.verdict = sc.verdict;
      o.selfOverall = sc.self ? sc.self.overall : null; o.cat = sc.cat;
      return o;
    });

  var raterMap = {};
  pairs.forEach(function (p) {
    var r = raterMap[p.rater] || (raterMap[p.rater] = { e: emps[p.rater], total: 0, done: 0, pending: [] });
    r.total++;
    if (doneSet[key(p.rater, p.ratee)]) r.done++;
    else r.pending.push(emps[p.ratee].name + (emps[p.ratee].nick ? ' (' + emps[p.ratee].nick + ')' : '') + ' – ' + p.rel);
  });
  var raters = Object.keys(raterMap).map(function (k) {
    var r = raterMap[k], o = pub_(r.e);
    o.total = r.total; o.done = r.done; o.pending = r.pending;
    return o;
  });

  return { title: s.title, openDate: s.openDate, closeDate: s.closeDate, testMode: s.testMode,
           cats: cats, ratees: ratees, raters: raters,
           totalForms: pairs.length, doneForms: resp.length,
           updated: nowTh_() };
}

function apiReport(pw, rateeId) {
  var s = checkHr_(pw);
  var emps = employees_(), items = items_(), cats = categories_(items);
  var pairs = pairs_(emps);
  var ok = {};
  pairs.forEach(function (p) { ok[p.rater + '|' + p.ratee] = true; });
  var resp = responses_(items).filter(function (r) { return ok[r.rater + '|' + r.ratee]; });
  var e = emps[id_(rateeId)];
  if (!e) throw new Error('ไม่พบผู้ถูกประเมิน');
  var sc = scoreRatee_(e.id, resp, s, items, cats);

  // คำถามปลายเปิด แยกตามกลุ่ม ไม่แสดงชื่อผู้ตอบ
  var labelOf = function (rel) {
    if (rel === 'ตนเอง') return 'ตนเอง';
    if (sc.subFew && (rel === 'เพื่อน' || rel === 'ลูกน้อง')) return 'เพื่อน + ลูกน้อง';
    return { 'หัวหน้า': 'หัวหน้า', 'เพื่อน': 'เพื่อนร่วมงาน', 'ลูกน้อง': 'ลูกน้อง' }[rel];
  };
  var comments = { Q1: [], Q2: [], Q3: [] };
  sc.responses.slice().sort(function (a, b) {
    return REL_ORDER.indexOf(a.rel) - REL_ORDER.indexOf(b.rel);
  }).forEach(function (r) {
    ['Q1', 'Q2', 'Q3'].forEach(function (q) {
      if (r[q]) comments[q].push({ group: labelOf(r.rel), text: r[q] });
    });
  });
  var assigned = {};
  pairs.filter(function (p) { return p.ratee === e.id; }).forEach(function (p) {
    assigned[p.rel] = (assigned[p.rel] || 0) + 1;
  });
  delete sc.responses;
  return { person: pub_(e), items: items, cats: cats, score: sc, comments: comments, assigned: assigned,
           weights: s.w, minSub: s.minSub,
           generated: nowTh_() };
}
