/*
 * หน้าทดสอบระยะ 0
 * ห้ามใส่ path โฟลเดอร์ข้อมูล ชื่อคน หรืออีเมลจริงในไฟล์นี้ (repo จะเป็นสาธารณะ)
 * ผู้ใช้เลือกโฟลเดอร์เองผ่านหน้าต่างของ Edge ทุกครั้ง
 */
(function () {
  'use strict';

  // ---------- ข้อมูลสมมติสำหรับ Word (ชื่อจากแบบเรียน ไม่ใช่บุคคลจริง) ----------
  var SAMPLE = {
    header: [
      'รายงานการประชุมสภาวิชาการ มหาวิทยาลัยหอการค้าไทย',
      'ครั้งที่ 5/2569',
      'วันพฤหัสบดีที่ 15 ตุลาคม พ.ศ. 2569 เวลา 9:30-12:00 น.',
      'รูปแบบออนไลน์ผ่านระบบ Zoom Meetings'
    ],
    advisors: [
      { name: 'ศาสตราจารย์ ดร.สมชาย ใจดี', role: 'ที่ปรึกษาสภาวิชาการ' },
      { name: 'รองศาสตราจารย์ ดร.สมหญิง รักเรียน', role: 'ที่ปรึกษาสภาวิชาการ' }
    ],
    members: [
      { name: 'ศาสตราจารย์ (วิจัย) ดร.มานะ ขยันยิ่ง', role: 'ประธานกรรมการ' },
      { name: 'รองศาสตราจารย์ ดร.ปิติ ยินดี', role: 'กรรมการ' },
      { name: 'ผู้ช่วยศาสตราจารย์ ดร.ชูใจ ใฝ่ดี', role: 'กรรมการ' },
      { name: 'อาจารย์ ดร.วีระ กล้าหาญ', role: 'กรรมการและเลขานุการ', note: '(รองอธิการบดีฝ่ายวิชาการ)' },
      { name: 'นางสาวมานี มีตา', role: 'ผู้ช่วยเลขานุการ', note: '(ผู้อำนวยการสำนักวิชาการ)' }
    ],
    absent: [
      { name: 'ศาสตราจารย์ ดร.ดวงใจ ปั้นน้ำ', role: 'ที่ปรึกษาสภาวิชาการ' }
    ],
    attendees: [
      { name: 'ผู้ช่วยศาสตราจารย์ ดร.ใฝ่ฝัน ตั้งใจ', note: 'คณบดี คณะตัวอย่างศาสตร์' }
    ],
    startTime: '9.30',
    chair: 'ศาสตราจารย์ (วิจัย) ดร.มานะ ขยันยิ่ง ประธานกรรมการ'
  };

  // ---------- โครงสร้างโฟลเดอร์จริงที่จะใช้ ----------
  var MEETING_PATH = ['สภาวิชาการ', '2569-05'];
  var ATTACH_DIR = 'เอกสารแนบ';
  var GENERATED_DIR = 'เอกสารที่สร้าง';
  var ATTACH_FILE = '4.1.1.pdf';
  var REPORT_FILE = 'รายงาน-ร่าง-v1.docx';
  var THAI_PROBE = 'ทดสอบภาษาไทย น้ำ ผู้ใหญ่ ปั๊ม กี่ ฤๅษี ฯลฯ ๑๒๓';

  var state = { dir: null, lastWrite: null };
  var results = [];

  // ---------- UI helpers ----------
  function $(id) { return document.getElementById(id); }

  function log(kind, msg) {
    var icon = { pass: '✅', fail: '❌', info: 'ℹ️', warn: '⚠️' }[kind] || '•';
    var time = new Date().toLocaleTimeString('th-TH');
    results.push('[' + time + '] ' + icon + ' ' + msg);
    var li = document.createElement('li');
    li.className = kind;
    li.textContent = icon + ' ' + msg;
    $('log').appendChild(li);
    li.scrollIntoView({ block: 'nearest' });
  }

  function errText(e) { return (e && e.name ? e.name + ': ' : '') + (e && e.message ? e.message : String(e)); }

  function setDirUI() {
    $('dir-name').textContent = state.dir ? state.dir.name : '(ยังไม่ได้เลือก)';
    document.querySelectorAll('[data-needs-dir]').forEach(function (b) { b.disabled = !state.dir; });
  }

  // ---------- IndexedDB: เก็บตัวชี้โฟลเดอร์ ----------
  function idb() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open('oaa-meeting-phase0', 1);
      req.onupgradeneeded = function () { req.result.createObjectStore('kv'); };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function idbOp(mode, fn) {
    return idb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction('kv', mode);
        var req = fn(tx.objectStore('kv'));
        tx.oncomplete = function () { resolve(req && req.result); };
        tx.onerror = function () { reject(tx.error); };
        tx.onabort = function () { reject(tx.error); };
      });
    });
  }
  function idbGet(key) { return idbOp('readonly', function (s) { return s.get(key); }); }
  function idbSet(key, val) { return idbOp('readwrite', function (s) { return s.put(val, key); }); }
  function idbDel(key) { return idbOp('readwrite', function (s) { return s.delete(key); }); }

  // ---------- File helpers ----------
  function writeFile(dirHandle, name, data) {
    return dirHandle.getFileHandle(name, { create: true }).then(function (fh) {
      return fh.createWritable().then(function (w) {
        return w.write(data).then(function () { return w.close(); });
      }).then(function () { return fh.getFile(); });
    });
  }
  function readText(dirHandle, name) {
    return dirHandle.getFileHandle(name).then(function (fh) { return fh.getFile(); })
      .then(function (f) { return f.text(); });
  }
  function getDirPath(root, parts) {
    return parts.reduce(function (p, name) {
      return p.then(function (d) { return d.getDirectoryHandle(name, { create: true }); });
    }, Promise.resolve(root));
  }
  async function listTree(dirHandle, prefix, depth, out) {
    for await (var entry of dirHandle.values()) {
      var p = prefix + entry.name;
      out.push({ path: p, kind: entry.kind });
      if (entry.kind === 'directory' && depth > 0) await listTree(entry, p + '\\', depth - 1, out);
    }
    return out;
  }

  // PDF 1 หน้าแบบเล็กที่สุด (ข้อความอังกฤษ) ใช้แทนเอกสารแนบ
  function makeTestPdf() {
    var objs = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
      '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    ];
    var stream = 'BT /F1 20 Tf 72 760 Td (Phase 0 test attachment: agenda 4.1.1) Tj ET';
    objs.push('<< /Length ' + stream.length + ' >>\nstream\n' + stream + '\nendstream');
    var pdf = '%PDF-1.4\n';
    var offsets = [];
    objs.forEach(function (o, i) {
      offsets.push(pdf.length);
      pdf += (i + 1) + ' 0 obj\n' + o + '\nendobj\n';
    });
    var xref = pdf.length;
    pdf += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n';
    offsets.forEach(function (off) { pdf += ('0000000000' + off).slice(-10) + ' 00000 n \n'; });
    pdf += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';
    return new TextEncoder().encode(pdf);
  }

  // ---------- ข้อ 1: เปิดโฟลเดอร์ ----------
  async function pickFolder() {
    try {
      var handle = await window.showDirectoryPicker({ id: 'oaa-data', mode: 'readwrite' });
      state.dir = handle;
      log('pass', '1.1 เลือกโฟลเดอร์ได้: "' + handle.name + '" และได้สิทธิ์อ่าน/เขียน');
      try {
        await idbSet('dataDir', handle);
        log('pass', '1.3a บันทึกตัวชี้โฟลเดอร์ลง IndexedDB ได้');
      } catch (e) {
        log('fail', '1.3a บันทึกตัวชี้โฟลเดอร์ลง IndexedDB ไม่ได้ (' + errText(e) + ') — เปิดหน้าใหม่ต้องเลือกโฟลเดอร์ซ้ำ');
      }
    } catch (e) {
      if (e.name === 'AbortError') log('info', 'ยกเลิกการเลือกโฟลเดอร์');
      else log('fail', '1.1 เลือกโฟลเดอร์ไม่ได้: ' + errText(e));
    }
    setDirUI();
    showStored();
  }

  async function reconnect() {
    try {
      var handle = await idbGet('dataDir');
      if (!handle) { log('fail', '1.3 ไม่พบโฟลเดอร์ที่จำไว้'); return; }
      var perm = await handle.requestPermission({ mode: 'readwrite' });
      if (perm === 'granted') {
        state.dir = handle;
        log('pass', '1.3 เชื่อมต่อโฟลเดอร์เดิม "' + handle.name + '" ได้โดยไม่ต้องเลือกใหม่');
        $('reconnect').hidden = true;
      } else {
        log('fail', '1.3 ไม่ได้รับอนุญาต (ผล: ' + perm + ')');
      }
    } catch (e) {
      log('fail', '1.3 เชื่อมต่อโฟลเดอร์เดิมไม่ได้: ' + errText(e));
    }
    setDirUI();
  }

  async function showStored() {
    var handle = null;
    try { handle = await idbGet('dataDir'); } catch (e) {
      log('warn', 'อ่าน IndexedDB ไม่ได้: ' + errText(e));
    }
    if (!handle) { $('stored').textContent = '(ไม่มี)'; $('reconnect').hidden = true; return; }
    $('stored').textContent = handle.name;
    if (state.dir) { $('reconnect').hidden = true; return; }
    var perm = await handle.queryPermission({ mode: 'readwrite' });
    if (perm === 'granted') {
      state.dir = handle;
      log('pass', '1.3 จำโฟลเดอร์ "' + handle.name + '" ได้ และได้สิทธิ์ทันทีโดยไม่ต้องกด (Edge จำการอนุญาตไว้)');
      setDirUI();
    } else {
      $('reconnect').hidden = false;
      log('info', 'พบโฟลเดอร์ที่จำไว้: "' + handle.name + '" (สถานะสิทธิ์: ' + perm + ') — กดปุ่ม "เชื่อมต่อโฟลเดอร์เดิม"');
    }
  }

  async function forget() {
    if (!confirm('ล้างโฟลเดอร์ที่หน้านี้จำไว้? (ไม่ลบไฟล์ใด ๆ ในโฟลเดอร์)')) return;
    await idbDel('dataDir');
    state.dir = null;
    log('info', 'ล้างโฟลเดอร์ที่จำไว้แล้ว');
    setDirUI();
    showStored();
  }

  async function writeReadTest() {
    try {
      var payload = {
        _test: 'phase0 — ลบได้',
        text: THAI_PROBE,
        written_at: new Date().toISOString(),
        origin: location.origin || location.protocol,
        browser: navigator.userAgent
      };
      var json = JSON.stringify(payload, null, 2);
      var file = await writeFile(state.dir, 'test.json', json);
      state.lastWrite = file.lastModified;
      await idbSet('lastWrite', file.lastModified);
      log('pass', '1.2 เขียน test.json ได้ (' + file.size + ' ไบต์)');
      var back = await readText(state.dir, 'test.json');
      if (back === json && JSON.parse(back).text === THAI_PROBE) {
        log('pass', '1.2 อ่านกลับมาตรงกันทุกไบต์ ภาษาไทยไม่เพี้ยน: "' + JSON.parse(back).text + '"');
      } else {
        log('fail', '1.2 อ่านกลับมาไม่ตรงกับที่เขียน');
      }
    } catch (e) {
      log('fail', '1.2 เขียน/อ่าน test.json ไม่ได้: ' + errText(e));
    }
  }

  async function readOnlyTest() {
    try {
      var back = await readText(state.dir, 'test.json');
      var obj = JSON.parse(back);
      log('pass', '1.3 อ่าน test.json หลังเปิดหน้าใหม่ได้ (เขียนไว้เมื่อ ' +
        new Date(obj.written_at).toLocaleString('th-TH') + '): "' + obj.text + '"');
    } catch (e) {
      log('fail', '1.3 อ่าน test.json ไม่ได้: ' + errText(e));
    }
  }

  async function changeCheck() {
    try {
      var last = state.lastWrite || await idbGet('lastWrite');
      if (!last) { log('warn', '1.4 ยังไม่เคยเขียน test.json จากหน้านี้ — กดปุ่มเขียนก่อน'); return; }
      var f = await (await state.dir.getFileHandle('test.json')).getFile();
      if (f.lastModified !== last) {
        log('pass', '1.4 ตรวจพบว่า test.json ถูกแก้จากที่อื่นหลังหน้านี้เขียน (เวลาแก้ล่าสุด ' +
          new Date(f.lastModified).toLocaleString('th-TH') + ')');
      } else {
        log('info', '1.4 test.json ยังไม่ถูกแก้จากที่อื่น (ถ้าเพิ่งแก้ใน Notepad แล้วยังขึ้นข้อความนี้ = ตรวจไม่เจอ)');
      }
    } catch (e) {
      log('fail', '1.4 ตรวจไม่ได้: ' + errText(e));
    }
  }

  async function structureTest() {
    try {
      var meeting = await getDirPath(state.dir, MEETING_PATH);
      var attach = await meeting.getDirectoryHandle(ATTACH_DIR, { create: true });
      await meeting.getDirectoryHandle(GENERATED_DIR, { create: true });
      var pdf = makeTestPdf();
      var f = await writeFile(attach, ATTACH_FILE, pdf);
      log('pass', '1.5 เขียน ' + MEETING_PATH.join('\\') + '\\' + ATTACH_DIR + '\\' + ATTACH_FILE + ' ได้ (' + f.size + ' ไบต์)');

      var agenda = {
        _test: 'phase0 — ลบได้',
        items: [{
          no: '4.1.1',
          title: 'ขออนุมัติหลักสูตรวิทยาศาสตรบัณฑิต สาขาวิชาตัวอย่างเทคโนโลยี (หลักสูตรปรับปรุง พ.ศ. 2570) ซึ่งเป็นชื่อเรื่องยาวที่ไม่ควรเป็นชื่อไฟล์',
          attachments: [{ file: ATTACH_DIR + '/' + ATTACH_FILE, title: 'เอกสารประกอบการพิจารณาหลักสูตร (ฉบับเต็ม)' }]
        }]
      };
      var json = JSON.stringify(agenda, null, 2);
      await writeFile(meeting, 'agenda.json', json);
      var back = JSON.parse(await readText(meeting, 'agenda.json'));
      if (back.items[0].title === agenda.items[0].title) {
        log('pass', '1.5 agenda.json เก็บชื่อเรื่องเต็มภาษาไทยและอ่านกลับได้ตรงกัน');
      } else {
        log('fail', '1.5 agenda.json อ่านกลับไม่ตรง');
      }
      var head = new TextDecoder().decode((await (await (await attach.getFileHandle(ATTACH_FILE)).getFile()).slice(0, 8).arrayBuffer()));
      log(head.indexOf('%PDF') === 0 ? 'pass' : 'fail', '1.5 อ่าน 4.1.1.pdf กลับได้ (ขึ้นต้นด้วย ' + head.trim() + ')');

      var rel = [state.dir.name].concat(MEETING_PATH, [GENERATED_DIR, REPORT_FILE]).join('\\');
      log('info', '1.5 path ยาวสุดที่ระบบจะใช้ (นับจากชื่อโฟลเดอร์ข้อมูล): ' + rel + ' = ' + rel.length +
        ' ตัวอักษร — ความยาวเต็มดูด้วยคำสั่ง PowerShell ใน checklist');
    } catch (e) {
      log('fail', '1.5 สร้างโครงสร้างโฟลเดอร์ไม่ได้: ' + errText(e));
    }
  }

  async function scanTemp() {
    try {
      var all = await listTree(state.dir, '', 4, []);
      var swap = all.filter(function (x) { return /\.crswap$/i.test(x.path); });
      var conflict = all.filter(function (x) { return /-[^\\-]+\.(json|docx|pdf)$/i.test(x.path) && !/รายงาน-ร่าง-v\d+\.docx$/.test(x.path); });
      log('info', '1.4 ไฟล์ในโฟลเดอร์ข้อมูล: ' + (all.map(function (x) { return x.path + (x.kind === 'directory' ? '\\' : ''); }).join(' | ') || '(ว่าง)'));
      if (swap.length) log('fail', '1.4 พบไฟล์ .crswap ค้าง: ' + swap.map(function (x) { return x.path; }).join(', '));
      else log('pass', '1.4 ไม่พบไฟล์ .crswap ค้าง');
      if (conflict.length) log('warn', '1.4 อาจเป็นไฟล์ชนกันจาก OneDrive (ชื่อมี -ชื่อเครื่อง): ' + conflict.map(function (x) { return x.path; }).join(', '));
    } catch (e) {
      log('fail', '1.4 อ่านรายการไฟล์ไม่ได้: ' + errText(e));
    }
  }

  // ---------- ข้อ 2: Word ----------
  function buildDocx() {
    return window.DocxPhase0.buildReportHeader(SAMPLE, { thaiTest: $('thai-test').checked });
  }

  function downloadDocx() {
    try {
      var blob = buildDocx();
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'phase0-' + REPORT_FILE;
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
      log('pass', '2 สร้าง Word ได้ (' + blob.size + ' ไบต์) และสั่งดาวน์โหลด ' + a.download + ' แล้ว — เปิดใน Word เพื่อตรวจ');
    } catch (e) {
      log('fail', '2 สร้าง Word ไม่ได้: ' + errText(e));
    }
  }

  async function saveDocxToFolder() {
    try {
      var blob = buildDocx();
      var meeting = await getDirPath(state.dir, MEETING_PATH);
      var gen = await meeting.getDirectoryHandle(GENERATED_DIR, { create: true });
      var f = await writeFile(gen, REPORT_FILE, blob);
      log('pass', '2 บันทึก Word ลง ' + MEETING_PATH.join('\\') + '\\' + GENERATED_DIR + '\\' + REPORT_FILE + ' ได้ (' + f.size + ' ไบต์)');
    } catch (e) {
      var hint = /NoModificationAllowed|InvalidState|lock/i.test(errText(e)) ? ' — ถ้าเปิดไฟล์นี้ค้างใน Word ให้ปิดก่อน' : '';
      log('fail', '2 บันทึก Word ลงโฟลเดอร์ไม่ได้: ' + errText(e) + hint);
    }
  }

  // ---------- ทั่วไป ----------
  function copyResults() {
    var text = 'ผลทดสอบระยะ 0\n' + $('env').innerText + '\n\n' + results.join('\n');
    var done = function () { log('info', 'คัดลอกผลแล้ว วางในแชทหรือ checklist ได้เลย'); };
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(done, fallback);
    } else { fallback(); }
    function fallback() {
      var ta = document.createElement('textarea');
      ta.value = text; document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); ta.remove(); done();
    }
  }

  function showEnv() {
    var mode = location.protocol === 'file:' ? 'A: เปิดจากไฟล์ในเครื่อง (file://)' :
      (/^(localhost|127\.0\.0\.1|\[::1\])$/.test(location.hostname) ? 'B: localhost (จำลอง GitHub Pages)' : 'B: เว็บ (' + location.host + ')');
    var edge = (navigator.userAgent.match(/Edg\/([\d.]+)/) || [])[1];
    var chrome = (navigator.userAgent.match(/Chrome\/([\d.]+)/) || [])[1];
    $('env').innerText =
      'โหมด: ' + mode + '\n' +
      'เบราว์เซอร์: ' + (edge ? 'Microsoft Edge ' + edge : (chrome ? 'Chrome ' + chrome : navigator.userAgent)) + '\n' +
      'isSecureContext: ' + window.isSecureContext + '\n' +
      'showDirectoryPicker: ' + ('showDirectoryPicker' in window ? 'มี' : 'ไม่มี') + '\n' +
      'IndexedDB: ' + ('indexedDB' in window ? 'มี' : 'ไม่มี');
    if (!('showDirectoryPicker' in window)) {
      log('fail', 'เบราว์เซอร์นี้ไม่มี File System Access API (ต้องใช้ Edge หรือ Chrome บนคอมพิวเตอร์)');
    }
  }

  document.addEventListener('DOMContentLoaded', function () {
    $('pick').onclick = pickFolder;
    $('reconnect').onclick = reconnect;
    $('forget').onclick = forget;
    $('write-read').onclick = writeReadTest;
    $('read-only').onclick = readOnlyTest;
    $('change-check').onclick = changeCheck;
    $('scan').onclick = scanTemp;
    $('structure').onclick = structureTest;
    $('docx-download').onclick = downloadDocx;
    $('docx-save').onclick = saveDocxToFolder;
    $('copy').onclick = copyResults;
    $('font-name').textContent = window.DocxPhase0.FONT;
    showEnv();
    setDirUI();
    showStored();
  });
})();
