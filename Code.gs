/**
 * ============================================================================
 *  STAFF PERFORMANCE TRACKING SYSTEM  -  Code.gs  (Google Apps Script, V8)
 * ============================================================================
 *  Dipasang di Google Sheets hasil konversi STAFF_PERFORMANCE_TRACKING.xlsx
 *  (Extensions > Apps Script). Lihat README.md untuk langkah instalasi.
 *
 *  Modul:
 *    CONFIG     konfigurasi terpusat (nama sheet, kolom, format, default)
 *    Util       helper umum (tanggal, header, baris terakhir, dll.)
 *    Log        pencatatan error/info (Logger + console)
 *    Settings   membaca SETTINGS (daftar, konfigurasi, admin)
 *    Masters    membaca STAFF MASTER & PROJECT MASTER
 *    Db         akses TASK DATABASE (append-only, perluasan rumus otomatis)
 *    Ids        pembuatan ID unik (Task / Log / Evidence)
 *    Validate   validasi baris TASK INPUT
 *    Submit     SUBMIT TASK  (TASK INPUT -> TASK DATABASE)
 *    Progress   PROGRESS LOG -> update status TASK DATABASE
 *    Evidence   EVIDENCE LOG (+ penerima Google Form)
 *    Overdue    deteksi overdue & mendekati deadline
 *    Notify     email (penugasan, progres, reminder)
 *    Reminders  job harian
 *    Guard      mencegah edit baris yang sudah dikirim / sudah dicatat
 *    Triggers   pasang / lepas trigger
 *    SelfTest   pemeriksaan instalasi
 *
 *  PENTING: script TIDAK berjalan otomatis hanya karena workbook dibuat.
 *  Jalankan selfTest() dan ikuti README.md bagian "Verifikasi instalasi".
 * ============================================================================
 */

// ============================================================================
//  CONFIG  (satu-satunya tempat yang perlu diubah bila layout sheet berubah)
// ============================================================================
const CONFIG = {
  SHEETS: {
    INPUT: 'TASK INPUT', DB: 'TASK DATABASE', PROGRESS: 'PROGRESS LOG', EVIDENCE: 'EVIDENCE LOG',
    STAFF: 'STAFF MASTER', PROJECT: 'PROJECT MASTER', SETTINGS: 'SETTINGS',
    DASHBOARD: 'DASHBOARD', STAFF_PERF: 'STAFF PERFORMANCE', PROJECT_REPORT: 'PROJECT REPORT',
    SYSTEM_LOG: 'SYSTEM LOG'
  },

  // TASK INPUT: posisi kolom tetap (1-based)
  INPUT: {
    HEADER_ROW: 4, FIRST_ROW: 5, LAST_ROW: 304, NCOLS: 15,
    COL: { NO: 1, TASK_ID: 2, INPUT_DATE: 3, CREATED_BY: 4, PROJECT: 5, PIC: 6, PARENT: 7, TITLE: 8, DESC: 9,
           PRIORITY: 10, START: 11, DEADLINE_DATE: 12, NOTES: 13, STATUS: 14, MESSAGE: 15 },
    SUBMITTED: 'SUBMITTED'
  },
  // Deadline = akhir hari pada tanggal deadline (tidak ada input jam)
  DEADLINE_END_OF_DAY: { h: 23, m: 59 },

  // Nama header (dicari berdasarkan teks, jadi urutan kolom boleh bergeser)
  DB_H: {
    TASK_ID: 'Task ID', CREATED_AT: 'Created At', CREATED_BY: 'Created By', PROJECT_ID: 'Project ID',
    PROJECT_NAME: 'Project Name', ASSIGNED_TO: 'Assigned To', PARENT: 'Parent Task ID', TITLE: 'Task Title',
    DESC: 'Task Description', PRIORITY: 'Priority', START: 'Start Date', DEADLINE: 'Deadline',
    STATUS: 'Current Status', PROGRESS: 'Progress Percentage', COMPLETION: 'Completion Date',
    DELAY_REASON: 'Delay Reason', RESULT_NOTES: 'Result Notes', LAST_UPDATED: 'Last Updated',
    EVIDENCE: 'Evidence Link', DELAY_CATEGORY: 'Delay Category', QUALITY: 'Quality Rating',
    SUBMITTED_BY: 'Submitted By (Email)',
    AUTO_FIRST: 'Effective Status'            // kolom pertama yang berisi rumus otomatis
  },
  PL_H: { LOG_ID: 'Log ID', TS: 'Timestamp', TASK_ID: 'Task ID', UPDATED_BY: 'Updated By', PREV: 'Previous Status',
          NEW: 'New Status', PROGRESS: 'Progress Percentage', NOTES: 'Update Notes', DELAY: 'Delay Reason',
          REVISED: 'Revised Estimated Completion', NOTE: 'Processing Note' },
  EV_H: { ID: 'Evidence ID', TASK_ID: 'Task ID', DATE: 'Upload Date', BY: 'Uploaded By',
          DESC: 'Evidence Description', URL: 'Google Drive URL', NOTE: 'Processing Note' },

  // SETTINGS: posisi daftar & konfigurasi
  SETTINGS_LAYOUT: {
    LIST_FIRST: 5, LIST_LAST: 14,
    COL: { PRIORITY: 1, STATUS: 2, ADMIN_NAME: 3, ADMIN_EMAIL: 4, DELAY_CAT: 5, DELAY_COUNTS: 6,
           Q_LABEL: 7, Q_SCORE: 8, PERIOD: 9, ACTIVE: 10, YESNO: 11, ADMIN_CODE: 12 },
    CFG_KEY_COL: 13, CFG_VAL_COL: 14, CFG_FIRST: 5, CFG_LAST: 80,
    WEIGHT_FIRST: 24, WEB_BLOCK_ROW: 31
  },
  // Nilai default bila key tidak ada di SETTINGS
  DEFAULTS: {
    SYSTEM_NAME: 'Staff Performance Tracking', TASK_ID_PREFIX: 'TSK', TASK_ID_DIGITS: '5', DUE_SOON_DAYS: '3',
    UPDATE_STALE_DAYS: '3', ALLOW_PAST_DEADLINE: 'NO',
    ENFORCE_CREATED_BY_MATCH: 'NO', REQUIRE_DELAY_REASON: 'YES', ALLOW_REOPEN_CLOSED: 'NO', EMAIL_ENABLED: 'YES',
    EMAIL_NOTIFY_ON_SUBMIT: 'YES', EMAIL_NOTIFY_ON_PROGRESS: 'YES', EMAIL_CC_ADMINS: 'YES',
    REMINDER_ENABLED: 'YES', REMINDER_HOUR: '8', REMINDER_ADMIN_DIGEST: 'YES',
    WEB_ENABLED: 'YES', WEB_APP_URL: '', EVIDENCE_PUBLIC_LINK: 'YES', REQUIRE_EVIDENCE_ON_COMPLETE: 'NO', MAX_PHOTOS_PER_TASK: '20', REQUIRE_BEFORE_AFTER: 'NO'
  },

  STATUS: { NEW: 'NEW', IN_PROGRESS: 'IN PROGRESS', PENDING: 'PENDING', COMPLETED: 'COMPLETED', CANCELLED: 'CANCELLED' },
  LOG_PREFIX: 'LOG', EVD_PREFIX: 'EVD',
  PROP: { TASK_SEQ: 'TASK_SEQ', LOG_SEQ: 'LOG_SEQ', EVD_SEQ: 'EVD_SEQ', LAST_DAILY_RUN: 'LAST_DAILY_RUN',
          FORM_TASK: 'FORM_TASK_ID', FORM_PROGRESS: 'FORM_PROGRESS_ID', EVIDENCE_FOLDER: 'EVIDENCE_FOLDER_ID',
          PATCH_VERSION: 'PATCH_VERSION' },

  FORMATS: { DATETIME: 'dd/mm/yyyy hh:mm', DATE: 'dd/mm/yyyy', PERCENT: '0%' },

  // Judul pertanyaan Google Form (harus sama persis dengan yang dibuat script / dibuat manual)
  FORMS: {
    // Form tugas HP (dibuat otomatis oleh Mobile Forms > Setup). Created By diambil dari email login admin.
    TASK: { PROJECT: 'Project', PIC: 'Assigned To (PIC)', TITLE: 'Task Title', DESC: 'Task Description', PRIORITY: 'Priority',
            START: 'Start Date (opsional)', DEADLINE: 'Deadline Date', PARENT: 'Parent Task ID (opsional)', NOTES: 'Notes (opsional)' },
    // Form progres HP (dibuat otomatis). Updated By diambil dari email login staf/admin.
    PROGRESS: { TASK: 'Tugas', STATUS: 'New Status', PROGRESS: 'Progress Percentage', NOTES: 'Update Notes',
                DELAY: 'Delay Reason', REVISED: 'Revised Estimated Completion (opsional)' },
    // Form evidence (dibuat MANUAL karena script tidak bisa membuat pertanyaan upload file)
    EVIDENCE: { TASK_ID: 'Task ID', UPLOADED_BY: 'Uploaded By', DESCRIPTION: 'Evidence Description', FILES: 'Foto / File' },
    TITLE_TASK: 'Input Tugas Baru', TITLE_PROGRESS: 'Update Progres Tugas',
    KEEP_PROGRESS: '(tidak berubah)', MAX_TASK_CHOICES: 400
  },

  LIMITS: { TITLE: 200, TEXT: 2000 },
  DB_MAX_ROWS: 10000,          // harus sama dengan kapasitas rumus di workbook
  FORMULA_EXTEND_CHUNK: 500,   // jumlah baris rumus yang ditambahkan sekali perluasan
  LOCK_TIMEOUT_MS: 30000,
  MAX_ERRORS_IN_ALERT: 15,
  PROTECT_TAG: 'TSK-SYS'
};

// ============================================================================
//  Util
// ============================================================================
const Util = {
  ss() { return SpreadsheetApp.getActiveSpreadsheet(); },

  sheet(name) {
    const sh = Util.ss().getSheetByName(name);
    if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan. Jangan ubah nama sheet (lihat CONFIG.SHEETS).');
    return sh;
  },

  tz() { return Util.ss().getSpreadsheetTimeZone(); },

  str(v) { return v === null || v === undefined ? '' : String(v).trim(); },
  up(v) { return Util.str(v).toUpperCase(); },
  low(v) { return Util.str(v).toLowerCase(); },
  norm(v) { return Util.str(v).toLowerCase().replace(/\s+/g, ' '); },
  isYes(v) { return Util.up(v) === 'YES'; },
  isDate(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); },
  pad(n, w) { let s = String(n); while (s.length < w) s = '0' + s; return s; },
  esc(s) {
    return Util.str(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  },

  fmt(d, pattern) { return Util.isDate(d) ? Utilities.formatDate(d, Util.tz(), pattern || 'dd/MM/yyyy HH:mm') : ''; },

  startOfDay(d) {
    const p = Utilities.formatDate(d, Util.tz(), 'yyyy-MM-dd').split('-');
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 0, 0, 0, 0);
  },

  /** Gabungkan tanggal (Date dari sel) + jam/menit -> Date di zona waktu script. */
  combine(dateCell, h, m) {
    const p = Utilities.formatDate(dateCell, Util.tz(), 'yyyy-MM-dd').split('-');
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), h, m, 0, 0);
  },

  /** "17:00", "17.00", "5:00 PM" -> {h, m} atau null. */
  parseTime(s) {
    const mt = /^(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(AM|PM)?$/i.exec(Util.str(s));
    if (!mt) return null;
    let h = Number(mt[1]); const m = Number(mt[2]); const ap = mt[3] ? mt[3].toUpperCase() : '';
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    if (h > 23 || m > 59) return null;
    return { h: h, m: m };
  },

  /** Peta nama header -> nomor kolom (1-based). */
  headerMap(sh, row) {
    const lc = sh.getLastColumn();
    const vals = sh.getRange(row || 1, 1, 1, lc).getValues()[0];
    const map = {};
    vals.forEach(function (v, i) { const k = Util.str(v); if (k && !(k in map)) map[k] = i + 1; });
    return map;
  },

  requireHeaders(map, names, sheetName) {
    const missing = names.filter(function (n) { return !(n in map); });
    if (missing.length) {
      throw new Error('Sheet "' + sheetName + '": header tidak ditemukan -> ' + missing.join(', ') +
        '. Jangan mengubah/menghapus teks header.');
    }
  },

  /** Baris terakhir yang kolom `col`-nya terisi (>= firstRow-1 bila kosong). */
  lastNonEmptyRow(sh, col, firstRow) {
    const lr = sh.getLastRow();
    if (lr < firstRow) return firstRow - 1;
    const vals = sh.getRange(firstRow, col, lr - firstRow + 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) if (Util.str(vals[i][0]) !== '') return firstRow + i;
    return firstRow - 1;
  },

  activeEmail() {
    let e = '';
    try { e = Session.getActiveUser().getEmail(); } catch (x) { e = ''; }
    if (!e) { try { e = Session.getEffectiveUser().getEmail(); } catch (x2) { e = ''; } }
    return Util.str(e);
  },

  ui() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } },

  alert(title, msg) {
    const ui = Util.ui();
    if (ui) ui.alert(title, msg, ui.ButtonSet.OK); else Logger.log('[' + title + '] ' + msg);
  },

  toast(msg, title) { try { Util.ss().toast(msg, title || 'Task System', 6); } catch (e) { Logger.log(msg); } },

  /** Jalankan fn dengan error handling terpusat (untuk fungsi menu). */
  run(label, fn) {
    try { Util.assertInternal(); return fn(); }
    catch (e) {
      Log.error(label, e);
      Util.alert(label + ' - GAGAL', String(e && e.message ? e.message : e) +
        '\n\nTidak ada data yang dihapus. Perbaiki penyebabnya lalu ulangi.');
      return null;
    }
  },

  /** Fungsi menu hanya boleh dari Sheet/editor, bukan dari halaman web app (google.script.run oleh pengunjung). */
  /** 'yyyy-MM-dd' -> Date lokal, atau null bila format/tanggal tidak nyata (mis. 2026-13-01, 2026-02-30). */
  ymd(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
    if (!m) return null;
    const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
    const dt = new Date(y, mo - 1, d);
    return dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d && y >= 2000 && y <= 2100 ? dt : null;
  },

  assertInternal() {
    if (Util.ui()) return;
    let a = '', e = '';
    try { a = Session.getActiveUser().getEmail(); } catch (x) { a = ''; }
    try { e = Session.getEffectiveUser().getEmail(); } catch (x) { e = ''; }
    if (!a || !e || a.toLowerCase() !== e.toLowerCase()) throw new Error('Fungsi ini hanya dapat dijalankan dari menu Sheet / editor Apps Script.');
  },

  /** Jalankan fn di dalam script lock. */
  withLock(fn) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      throw new UserError('Sistem sedang dipakai pengguna lain. Tunggu beberapa detik lalu coba lagi.');
    }
    try { return fn(); } finally { lock.releaseLock(); }
  },

  assertTimezone() {
    const a = Util.tz(); const b = Session.getScriptTimeZone();
    if (a !== b) {
      throw new Error('Zona waktu tidak sama: spreadsheet = ' + a + ', script = ' + b +
        '. Samakan di File > Settings (spreadsheet) dan Project Settings > appsscript.json "timeZone" (script).');
    }
  }
};

// ============================================================================
//  Log
// ============================================================================
const Log = {
  info(tag, msg) { try { console.log('[' + tag + '] ' + msg); } catch (e) { /* noop */ } Logger.log('[' + tag + '] ' + msg); },
  error(tag, err) {
    const m = '[' + tag + '] ' + (err && err.stack ? err.stack : err);
    try { console.error(m); } catch (e) { /* noop */ }
    Logger.log(m);
    Log.sheet('ERROR', tag, err && err.message ? err.message : String(err));
  },
  /** Catatan permanen di sheet SYSTEM LOG (form HP, email, penolakan). Tidak pernah melempar error. */
  audit(level, tag, msg) { Log.info(tag, msg); Log.sheet(level, tag, msg); },
  sheet(level, tag, msg) {
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SYSTEM_LOG);
      if (!sh) return;
      const row = Math.max(1, Util.lastNonEmptyRow(sh, 1, 2)) + 1;
      sh.getRange(row, 1, 1, 4).setValues([[new Date(), tag, level, String(msg).substring(0, 1500)]]);
      sh.getRange(row, 1).setNumberFormat(CONFIG.FORMATS.DATETIME);
    } catch (e) { /* log tidak boleh menggagalkan proses utama */ }
  }
};

// ============================================================================
//  Settings
// ============================================================================
const Settings = {
  load() {
    const L = CONFIG.SETTINGS_LAYOUT;
    const sh = Util.sheet(CONFIG.SHEETS.SETTINGS);
    const n = L.LIST_LAST - L.LIST_FIRST + 1;
    const col = function (c) {
      return sh.getRange(L.LIST_FIRST, c, n, 1).getDisplayValues().map(function (r) { return Util.str(r[0]); });
    };
    const priorities = col(L.COL.PRIORITY).filter(Boolean);
    const statuses = col(L.COL.STATUS).filter(Boolean).map(Util.up);
    const an = col(L.COL.ADMIN_NAME), ae = col(L.COL.ADMIN_EMAIL);
    const admins = [];
    an.forEach(function (name, i) { if (name) admins.push({ name: name, email: ae[i] || '' }); });
    const dc = col(L.COL.DELAY_CAT), dcs = col(L.COL.DELAY_COUNTS);
    const delayCategories = [];
    dc.forEach(function (name, i) { if (name) delayCategories.push({ name: name, countsAgainstStaff: Util.up(dcs[i]) !== 'NO' }); });
    const active = col(L.COL.ACTIVE).filter(Boolean);
    const ac = col(L.COL.ADMIN_CODE);
    admins.forEach(function (a) { a.code = ''; });
    an.forEach(function (name, i) { if (name) admins.filter(function (a) { return a.name === name; })[0].code = ac[i] || ''; });
    const ql = col(L.COL.Q_LABEL), qs = col(L.COL.Q_SCORE);
    const quality = [];
    ql.forEach(function (lab, i) { if (lab) { const sc = Number(qs[i]); quality.push({ label: lab, score: isNaN(sc) ? 0 : sc }); } });
    let weights = sh.getRange(L.WEIGHT_FIRST, L.CFG_VAL_COL, 4, 1).getValues().map(function (r) { return Number(r[0]); });
    if (weights.some(function (x) { return isNaN(x) || x < 0; }) || weights.reduce(function (a, b) { return a + b; }, 0) <= 0) weights = [0.25, 0.30, 0.30, 0.15];

    const cfg = {};
    const cn = L.CFG_LAST - L.CFG_FIRST + 1;
    const keys = sh.getRange(L.CFG_FIRST, L.CFG_KEY_COL, cn, 1).getDisplayValues();
    const vals = sh.getRange(L.CFG_FIRST, L.CFG_VAL_COL, cn, 1).getDisplayValues();
    for (let i = 0; i < cn; i++) { const k = Util.str(keys[i][0]); if (k) cfg[k] = Util.str(vals[i][0]); }
    Object.keys(CONFIG.DEFAULTS).forEach(function (k) { if (!(k in cfg) || cfg[k] === '') cfg[k] = CONFIG.DEFAULTS[k]; });

    return { cfg: cfg, priorities: priorities, statuses: statuses, admins: admins, delayCategories: delayCategories,
             activeLabel: active[0] || 'Active', quality: quality, weights: weights };
  },
  yes(S, key) { return Util.isYes(S.cfg[key]); },
  num(S, key) { const n = Number(S.cfg[key]); return isNaN(n) ? Number(CONFIG.DEFAULTS[key]) : n; }
};

// ============================================================================
//  Masters
// ============================================================================
const Masters = {
  _load(sheetName, S) {
    const sh = Util.sheet(sheetName);
    const last = Util.lastNonEmptyRow(sh, 2, 2);
    const out = { byName: {}, list: [] };
    if (last < 2) return out;
    const v = sh.getRange(2, 1, last - 1, 6).getDisplayValues();
    v.forEach(function (r) {
      const name = Util.str(r[1]);
      if (!name) return;
      const o = { id: Util.str(r[0]), name: name, c3: Util.str(r[2]), c4: Util.str(r[3]), code: Util.str(r[5]),
                  active: Util.low(r[4]) === Util.low(S.activeLabel) };
      out.byName[Util.low(name)] = o; out.list.push(o);
    });
    return out;
  },
  staff(S) {   // {id, name, email, position, active}
    const m = Masters._load(CONFIG.SHEETS.STAFF, S);
    m.list.forEach(function (o) { o.email = o.c3; o.position = o.c4; });
    return m;
  },
  projects(S) {   // {id, name, location, manager, active}
    const m = Masters._load(CONFIG.SHEETS.PROJECT, S);
    m.list.forEach(function (o) { o.location = o.c3; o.manager = o.c4; });
    return m;
  }
};

// ============================================================================
//  Db  (TASK DATABASE)
// ============================================================================
const Db = {
  open() {
    const sh = Util.sheet(CONFIG.SHEETS.DB);
    const h = Util.headerMap(sh, 1);
    const H = CONFIG.DB_H;
    Util.requireHeaders(h, Object.keys(H).map(function (k) { return H[k]; }), CONFIG.SHEETS.DB);
    return { sheet: sh, h: h, width: h[H.SUBMITTED_BY], autoStart: h[H.AUTO_FIRST] };
  },

  /** Baca semua baris data (kolom input A..Submitted By). */
  readAll(db) {
    const H = CONFIG.DB_H;
    const last = Util.lastNonEmptyRow(db.sheet, db.h[H.TASK_ID], 2);
    const out = { last: last, rows: [], byId: {}, fingerprints: {} };
    if (last < 2) return out;
    const vals = db.sheet.getRange(2, 1, last - 1, db.width).getValues();
    vals.forEach(function (v, i) {
      const o = { _row: i + 2, _vals: v };
      Object.keys(db.h).forEach(function (name) { if (db.h[name] <= db.width) o[name] = v[db.h[name] - 1]; });
      const id = Util.up(o[H.TASK_ID]);
      if (!id) return;
      out.rows.push(o); out.byId[id] = o;
      if (Util.isDate(o[H.DEADLINE])) out.fingerprints[Db.fingerprint(o[H.PROJECT_ID], o[H.ASSIGNED_TO], o[H.TITLE], o[H.DEADLINE])] = id;
    });
    return out;
  },

  fingerprint(projectId, pic, title, deadline) {
    return [Util.up(projectId), Util.norm(pic), Util.norm(title), Math.round(deadline.getTime() / 60000)].join('|');
  },

  /** Pastikan rumus kolom otomatis ada sampai `upToRow` (salin dari baris rumus terakhir). */
  ensureFormulas(db, upToRow) {
    if (upToRow > CONFIG.DB_MAX_ROWS + 1) {
      throw new Error('TASK DATABASE penuh (' + CONFIG.DB_MAX_ROWS + ' tugas). Arsipkan dulu atau naikkan DB_MAX_ROWS beserta rumus di workbook.');
    }
    const sh = db.sheet; const col = db.autoStart;
    let end = sh.getLastRow();
    if (end < 2) throw new Error('Rumus kolom otomatis TASK DATABASE tidak ditemukan. Gunakan file XLSX asli.');
    if (!Util.str(sh.getRange(end, col).getFormula())) {
      throw new Error('Rumus kolom otomatis di TASK DATABASE baris ' + end + ' kosong. Salin rumus dari baris 2 ke bawah atau gunakan file XLSX asli.');
    }
    if (upToRow <= end) return;
    const target = Math.min(CONFIG.DB_MAX_ROWS + 1, Math.max(upToRow, end + CONFIG.FORMULA_EXTEND_CHUNK));
    const nAuto = sh.getLastColumn() - col + 1;
    sh.getRange(end, col, 1, nAuto).copyTo(sh.getRange(end + 1, col, target - end, nAuto));
  },

  /** Tambah baris baru di akhir (tidak pernah menimpa baris yang ada). `objs` = array {header: value}. */
  append(db, objs) {
    if (!objs.length) return 0;
    const H = CONFIG.DB_H;
    const start = Util.lastNonEmptyRow(db.sheet, db.h[H.TASK_ID], 2) + 1;
    Db.ensureFormulas(db, start + objs.length - 1);
    const names = Object.keys(db.h).filter(function (n) { return db.h[n] <= db.width; }).sort(function (a, b) { return db.h[a] - db.h[b]; });
    const arr = objs.map(function (o) {
      const row = new Array(db.width).fill('');
      names.forEach(function (n) { if (o[n] !== undefined && o[n] !== null) row[db.h[n] - 1] = o[n]; });
      return row;
    });
    // kolom teks bebas diformat "plain text" lebih dulu: isi yang diawali "=" tidak boleh menjadi rumus
    [H.TITLE, H.DESC].forEach(function (n) { db.sheet.getRange(start, db.h[n], arr.length, 1).setNumberFormat('@'); });
    db.sheet.getRange(start, 1, arr.length, db.width).setValues(arr);
    const F = CONFIG.FORMATS;
    const fmts = {};
    fmts[H.CREATED_AT] = F.DATETIME; fmts[H.START] = F.DATE; fmts[H.DEADLINE] = F.DATETIME; fmts[H.PROGRESS] = F.PERCENT;
    fmts[H.COMPLETION] = F.DATETIME; fmts[H.LAST_UPDATED] = F.DATETIME;
    Object.keys(fmts).forEach(function (n) { db.sheet.getRange(start, db.h[n], arr.length, 1).setNumberFormat(fmts[n]); });
    return start;
  },

  /** Tulis kembali blok kolom Current Status..Last Updated untuk satu baris. */
  writeStatusBlock(db, row, o) {
    const H = CONFIG.DB_H;
    const c1 = db.h[H.STATUS], c2 = db.h[H.LAST_UPDATED];
    const vals = [];
    for (let c = c1; c <= c2; c++) {
      const name = Object.keys(db.h).filter(function (n) { return db.h[n] === c; })[0];
      vals.push(o[name] === undefined || o[name] === null ? '' : o[name]);
    }
    [H.DELAY_REASON, H.RESULT_NOTES].forEach(function (n) { db.sheet.getRange(row, db.h[n]).setNumberFormat('@'); });
    db.sheet.getRange(row, c1, 1, c2 - c1 + 1).setValues([vals]);
    db.sheet.getRange(row, db.h[H.PROGRESS]).setNumberFormat(CONFIG.FORMATS.PERCENT);
    db.sheet.getRange(row, db.h[H.COMPLETION]).setNumberFormat(CONFIG.FORMATS.DATETIME);
    db.sheet.getRange(row, db.h[H.LAST_UPDATED]).setNumberFormat(CONFIG.FORMATS.DATETIME);
  }
};

// ============================================================================
//  Ids
// ============================================================================
const Ids = {
  parseSeq(id, prefix) {
    const m = new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-(\\d+)$', 'i').exec(Util.str(id));
    return m ? Number(m[1]) : 0;
  },
  maxSeq(ids, prefix) { let mx = 0; ids.forEach(function (i) { mx = Math.max(mx, Ids.parseSeq(i, prefix)); }); return mx; },
  /** Reservasi `count` nomor berurutan. Wajib dipanggil di dalam lock. Mengembalikan nomor pertama. */
  reserve(propKey, existingIds, prefix, count) {
    const props = PropertiesService.getScriptProperties();
    const saved = Number(props.getProperty(propKey) || 0) || 0;
    const first = Math.max(saved, Ids.maxSeq(existingIds, prefix)) + 1;
    return { first: first, commit: function () { props.setProperty(propKey, String(first + count - 1)); } };
  },
  format(prefix, n, digits) { return prefix + '-' + Util.pad(n, digits); }
};

// ============================================================================
//  Validate  (baris TASK INPUT)
// ============================================================================
const Validate = {
  /**
   * @param raw  {createdBy, project, pic, parent, title, desc, priority, inputDate, start, deadlineDate, notes}
   * @param ctx  {S, staff, projects, db, batchFp, activeEmail, now}
   * @return {errors: string[], clean: object}
   */
  inputRow(raw, ctx) {
    const E = []; const S = ctx.S; const clean = {};

    // Created By
    const cb = Util.str(raw.createdBy);
    if (!cb) E.push('Created By wajib diisi');
    else {
      const adm = S.admins.filter(function (a) { return Util.low(a.name) === Util.low(cb); })[0];
      if (!adm) E.push('Created By "' + cb + '" bukan admin yang diotorisasi (SETTINGS > Admin Name)');
      else {
        clean.createdBy = adm.name; clean.adminEmail = adm.email;
        if (Settings.yes(S, 'ENFORCE_CREATED_BY_MATCH') && adm.email && ctx.activeEmail &&
            Util.low(adm.email) !== Util.low(ctx.activeEmail)) {
          E.push('Akun yang login (' + ctx.activeEmail + ') tidak sama dengan email admin "' + adm.name + '"');
        }
      }
    }

    // Project
    const pj = Util.str(raw.project);
    if (!pj) E.push('Project wajib diisi');
    else {
      const p = ctx.projects.byName[Util.low(pj)];
      if (!p) E.push('Project "' + pj + '" tidak ada di PROJECT MASTER');
      else if (!p.active) E.push('Project "' + pj + '" berstatus Inactive');
      else clean.project = p;
    }

    // PIC
    const pic = Util.str(raw.pic);
    if (!pic) E.push('Assigned To (PIC) wajib diisi');
    else {
      const s = ctx.staff.byName[Util.low(pic)];
      if (!s) E.push('PIC "' + pic + '" tidak ada di STAFF MASTER');
      else if (!s.active) E.push('PIC "' + pic + '" berstatus Inactive');
      else clean.staff = s;
    }

    // Title & priority
    const title = Util.str(raw.title);
    if (!title) E.push('Task Title wajib diisi');
    else if (title.length > CONFIG.LIMITS.TITLE) E.push('Task Title terlalu panjang (maks ' + CONFIG.LIMITS.TITLE + ' karakter)');
    else clean.title = title;
    if (Util.str(raw.desc).length > CONFIG.LIMITS.TEXT) E.push('Task Description terlalu panjang (maks ' + CONFIG.LIMITS.TEXT + ' karakter)');
    if (Util.str(raw.notes).length > CONFIG.LIMITS.TEXT) E.push('Notes terlalu panjang (maks ' + CONFIG.LIMITS.TEXT + ' karakter)');
    const pr = Util.str(raw.priority);
    if (!pr) E.push('Priority wajib diisi');
    else {
      const m = S.priorities.filter(function (x) { return Util.low(x) === Util.low(pr); })[0];
      if (!m) E.push('Priority "' + pr + '" tidak valid (' + S.priorities.join(', ') + ')'); else clean.priority = m;
    }

    // Dates
    const today = Util.startOfDay(ctx.now);
    let inputDate = today;
    if (Util.str(raw.inputDate) !== '') {
      if (Util.isDate(raw.inputDate)) inputDate = raw.inputDate; else E.push('Format Input Date tidak valid');
    }
    clean.inputDate = inputDate;

    let start = inputDate;
    if (Util.str(raw.start) !== '') {
      if (Util.isDate(raw.start)) start = raw.start; else E.push('Format Start Date tidak valid');
    }
    clean.start = start;

    let deadline = null;
    if (Util.str(raw.deadlineDate) === '') E.push('Deadline Date wajib diisi');
    else if (!Util.isDate(raw.deadlineDate)) E.push('Format Deadline Date tidak valid (gunakan dd/mm/yyyy)');
    else {
      const eod = CONFIG.DEADLINE_END_OF_DAY;
      deadline = Util.combine(raw.deadlineDate, eod.h, eod.m);   // akhir hari pada tanggal deadline
      clean.deadline = deadline;
      if (Util.startOfDay(deadline).getTime() < Util.startOfDay(start).getTime()) E.push('Deadline lebih awal dari Start Date');
      if (!Settings.yes(S, 'ALLOW_PAST_DEADLINE') && Util.startOfDay(deadline).getTime() < today.getTime()) {
        E.push('Deadline sudah lewat (ubah tanggal, atau set ALLOW_PAST_DEADLINE = YES di SETTINGS)');
      }
    }

    // Parent
    const parent = Util.up(raw.parent);
    if (parent) {
      if (!ctx.db.byId[parent]) E.push('Parent Task ID "' + parent + '" tidak ditemukan di TASK DATABASE (kirim tugas induk terlebih dulu)');
      else clean.parent = ctx.db.byId[parent][CONFIG.DB_H.TASK_ID];
    }

    // Duplicate (hanya bila field kunci valid)
    if (clean.project && clean.staff && clean.title && deadline) {
      const fp = Db.fingerprint(clean.project.id, clean.staff.name, clean.title, deadline);
      if (ctx.db.fingerprints[fp]) E.push('Duplikat: tugas identik sudah ada di database (' + ctx.db.fingerprints[fp] + '). Ubah judul/deadline bila memang tugas berbeda');
      else if (ctx.batchFp[fp]) E.push('Duplikat dengan baris ' + ctx.batchFp[fp] + ' pada pengiriman yang sama');
      clean.fp = fp;
    }

    clean.desc = Util.str(raw.desc); clean.notes = Util.str(raw.notes);
    return { errors: E, clean: clean };
  }
};

// ============================================================================
//  Submit  (TASK INPUT -> TASK DATABASE)
// ============================================================================
const Submit = {
  /** Jalur SHEET (laptop): kirim baris TASK INPUT ke TASK DATABASE. */
  run() {
    Util.assertTimezone();
    let result = null;
    Util.withLock(function () { result = Submit.process_(new Date()); });
    // email di luar lock; kegagalan email tidak membatalkan data
    let mail = { sent: 0, failed: [], skipped: 'tidak ada tugas baru' };
    if (result.submitted.length) {
      try { mail = Notify.onSubmit(result); } catch (e) { Log.error('Notify.onSubmit', e); mail = { sent: 0, failed: [String(e.message || e)], skipped: '' }; }
      Forms.refreshQuiet();
    }
    result.mail = mail;
    return result;
  },

  /** Konteks validasi (settings, master, database). Wajib dipanggil di dalam lock. */
  context_(now, email) {
    const S = Settings.load();
    const dbh = Db.open();
    return { S: S, staff: Masters.staff(S), projects: Masters.projects(S), db: Db.readAll(dbh), dbh: dbh,
             batchFp: {}, activeEmail: email, now: now };
  },

  /**
   * Tulis tugas valid ke TASK DATABASE (hanya tambah baris). Wajib di dalam lock.
   * @param valid [{clean, rowNum?}]  -> mengisi .id dan mengembalikan ringkasan untuk email/UI
   */
  write_(ctx, valid) {
    const H = CONFIG.DB_H, S = ctx.S, now = ctx.now;
    const existingIds = ctx.db.rows.map(function (o) { return o[H.TASK_ID]; });
    const prefix = S.cfg.TASK_ID_PREFIX || 'TSK';
    const digits = Math.max(1, Number(S.cfg.TASK_ID_DIGITS) || 5);
    const res = Ids.reserve(CONFIG.PROP.TASK_SEQ, existingIds, prefix, valid.length);
    const objs = valid.map(function (it, k) {
      const c = it.clean;
      it.id = Ids.format(prefix, res.first + k, digits);
      const desc = c.notes ? (c.desc ? c.desc + '\n' : '') + 'Catatan: ' + c.notes : c.desc;
      const o = {};
      o[H.TASK_ID] = it.id; o[H.CREATED_AT] = now; o[H.CREATED_BY] = c.createdBy; o[H.PROJECT_ID] = c.project.id;
      o[H.PROJECT_NAME] = c.project.name; o[H.ASSIGNED_TO] = c.staff.name; o[H.PARENT] = c.parent || '';
      o[H.TITLE] = c.title; o[H.DESC] = desc; o[H.PRIORITY] = c.priority; o[H.START] = c.start; o[H.DEADLINE] = c.deadline;
      o[H.STATUS] = CONFIG.STATUS.NEW; o[H.PROGRESS] = 0; o[H.LAST_UPDATED] = now; o[H.SUBMITTED_BY] = ctx.activeEmail;
      return o;
    });
    Db.append(ctx.dbh, objs);   // gagal -> exception: ID belum di-commit, tidak ada yang ditandai terkirim
    res.commit();
    return valid.map(function (it) {
      const c = it.clean;
      return { row: it.rowNum, id: it.id, title: c.title, project: c.project.name, pic: c.staff.name, picEmail: c.staff.email,
               priority: c.priority, deadline: c.deadline, createdBy: c.createdBy, adminEmail: c.adminEmail };
    });
  },

  process_(now) {
    const C = CONFIG.INPUT.COL, IN = CONFIG.INPUT, H = CONFIG.DB_H;
    const email = Util.activeEmail();
    const ctx = Submit.context_(now, email);
    const dbData = ctx.db;
    const sh = Util.sheet(CONFIG.SHEETS.INPUT);
    const n = IN.LAST_ROW - IN.FIRST_ROW + 1;
    const vals = sh.getRange(IN.FIRST_ROW, 1, n, C.STATUS).getValues();

    const result = { submitted: [], errors: [], healed: [], alreadySubmitted: 0, empty: 0, emailUsed: email };
    const valid = [];      // {rowNum, clean, raw}
    const healRows = [];   // baris yang ID-nya sudah ada di database tetapi status belum SUBMITTED

    for (let i = 0; i < n; i++) {
      const v = vals[i]; const rowNum = IN.FIRST_ROW + i;
      if (Util.up(v[C.STATUS - 1]) === IN.SUBMITTED) { result.alreadySubmitted++; continue; }
      let hasData = false;
      for (let c = C.INPUT_DATE; c <= C.NOTES; c++) if (Util.str(v[c - 1]) !== '') { hasData = true; break; }
      if (!hasData) { result.empty++; continue; }

      const existingId = Util.up(v[C.TASK_ID - 1]);
      if (existingId && dbData.byId[existingId]) { healRows.push({ rowNum: rowNum, id: dbData.byId[existingId][H.TASK_ID] }); continue; }

      const raw = {
        createdBy: v[C.CREATED_BY - 1], project: v[C.PROJECT - 1], pic: v[C.PIC - 1], parent: v[C.PARENT - 1],
        title: v[C.TITLE - 1], desc: v[C.DESC - 1], priority: v[C.PRIORITY - 1], inputDate: v[C.INPUT_DATE - 1],
        start: v[C.START - 1], deadlineDate: v[C.DEADLINE_DATE - 1], notes: v[C.NOTES - 1]
      };
      const r = Validate.inputRow(raw, ctx);
      if (r.errors.length) { result.errors.push({ row: rowNum, messages: r.errors }); continue; }
      ctx.batchFp[r.clean.fp] = rowNum;
      valid.push({ rowNum: rowNum, clean: r.clean, raw: raw });
    }

    // Pulihkan baris yang sudah ada di database (mis. script terhenti setelah menulis database)
    healRows.forEach(function (h) {
      sh.getRange(h.rowNum, C.TASK_ID).setValue(h.id);
      sh.getRange(h.rowNum, C.STATUS).setValue(IN.SUBMITTED);
      result.healed.push({ row: h.rowNum, id: h.id });
    });
    if (!valid.length) return result;

    const summary = Submit.write_(ctx, valid);
    valid.forEach(function (it, k) {
      const c = it.clean;
      sh.getRange(it.rowNum, C.TASK_ID).setValue(it.id);
      if (Util.str(it.raw.inputDate) === '') sh.getRange(it.rowNum, C.INPUT_DATE).setValue(c.inputDate);
      if (Util.str(it.raw.start) === '') sh.getRange(it.rowNum, C.START).setValue(c.start);
      sh.getRange(it.rowNum, C.STATUS).setValue(IN.SUBMITTED);
      result.submitted.push(summary[k]);
    });
    return result;
  },

  /** Jalur WEB (admin): banyak tugas sekaligus; sebagian boleh gagal. items = [{project,pic,title,desc,priority,start,deadline,parent,notes}] */
  fromWeb(user, items) {
    Util.assertTimezone();
    let out = null;
    Util.withLock(function () {
      const now = new Date();
      const ctx = Submit.context_(now, user.email);
      const results = []; const valid = [];
      items.forEach(function (it, i) {
        const raw = { createdBy: user.name, project: it.project, pic: it.pic, parent: it.parent, title: it.title, desc: it.desc,
                      priority: it.priority, inputDate: '', start: Submit.formDate_(it.start), deadlineDate: Submit.formDate_(it.deadline), notes: it.notes };
        const r = Validate.inputRow(raw, ctx);
        if (r.errors.length) { results[i] = { ok: false, errors: r.errors }; return; }
        ctx.batchFp[r.clean.fp] = 'ke-' + (i + 1);
        valid.push({ idx: i, clean: r.clean, raw: raw });
      });
      const summary = valid.length ? Submit.write_(ctx, valid) : [];
      valid.forEach(function (v, k) { results[v.idx] = { ok: true, id: summary[k].id }; });
      out = { results: results, submitted: summary };
    });
    let mail = { sent: 0, failed: [], skipped: 'tidak ada tugas baru' };
    if (out.submitted.length) {
      Log.audit('INFO', 'WEB-TASK', 'Tugas ' + out.submitted.map(function (s) { return s.id; }).join(', ') + ' dibuat oleh ' + user.name);
      try { mail = Notify.onSubmit({ submitted: out.submitted, errors: [], healed: [] }); } catch (e) { Log.error('Notify.onSubmit', e); mail = { sent: 0, failed: [String(e.message || e)], skipped: '' }; }
      Forms.refreshQuiet();
    }
    return { results: out.results, created: out.submitted.length, mail: { sent: mail.sent || 0, failed: mail.failed || [], skipped: mail.skipped || '' } };
  },

  /** Jalur FORM HP: satu tanggapan Google Form = satu tugas. `map` = {judul pertanyaan: jawaban}. */
  fromForm(map, email) {
    Util.assertTimezone();
    let res = null;
    Util.withLock(function () { res = Submit.formCore_(map, email, new Date()); });
    if (res.submitted.length) {
      Log.audit('INFO', 'FORM-TASK', 'Tugas ' + res.submitted.map(function (s) { return s.id; }).join(', ') + ' dibuat oleh ' + email);
      try { res.mail = Notify.onSubmit(res); } catch (e) { Log.error('Notify.onSubmit', e); res.mail = { sent: 0, failed: [String(e.message || e)], skipped: '' }; }
      Forms.refreshQuiet();
    } else {
      const msgs = res.errors[0].messages;
      Log.audit('WARN', 'FORM-TASK', 'DITOLAK (' + (email || 'tanpa email') + '): ' + msgs.join('; '));
      Notify.rejectToSender(email, 'Tugas TIDAK tercatat', msgs);
    }
    return res;
  },

  formCore_(map, email, now) {
    const ctx = Submit.context_(now, email);
    const T = CONFIG.FORMS.TASK; const S = ctx.S;
    const adm = S.admins.filter(function (a) { return a.email && email && Util.low(a.email) === Util.low(email); })[0];
    if (!adm) {
      return { submitted: [], errors: [{ ref: 'form', messages: ['Akun ' + (email || '(email tidak terbaca)') +
        ' bukan admin yang diotorisasi. Daftarkan emailnya di SETTINGS > Admin Email.'] }] };
    }
    const raw = {
      createdBy: adm.name, project: map[T.PROJECT], pic: map[T.PIC], parent: map[T.PARENT], title: map[T.TITLE],
      desc: map[T.DESC], priority: map[T.PRIORITY], inputDate: '', start: Submit.formDate_(map[T.START]),
      deadlineDate: Submit.formDate_(map[T.DEADLINE]), notes: map[T.NOTES]
    };
    const r = Validate.inputRow(raw, ctx);
    if (r.errors.length) return { submitted: [], errors: [{ ref: 'form', messages: r.errors }] };
    return { submitted: Submit.write_(ctx, [{ clean: r.clean, raw: raw }]), errors: [] };
  },

  /** Jawaban pertanyaan tanggal Google Form: 'yyyy-MM-dd' -> Date (tengah malam, zona script). */
  formDate_(v) {
    const s = Util.str(v);
    if (!s) return '';
    return Util.ymd(s.trim()) || s;
  },

  summaryText(r) {
    const lines = [];
    lines.push('Terkirim ke database: ' + r.submitted.length + ' tugas');
    if (r.submitted.length) lines.push('ID: ' + r.submitted.map(function (s) { return s.id; }).slice(0, 12).join(', ') + (r.submitted.length > 12 ? ' ...' : ''));
    if (r.healed.length) lines.push('Dipulihkan (sudah ada di database, tidak dikirim ulang): ' + r.healed.length);
    if (r.alreadySubmitted) lines.push('Dilewati (sudah SUBMITTED): ' + r.alreadySubmitted);
    if (r.errors.length) {
      lines.push(''); lines.push('GAGAL VALIDASI: ' + r.errors.length + ' baris (belum dikirim, silakan perbaiki lalu Submit lagi):');
      r.errors.slice(0, CONFIG.MAX_ERRORS_IN_ALERT).forEach(function (e) { lines.push('  Baris ' + e.row + ': ' + e.messages.join('; ')); });
      if (r.errors.length > CONFIG.MAX_ERRORS_IN_ALERT) lines.push('  ... dan ' + (r.errors.length - CONFIG.MAX_ERRORS_IN_ALERT) + ' baris lain');
    }
    if (!r.submitted.length && !r.errors.length && !r.healed.length) lines.push('Tidak ada baris baru untuk dikirim.');
    if (r.submitted.length && r.mail) {
      lines.push('');
      if (r.mail.skipped) lines.push('EMAIL TIDAK DIKIRIM: ' + r.mail.skipped);
      else {
        lines.push('Email terkirim: ' + r.mail.sent);
        if (r.mail.failed.length) lines.push('Email GAGAL: ' + r.mail.failed.join('; '));
      }
      lines.push('(rincian email ada di sheet SYSTEM LOG)');
    }
    return lines.join('\n');
  }
};

// ============================================================================
//  Progress  (PROGRESS LOG -> TASK DATABASE)
// ============================================================================
const Progress = {
  /** Proses semua baris PROGRESS LOG yang belum punya Log ID. */
  run() {
    Util.assertTimezone();
    let result = null;
    Util.withLock(function () { result = Progress.process_(new Date()); });
    try { Notify.onProgress(result); } catch (e) { Log.error('Notify.onProgress', e); }
    if (result.processed.length) Forms.refreshQuiet();
    return result;
  },

  process_(now) {
    const P = CONFIG.PL_H, H = CONFIG.DB_H, ST = CONFIG.STATUS;
    const S = Settings.load();
    const db = Db.open(); const dbData = Db.readAll(db);
    const sh = Util.sheet(CONFIG.SHEETS.PROGRESS);
    const h = Util.headerMap(sh, 1);
    Util.requireHeaders(h, Object.keys(P).map(function (k) { return P[k]; }), CONFIG.SHEETS.PROGRESS);

    const last = Math.max(Util.lastNonEmptyRow(sh, h[P.LOG_ID], 2), Util.lastNonEmptyRow(sh, h[P.TASK_ID], 2),
                          Util.lastNonEmptyRow(sh, h[P.NEW], 2), Util.lastNonEmptyRow(sh, h[P.UPDATED_BY], 2));
    const result = { processed: [], errors: [], skipped: 0 };
    if (last < 2) return result;
    const width = sh.getLastColumn();
    const vals = sh.getRange(2, 1, last - 1, width).getValues();
    const existingLogIds = vals.map(function (r) { return r[h[P.LOG_ID] - 1]; });
    const prefix = CONFIG.LOG_PREFIX;
    const cands = [];
    vals.forEach(function (r, i) {
      if (Util.str(r[h[P.LOG_ID] - 1]) !== '') { result.skipped++; return; }
      let any = false;
      [P.TASK_ID, P.UPDATED_BY, P.NEW, P.PROGRESS, P.NOTES, P.DELAY, P.REVISED].forEach(function (k) { if (Util.str(r[h[k] - 1]) !== '') any = true; });
      if (any) cands.push({ i: i, rowNum: i + 2, r: r });
    });
    if (!cands.length) return result;

    const res = Ids.reserve(CONFIG.PROP.LOG_SEQ, existingLogIds, prefix, cands.length);
    let seq = res.first; let consumed = 0;
    const dirty = {};     // rowNum(DB) -> obj
    const closed = [ST.COMPLETED, ST.CANCELLED];

    cands.forEach(function (cd) {
      const r = cd.r; const E = [];
      const taskId = Util.up(r[h[P.TASK_ID] - 1]);
      const by = Util.str(r[h[P.UPDATED_BY] - 1]);
      const ns = Util.up(r[h[P.NEW] - 1]);
      const notes = Util.str(r[h[P.NOTES] - 1]);
      const delay = Util.str(r[h[P.DELAY] - 1]);
      const revised = r[h[P.REVISED] - 1];
      const t = taskId ? dbData.byId[taskId] : null;

      if (!taskId) E.push('Task ID wajib diisi'); else if (!t) E.push('Task ID "' + taskId + '" tidak ada di TASK DATABASE');
      if (!by) E.push('Updated By wajib diisi');
      if (!ns) E.push('New Status wajib diisi'); else if (S.statuses.indexOf(ns) < 0) E.push('New Status "' + ns + '" tidak valid');
      if (Util.str(revised) !== '' && !Util.isDate(revised)) E.push('Format Revised Estimated Completion tidak valid');

      let prog; const pr = r[h[P.PROGRESS] - 1];
      if (Util.str(pr) !== '') {
        prog = Number(pr);
        if (isNaN(prog)) E.push('Progress Percentage harus angka / persen');
        else { if (prog > 1) prog = prog / 100; if (prog < 0 || prog > 1) E.push('Progress Percentage harus 0%-100%'); }
      }

      let prev = '';
      if (t && !E.length) {
        prev = Util.up(t[H.STATUS]);
        const reopen = Settings.yes(S, 'ALLOW_REOPEN_CLOSED');
        if (closed.indexOf(prev) >= 0 && !reopen) E.push('Tugas sudah ' + prev + ' dan tidak dapat diubah (ALLOW_REOPEN_CLOSED = NO)');
        if (ns === ST.NEW && prev !== ST.NEW) E.push('Status tidak dapat dikembalikan ke NEW');
        if (ns === ST.COMPLETED) {
          if (prog !== undefined && prog < 1) E.push('Status COMPLETED harus progress 100% (kosongkan kolom progress untuk otomatis)');
          prog = 1;
        } else if (prog === 1 && ns !== ST.CANCELLED) {
          E.push('Progress 100% harus berstatus COMPLETED');
        }
        if (Settings.yes(S, 'REQUIRE_DELAY_REASON') && !delay) {
          const dl = t[H.DEADLINE];
          const late = Util.isDate(dl) && dl.getTime() < now.getTime();
          if (ns === ST.PENDING) E.push('Status PENDING wajib mengisi Delay Reason');
          else if (late && ns !== ST.CANCELLED) E.push('Tugas sudah melewati deadline: Delay Reason wajib diisi');
        }
      }

      if (E.length) {
        sh.getRange(cd.rowNum, h[P.NOTE]).setValue('ERROR: ' + E.join('; '));
        result.errors.push({ row: cd.rowNum, messages: E });
        return;
      }

      // --- terapkan
      const logId = Ids.format(prefix, seq++, 5); consumed++;
      const row = t._row;
      const before = t[H.STATUS];
      t[H.STATUS] = ns;
      t[H.PROGRESS] = prog !== undefined ? prog : (Number(t[H.PROGRESS]) || 0);
      if (ns === ST.COMPLETED) t[H.COMPLETION] = now;
      else if (closed.indexOf(prev) >= 0) t[H.COMPLETION] = '';
      if (delay) t[H.DELAY_REASON] = delay;
      if (ns === ST.COMPLETED && notes) t[H.RESULT_NOTES] = notes;
      t[H.LAST_UPDATED] = now;
      dirty[row] = t;

      sh.getRange(cd.rowNum, h[P.LOG_ID], 1, 2).setValues([[logId, now]]);
      sh.getRange(cd.rowNum, h[P.TS]).setNumberFormat(CONFIG.FORMATS.DATETIME);
      sh.getRange(cd.rowNum, h[P.PREV]).setValue(before);
      if (Util.str(pr) === '') sh.getRange(cd.rowNum, h[P.PROGRESS]).setValue(t[H.PROGRESS]);
      sh.getRange(cd.rowNum, h[P.NOTE]).setValue('OK');
      result.processed.push({ row: cd.rowNum, logId: logId, taskId: t[H.TASK_ID], prev: before, status: ns, progress: t[H.PROGRESS],
                              by: by, notes: notes, delay: delay, title: t[H.TITLE], pic: t[H.ASSIGNED_TO], createdBy: t[H.CREATED_BY] });
    });

    Object.keys(dirty).forEach(function (row) { Db.writeStatusBlock(db, Number(row), dirty[row]); });
    if (consumed) { PropertiesService.getScriptProperties().setProperty(CONFIG.PROP.LOG_SEQ, String(res.first + consumed - 1)); }
    return result;
  },

  summaryText(r) {
    const L = ['Progress diproses: ' + r.processed.length];
    if (r.skipped) L.push('Dilewati (sudah punya Log ID): ' + r.skipped);
    if (r.errors.length) {
      L.push(''); L.push('ERROR: ' + r.errors.length + ' baris (lihat kolom Processing Note):');
      r.errors.slice(0, CONFIG.MAX_ERRORS_IN_ALERT).forEach(function (e) { L.push('  Baris ' + e.row + ': ' + e.messages.join('; ')); });
    }
    return L.join('\n');
  },

  /** Jalur FORM HP progres. Identitas = email login: staf (hanya tugas miliknya) atau admin. */
  fromFormResponse(map, email) {
    const F = CONFIG.FORMS.PROGRESS, P = CONFIG.PL_H, H = CONFIG.DB_H;
    const reject = function (msgs) {
      Log.audit('WARN', 'FORM-PROGRESS', 'DITOLAK (' + (email || 'tanpa email') + '): ' + msgs.join('; '));
      Notify.rejectToSender(email, 'Update progres TIDAK tercatat', msgs);
      return { processed: [], errors: [{ row: 0, messages: msgs }], skipped: 0 };
    };
    Util.assertTimezone();
    const S = Settings.load();
    const choice = Util.str(map[F.TASK]);
    const taskId = Util.up(choice.split('|')[0]);
    const staff = Masters.staff(S);
    const st = staff.list.filter(function (x) { return x.active && x.email && email && Util.low(x.email) === Util.low(email); })[0];
    const adm = S.admins.filter(function (a) { return a.email && email && Util.low(a.email) === Util.low(email); })[0];
    if (!st && !adm) return reject(['Akun ' + (email || '(email tidak terbaca)') + ' bukan staf aktif atau admin. Samakan email di STAFF MASTER / SETTINGS.']);
    const by = adm ? adm.name : st.name;
    const t = Db.readAll(Db.open()).byId[taskId];
    if (!t) return reject(['Tugas "' + choice + '" tidak ditemukan di TASK DATABASE.']);
    if (!adm && Util.low(t[H.ASSIGNED_TO]) !== Util.low(st.name)) {
      return reject(['Anda (' + st.name + ') bukan PIC tugas ' + taskId + ' (PIC: ' + t[H.ASSIGNED_TO] + ').']);
    }
    const ns = Util.up(map[F.STATUS]);
    if (ns === CONFIG.STATUS.CANCELLED && !adm) return reject(['Status CANCELLED hanya boleh oleh admin.']);

    let pn = '';
    const ptxt = Util.str(map[F.PROGRESS]);
    if (ptxt !== '' && ptxt !== CONFIG.FORMS.KEEP_PROGRESS) {
      pn = Number(ptxt.replace('%', ''));
      if (isNaN(pn)) return reject(['Progress "' + ptxt + '" tidak valid.']);
      pn = pn > 1 ? pn / 100 : pn;
    }
    const revised = Submit.formDate_(map[F.REVISED]);

    const ar = Progress.appendAndRun_({ taskId: t[H.TASK_ID], by: by, ns: ns, pn: pn, notes: Util.str(map[F.NOTES]),
                                        delay: Util.str(map[F.DELAY]), revised: revised });
    if (ar.mine) return reject(ar.mine.messages);   // baris form yang ditolak sudah dibersihkan (jawaban asli tetap di Google Form)
    Log.audit('INFO', 'FORM-PROGRESS', taskId + ' -> ' + ns + ' oleh ' + by);
    return ar.result;
  },

  /** Tulis satu baris PROGRESS LOG lalu proses (jalur form HP & web). o = {taskId, by, ns, pn, notes, delay, revised}. */
  appendAndRun_(o) {
    const P = CONFIG.PL_H; let rowNum = 0;
    Util.withLock(function () {
      const sh = Util.sheet(CONFIG.SHEETS.PROGRESS); const h = Util.headerMap(sh, 1);
      rowNum = Math.max(Util.lastNonEmptyRow(sh, h[P.LOG_ID], 2), Util.lastNonEmptyRow(sh, h[P.TASK_ID], 2),
                        Util.lastNonEmptyRow(sh, h[P.NEW], 2), Util.lastNonEmptyRow(sh, h[P.UPDATED_BY], 2)) + 1;
      sh.getRange(rowNum, h[P.NOTES]).setNumberFormat('@'); sh.getRange(rowNum, h[P.DELAY]).setNumberFormat('@');
      sh.getRange(rowNum, h[P.TASK_ID]).setValue(o.taskId);
      sh.getRange(rowNum, h[P.UPDATED_BY]).setValue(o.by);
      sh.getRange(rowNum, h[P.NEW]).setValue(o.ns);
      if (o.pn !== '' && o.pn !== undefined && o.pn !== null) sh.getRange(rowNum, h[P.PROGRESS]).setValue(o.pn);
      sh.getRange(rowNum, h[P.NOTES]).setValue(Util.str(o.notes));
      sh.getRange(rowNum, h[P.DELAY]).setValue(Util.str(o.delay));
      if (Util.isDate(o.revised)) sh.getRange(rowNum, h[P.REVISED]).setValue(o.revised);
    });
    const result = Progress.run();
    const mine = result.errors.filter(function (e) { return e.row === rowNum; })[0];
    if (mine) {
      const sh2 = Util.sheet(CONFIG.SHEETS.PROGRESS);
      sh2.getRange(rowNum, 1, 1, sh2.getLastColumn()).clearContent();
    }
    return { result: result, mine: mine, rowNum: rowNum };
  }
};

// ============================================================================
//  Evidence  (EVIDENCE LOG)
// ============================================================================
const Evidence = {
  run() {
    Util.assertTimezone();
    let result = null;
    Util.withLock(function () { result = Evidence.process_(new Date()); });
    return result;
  },

  process_(now) {
    const E_ = CONFIG.EV_H, H = CONFIG.DB_H;
    const db = Db.open(); const dbData = Db.readAll(db);
    const sh = Util.sheet(CONFIG.SHEETS.EVIDENCE);
    const h = Util.headerMap(sh, 1);
    Util.requireHeaders(h, Object.keys(E_).map(function (k) { return E_[k]; }), CONFIG.SHEETS.EVIDENCE);
    const last = Math.max(Util.lastNonEmptyRow(sh, h[E_.ID], 2), Util.lastNonEmptyRow(sh, h[E_.TASK_ID], 2), Util.lastNonEmptyRow(sh, h[E_.URL], 2));
    const result = { processed: [], errors: [], skipped: 0 };
    if (last < 2) return result;
    const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const existing = vals.map(function (r) { return r[h[E_.ID] - 1]; });
    const seen = {};
    vals.forEach(function (r) { if (Util.str(r[h[E_.ID] - 1]) !== '') seen[Util.up(r[h[E_.TASK_ID] - 1]) + '|' + Util.str(r[h[E_.URL] - 1])] = true; });
    const cands = [];
    vals.forEach(function (r, i) {
      if (Util.str(r[h[E_.ID] - 1]) !== '') { result.skipped++; return; }
      if (Util.str(r[h[E_.TASK_ID] - 1]) !== '' || Util.str(r[h[E_.URL] - 1]) !== '' || Util.str(r[h[E_.DESC] - 1]) !== '') cands.push({ rowNum: i + 2, r: r });
    });
    if (!cands.length) return result;
    const res = Ids.reserve(CONFIG.PROP.EVD_SEQ, existing, CONFIG.EVD_PREFIX, cands.length);
    let seq = res.first, consumed = 0;
    const links = {};

    cands.forEach(function (cd) {
      const r = cd.r; const E = [];
      const taskId = Util.up(r[h[E_.TASK_ID] - 1]); const url = Util.str(r[h[E_.URL] - 1]); const by = Util.str(r[h[E_.BY] - 1]);
      const t = taskId ? dbData.byId[taskId] : null;
      if (!taskId) E.push('Task ID wajib diisi'); else if (!t) E.push('Task ID "' + taskId + '" tidak ada di TASK DATABASE');
      if (!url) E.push('Google Drive URL wajib diisi'); else if (!/^https?:\/\//i.test(url)) E.push('URL harus diawali http:// atau https://');
      if (!by) E.push('Uploaded By wajib diisi');
      if (!E.length && seen[taskId + '|' + url]) E.push('Evidence yang sama sudah tercatat untuk tugas ini');
      if (E.length) { sh.getRange(cd.rowNum, h[E_.NOTE]).setValue('ERROR: ' + E.join('; ')); result.errors.push({ row: cd.rowNum, messages: E }); return; }
      const id = Ids.format(CONFIG.EVD_PREFIX, seq++, 5); consumed++;
      sh.getRange(cd.rowNum, h[E_.ID]).setValue(id);
      if (Util.str(r[h[E_.DATE] - 1]) === '') { sh.getRange(cd.rowNum, h[E_.DATE]).setValue(now); sh.getRange(cd.rowNum, h[E_.DATE]).setNumberFormat(CONFIG.FORMATS.DATETIME); }
      sh.getRange(cd.rowNum, h[E_.NOTE]).setValue('OK');
      seen[taskId + '|' + url] = true;
      links[t._row] = url;
      result.processed.push({ row: cd.rowNum, id: id, taskId: t[H.TASK_ID], url: url });
    });
    Object.keys(links).forEach(function (row) { db.sheet.getRange(Number(row), db.h[H.EVIDENCE]).setValue(links[row]); });
    if (consumed) PropertiesService.getScriptProperties().setProperty(CONFIG.PROP.EVD_SEQ, String(res.first + consumed - 1));
    return result;
  },

  summaryText(r) {
    const L = ['Evidence diproses: ' + r.processed.length];
    if (r.errors.length) { L.push('ERROR: ' + r.errors.length + ' baris'); r.errors.slice(0, CONFIG.MAX_ERRORS_IN_ALERT).forEach(function (e) { L.push('  Baris ' + e.row + ': ' + e.messages.join('; ')); }); }
    return L.join('\n');
  },

  /** Dipanggil trigger Google Form (upload foto langsung ke Drive). */
  fromForm(nv) {
    const F = CONFIG.FORMS.EVIDENCE, E_ = CONFIG.EV_H;
    const first = function (k) { return nv[k] && nv[k].length ? Util.str(nv[k][0]) : ''; };
    const urls = first(F.FILES).split(/\s*,\s*/).filter(Boolean);
    const sh = Util.sheet(CONFIG.SHEETS.EVIDENCE); const h = Util.headerMap(sh, 1);
    Util.withLock(function () {
      let row = Math.max(Util.lastNonEmptyRow(sh, h[E_.ID], 2), Util.lastNonEmptyRow(sh, h[E_.TASK_ID], 2), Util.lastNonEmptyRow(sh, h[E_.URL], 2)) + 1;
      urls.forEach(function (u) {
        sh.getRange(row, h[E_.TASK_ID]).setValue(first(F.TASK_ID));
        sh.getRange(row, h[E_.BY]).setValue(first(F.UPLOADED_BY));
        sh.getRange(row, h[E_.DESC]).setValue(first(F.DESCRIPTION));
        sh.getRange(row, h[E_.URL]).setValue(u);
        row++;
      });
    });
    return Evidence.run();
  }
};

// ============================================================================
//  Overdue  (deteksi overdue & mendekati deadline; sumber kebenaran = TASK DATABASE)
// ============================================================================
const Overdue = {
  /** Pure function: klasifikasi satu tugas. */
  classify(status, deadline, now, dueSoonDays) {
    const s = Util.up(status);
    if (s === CONFIG.STATUS.COMPLETED || s === CONFIG.STATUS.CANCELLED) return 'CLOSED';
    if (!Util.isDate(deadline)) return 'NO_DEADLINE';
    if (deadline.getTime() < now.getTime()) return 'OVERDUE';
    if (deadline.getTime() - now.getTime() <= dueSoonDays * 86400000) return 'DUE_SOON';
    return 'OK';
  },

  detect(now, S) {
    const H = CONFIG.DB_H;
    S = S || Settings.load();
    const days = Settings.num(S, 'DUE_SOON_DAYS');
    const db = Db.open(); const data = Db.readAll(db);
    const out = { overdue: [], dueSoon: [], total: data.rows.length };
    data.rows.forEach(function (o) {
      const k = Overdue.classify(o[H.STATUS], o[H.DEADLINE], now, days);
      if (k !== 'OVERDUE' && k !== 'DUE_SOON') return;
      const item = { id: o[H.TASK_ID], title: o[H.TITLE], project: o[H.PROJECT_NAME], pic: o[H.ASSIGNED_TO], status: Util.up(o[H.STATUS]),
                     deadline: o[H.DEADLINE], createdBy: o[H.CREATED_BY], priority: o[H.PRIORITY],
                     days: Math.round(Math.abs(now.getTime() - o[H.DEADLINE].getTime()) / 8640000) / 10 };
      (k === 'OVERDUE' ? out.overdue : out.dueSoon).push(item);
    });
    const byDl = function (a, b) { return a.deadline.getTime() - b.deadline.getTime(); };
    out.overdue.sort(byDl); out.dueSoon.sort(byDl);
    return out;
  }
};

// ============================================================================
//  Notify  (email)
// ============================================================================
const Notify = {
  enabled(S) { return Settings.yes(S, 'EMAIL_ENABLED'); },
  validEmail(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(Util.str(e)); },

  /** Kirim email. Selalu mencatat hasil (berhasil/gagal + alasan) ke SYSTEM LOG. */
  send(S, to, cc, subject, html) {
    if (!Notify.validEmail(to)) {
      Log.audit('WARN', 'MAIL', 'GAGAL ke "' + to + '": email tujuan tidak valid / kosong (' + subject + ')');
      return { ok: false, error: 'email tujuan tidak valid: "' + to + '"' };
    }
    try {
      if (MailApp.getRemainingDailyQuota() < 1) {
        Log.audit('ERROR', 'MAIL', 'GAGAL ke ' + to + ': kuota email harian habis');
        return { ok: false, error: 'kuota email harian habis' };
      }
      const opt = { to: to, subject: '[' + S.cfg.SYSTEM_NAME + '] ' + subject, body: Notify.plain(html), htmlBody: html, name: S.cfg.SYSTEM_NAME };
      if (cc) opt.cc = cc;
      MailApp.sendEmail(opt);
      Log.audit('INFO', 'MAIL', 'TERKIRIM ke ' + to + (cc ? ' (cc ' + cc + ')' : '') + ': ' + subject);
      return { ok: true };
    } catch (e) {
      Log.audit('ERROR', 'MAIL', 'GAGAL ke ' + to + ': ' + String(e.message || e));
      return { ok: false, error: String(e.message || e) };
    }
  },

  plain(html) {
    return String(html).replace(/<a\s[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, '$2 ($1)').replace(/<\/(p|tr|table)>/gi, '\n').replace(/<\/t[dh]>/gi, ' | ').replace(/<[^>]+>/g, '')
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&rarr;/g, '->');
  },

  /** Umpan balik ke pengirim Form HP yang ditolak (tidak ada UI lain untuk menampilkan error). Tidak bergantung EMAIL_ENABLED. */
  rejectToSender(email, subject, messages) {
    if (!Notify.validEmail(email)) return { ok: false, error: 'email pengirim tidak terbaca' };
    const S = Settings.load();
    const html = '<p><b>' + Util.esc(subject) + '</b></p><ul>' + messages.map(function (m) { return '<li>' + Util.esc(m) + '</li>'; }).join('') +
      '</ul><p style="font-size:12px;color:#666">Perbaiki lalu kirim ulang form.</p>';
    return Notify.send(S, email, '', subject, html);
  },

  /** Sel tabel berisi ID tugas yang menjadi link ke detail tugas di web app (teks biasa bila web app tidak aktif). */
  idCell(S, id) { const a = Web.a(S, id, { tab: 'tasks', task: id }); return a ? { html: a } : id; },
  /** Paragraf link; '' bila web app tidak aktif. */
  linkP(S, label, params, note) {
    const a = Web.a(S, label, params);
    return a ? '<p>' + a + (note ? ' <span style="font-size:12px;color:#666">' + Util.esc(note) + '</span>' : '') + '</p>' : '';
  },

  table(headers, rows) {
    const th = headers.map(function (x) { return '<th style="background:#1F2937;color:#fff;padding:4px 8px;text-align:left;font-size:12px">' + Util.esc(x) + '</th>'; }).join('');
    const tr = rows.map(function (r) {
      return '<tr>' + r.map(function (c) { return '<td style="border-bottom:1px solid #ddd;padding:4px 8px;font-size:12px">' + (c && typeof c === 'object' && 'html' in c ? c.html : Util.esc(c)) + '</td>'; }).join('') + '</tr>';
    }).join('');
    return '<table style="border-collapse:collapse;font-family:Arial,sans-serif"><tr>' + th + '</tr>' + tr + '</table>';
  },

  onSubmit(result) {
    const S = Settings.load();
    const out = { sent: 0, failed: [], skipped: '' };
    if (!Notify.enabled(S)) { out.skipped = 'EMAIL_ENABLED = NO di SETTINGS (ubah ke YES)'; Log.audit('WARN', 'MAIL', 'Email penugasan dilewati: EMAIL_ENABLED = NO'); return out; }
    if (!Settings.yes(S, 'EMAIL_NOTIFY_ON_SUBMIT')) { out.skipped = 'EMAIL_NOTIFY_ON_SUBMIT = NO di SETTINGS'; return out; }
    const groups = {};
    result.submitted.forEach(function (t) { (groups[t.pic] = groups[t.pic] || []).push(t); });
    Object.keys(groups).forEach(function (pic) {
      const list = groups[pic];
      const cc = Settings.yes(S, 'EMAIL_CC_ADMINS')
        ? list.map(function (t) { return t.adminEmail; }).filter(function (e, i, a) { return Notify.validEmail(e) && a.indexOf(e) === i; }).join(',') : '';
      const html = '<p>Halo ' + Util.esc(pic) + ', Anda mendapat ' + list.length + ' tugas baru:</p>' +
        Notify.table(['Task ID', 'Tugas', 'Proyek', 'Prioritas', 'Deadline', 'Pemberi tugas'],
          list.map(function (t) { return [Notify.idCell(S, t.id), t.title, t.project, t.priority, Util.fmt(t.deadline), t.createdBy]; })) +
        Notify.linkP(S, 'Buka tugas & laporan saya', { tab: 'report', staff: pic }, '(masuk dengan kode akses Anda; upload foto sebelum/sesudah dari sini)') +
        '<p style="font-size:12px;color:#666">Update progres melalui web app di atas atau sheet PROGRESS LOG.</p>';
      const r = Notify.send(S, list[0].picEmail, cc, list.length + ' tugas baru untuk ' + pic, html);
      if (r.ok) out.sent++; else out.failed.push(pic + ': ' + r.error);
    });
    return out;
  },

  onProgress(result) {
    const S = Settings.load();
    if (!Notify.enabled(S) || !Settings.yes(S, 'EMAIL_NOTIFY_ON_PROGRESS')) return;
    const interesting = [CONFIG.STATUS.COMPLETED, CONFIG.STATUS.PENDING, CONFIG.STATUS.CANCELLED];
    result.processed.filter(function (p) { return interesting.indexOf(p.status) >= 0; }).forEach(function (p) {
      const adm = S.admins.filter(function (a) { return Util.low(a.name) === Util.low(p.createdBy); })[0];
      if (!adm) return;
      const html = '<p>' + Util.esc(p.by) + ' memperbarui tugas <b>' + Util.esc(p.taskId) + '</b> - ' + Util.esc(p.title) + '</p>' +
        '<p>Status: ' + Util.esc(p.prev) + ' &rarr; <b>' + Util.esc(p.status) + '</b> (' + Math.round(p.progress * 100) + '%)</p>' +
        (p.delay ? '<p>Alasan keterlambatan: ' + Util.esc(p.delay) + '</p>' : '') + (p.notes ? '<p>Catatan: ' + Util.esc(p.notes) + '</p>' : '') +
        Notify.linkP(S, 'Lihat tugas & foto bukti', { tab: 'tasks', task: p.taskId }) +
        Notify.linkP(S, 'Laporan ' + p.by, { tab: 'report', staff: p.by });
      const r = Notify.send(S, adm.email, '', p.taskId + ' -> ' + p.status, html);
      if (!r.ok) Log.info('Notify.onProgress', 'gagal kirim ke ' + adm.email + ': ' + r.error);
    });
  },

  digest(S, det, now) {
    const out = { sent: 0, failed: [] };
    const byPic = {};
    det.overdue.concat(det.dueSoon).forEach(function (t) { (byPic[t.pic] = byPic[t.pic] || { overdue: [], dueSoon: [] }); });
    det.overdue.forEach(function (t) { byPic[t.pic].overdue.push(t); });
    det.dueSoon.forEach(function (t) { byPic[t.pic].dueSoon.push(t); });
    const staff = Masters.staff(S);
    const rowsOf = function (a, lbl) { return a.map(function (t) { return [Notify.idCell(S, t.id), t.title, t.project, Util.fmt(t.deadline), t.days + ' hari ' + lbl]; }); };
    Object.keys(byPic).forEach(function (pic) {
      const s = staff.byName[Util.low(pic)];
      if (!s || !s.active) return;
      const g = byPic[pic]; let html = '<p>Halo ' + Util.esc(pic) + ',</p>';
      if (g.overdue.length) html += '<p><b style="color:#C00000">OVERDUE (' + g.overdue.length + ')</b></p>' + Notify.table(['Task ID', 'Tugas', 'Proyek', 'Deadline', 'Terlambat'], rowsOf(g.overdue, 'terlambat'));
      if (g.dueSoon.length) html += '<p><b>Mendekati deadline (' + g.dueSoon.length + ')</b></p>' + Notify.table(['Task ID', 'Tugas', 'Proyek', 'Deadline', 'Sisa'], rowsOf(g.dueSoon, 'lagi'));
      html += Notify.linkP(S, 'Buka tugas & laporan saya', { tab: 'report', staff: pic });
      const r = Notify.send(S, s.email, '', 'Reminder: ' + g.overdue.length + ' overdue, ' + g.dueSoon.length + ' mendekati deadline', html);
      if (r.ok) out.sent++; else out.failed.push(pic + ': ' + r.error);
    });
    if (Settings.yes(S, 'REMINDER_ADMIN_DIGEST') && (det.overdue.length || det.dueSoon.length)) {
      const all = det.overdue.map(function (t) { return [Notify.idCell(S, t.id), t.title, t.pic, t.project, Util.fmt(t.deadline), 'OVERDUE ' + t.days + ' hari']; })
        .concat(det.dueSoon.map(function (t) { return [Notify.idCell(S, t.id), t.title, t.pic, t.project, Util.fmt(t.deadline), 'sisa ' + t.days + ' hari']; }));
      const html = '<p>Ringkasan harian ' + Util.fmt(now, 'dd/MM/yyyy') + ': ' + det.overdue.length + ' overdue, ' + det.dueSoon.length + ' mendekati deadline.</p>' +
        Notify.table(['Task ID', 'Tugas', 'PIC', 'Proyek', 'Deadline', 'Status'], all) +
        Notify.linkP(S, 'Buka monitoring', { tab: 'monitor' });
      S.admins.forEach(function (a) {
        const r = Notify.send(S, a.email, '', 'Ringkasan harian: ' + det.overdue.length + ' overdue', html);
        if (r.ok) out.sent++; else out.failed.push(a.name + ': ' + r.error);
      });
    }
    return out;
  }
};

// ============================================================================
//  Reminders  (job harian)
// ============================================================================
const Reminders = {
  run(force) {
    const props = PropertiesService.getScriptProperties();
    const now = new Date();
    const S = Settings.load();
    const today = Util.fmt(now, 'yyyy-MM-dd');
    const res = { ran: false, reason: '', overdue: 0, dueSoon: 0, mail: null };
    if (!force && !Settings.yes(S, 'REMINDER_ENABLED')) { res.reason = 'REMINDER_ENABLED = NO'; return res; }
    if (!force && props.getProperty(CONFIG.PROP.LAST_DAILY_RUN) === today) { res.reason = 'sudah berjalan hari ini'; return res; }
    const det = Overdue.detect(now, S);
    res.overdue = det.overdue.length; res.dueSoon = det.dueSoon.length; res.ran = true; res.det = det;
    if (!Notify.enabled(S)) { res.reason = 'EMAIL_ENABLED = NO (deteksi saja, email tidak dikirim)'; return res; }
    res.mail = Notify.digest(S, det, now);
    props.setProperty(CONFIG.PROP.LAST_DAILY_RUN, today);
    return res;
  }
};

// ============================================================================
//  Forms  (input via HP: Google Form tugas + Google Form progres)
//  Script membuat form, mengisi dropdown dari STAFF/PROJECT MASTER & SETTINGS,
//  dan memproses setiap kiriman. Identitas pengirim = email Google yang login.
// ============================================================================
const Forms = {
  props() { return PropertiesService.getScriptProperties(); },

  open_(key) {
    const id = Forms.props().getProperty(key);
    if (!id) return null;
    try { return FormApp.openById(id); }
    catch (e) { Log.audit('WARN', 'FORMS', 'Form (' + key + ') tidak bisa dibuka: ' + String(e.message || e)); return null; }
  },

  configure_(form, description, confirmation) {
    form.setDescription(description);
    form.setConfirmationMessage(confirmation);
    form.setLimitOneResponsePerUser(false);
    form.setShowLinkToRespondAgain(true);   // tombol "Kirim tanggapan lain" -> input banyak tugas cepat
    form.setAllowResponseEdits(false);
    form.setProgressBar(false);
    try {
      if (FormApp.EmailCollectionType && typeof form.setEmailCollectionType === 'function') form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      else form.setCollectEmail(true);
    } catch (e) { form.setCollectEmail(true); }
    try { form.setRequireLogin(true); } catch (e2) { /* akun non-Workspace: login tetap diminta oleh email terverifikasi */ }
  },

  setChoices_(form, title, values) {
    const it = form.getItems().filter(function (i) { return i.getTitle() === title; })[0];
    if (!it) return false;
    const vals = values.length ? values : ['(belum ada data)'];
    if (it.getType() === FormApp.ItemType.LIST) it.asListItem().setChoiceValues(vals);
    else it.asMultipleChoiceItem().setChoiceValues(vals);
    return true;
  },

  activeNames_(m) { return m.list.filter(function (x) { return x.active; }).map(function (x) { return x.name; }); },

  openTaskChoices_() {
    const H = CONFIG.DB_H, ST = CONFIG.STATUS;
    const rows = Db.readAll(Db.open()).rows.filter(function (o) {
      const s = Util.up(o[H.STATUS]); return s !== ST.COMPLETED && s !== ST.CANCELLED;
    });
    rows.sort(function (a, b) {
      const x = Util.isDate(a[H.DEADLINE]) ? a[H.DEADLINE].getTime() : 9e15, y = Util.isDate(b[H.DEADLINE]) ? b[H.DEADLINE].getTime() : 9e15;
      return x - y;
    });
    return rows.slice(0, CONFIG.FORMS.MAX_TASK_CHOICES).map(function (o) {
      const t = Util.str(o[H.TITLE]); return o[H.TASK_ID] + ' | ' + (t.length > 60 ? t.substring(0, 57) + '...' : t) + ' | ' + o[H.ASSIGNED_TO];
    });
  },

  /** Buat form (bila belum ada), isi dropdown, dan pasang trigger. Aman dijalankan berulang. */
  setup() {
    Util.assertTimezone();
    const S = Settings.load();
    const staff = Forms.activeNames_(Masters.staff(S)), projects = Forms.activeNames_(Masters.projects(S));
    if (!projects.length) throw new Error('Tidak ada proyek Active di PROJECT MASTER.');
    if (!staff.length) throw new Error('Tidak ada staf Active di STAFF MASTER.');
    if (!S.admins.some(function (a) { return Notify.validEmail(a.email); })) throw new Error('Isi minimal satu Admin Email yang valid di SETTINGS.');
    const T = CONFIG.FORMS.TASK, P = CONFIG.FORMS.PROGRESS;
    const created = [];

    if (!Forms.open_(CONFIG.PROP.FORM_TASK)) {
      const f = FormApp.create(CONFIG.FORMS.TITLE_TASK + ' - ' + S.cfg.SYSTEM_NAME);
      Forms.configure_(f, 'Khusus admin. Login dengan akun Google yang emailnya terdaftar di SETTINGS > Admin Email.',
        'Tugas diterima. Cek email/sheet untuk Task ID. Tekan "Kirim tanggapan lain" untuk menambah tugas berikutnya.');
      f.addListItem().setTitle(T.PROJECT).setRequired(true).setChoiceValues(projects);
      f.addListItem().setTitle(T.PIC).setRequired(true).setChoiceValues(staff);
      f.addTextItem().setTitle(T.TITLE).setRequired(true);
      f.addParagraphTextItem().setTitle(T.DESC);
      f.addMultipleChoiceItem().setTitle(T.PRIORITY).setRequired(true).setChoiceValues(S.priorities);
      f.addDateItem().setTitle(T.DEADLINE).setRequired(true).setHelpText('Batas akhir = akhir hari (23:59) pada tanggal ini.');
      f.addDateItem().setTitle(T.START).setHelpText('Kosongkan = hari ini.');
      f.addTextItem().setTitle(T.PARENT).setHelpText('Isi Task ID induk (mis. TSK-00012) bila ini sub-tugas.');
      f.addParagraphTextItem().setTitle(T.NOTES);
      Forms.props().setProperty(CONFIG.PROP.FORM_TASK, f.getId());
      created.push('Form Tugas');
    }
    if (!Forms.open_(CONFIG.PROP.FORM_PROGRESS)) {
      const f = FormApp.create(CONFIG.FORMS.TITLE_PROGRESS + ' - ' + S.cfg.SYSTEM_NAME);
      Forms.configure_(f, 'Login dengan akun Google yang emailnya terdaftar di STAFF MASTER (hanya tugas milik Anda) atau admin.',
        'Update progres tercatat. Jika ditolak, alasan dikirim ke email Anda.');
      f.addListItem().setTitle(P.TASK).setRequired(true).setChoiceValues(['(belum ada data)']);
      f.addMultipleChoiceItem().setTitle(P.STATUS).setRequired(true).setChoiceValues(['IN PROGRESS']);
      const pc = [CONFIG.FORMS.KEEP_PROGRESS]; for (let i = 0; i <= 90; i += 10) pc.push(i + '%');
      f.addListItem().setTitle(P.PROGRESS).setChoiceValues(pc).setHelpText('COMPLETED otomatis 100%.');
      f.addParagraphTextItem().setTitle(P.NOTES);
      f.addParagraphTextItem().setTitle(P.DELAY).setHelpText('Wajib bila status PENDING atau tugas sudah lewat deadline.');
      f.addDateItem().setTitle(P.REVISED);
      Forms.props().setProperty(CONFIG.PROP.FORM_PROGRESS, f.getId());
      created.push('Form Progres');
    }
    Forms.refresh();
    Triggers.removeByHandler_('onTaskFormSubmit'); Triggers.removeByHandler_('onProgressFormSubmit');
    ScriptApp.newTrigger('onTaskFormSubmit').forForm(Forms.open_(CONFIG.PROP.FORM_TASK)).onFormSubmit().create();
    ScriptApp.newTrigger('onProgressFormSubmit').forForm(Forms.open_(CONFIG.PROP.FORM_PROGRESS)).onFormSubmit().create();
    Log.audit('INFO', 'FORMS', 'Setup form selesai. Dibuat baru: ' + (created.join(', ') || '-'));
    return { created: created, links: Forms.links() };
  },

  /** Perbarui dropdown form dari master data / SETTINGS / tugas terbuka. */
  refresh() {
    const S = Settings.load(); const out = { task: false, progress: false };
    const tf = Forms.open_(CONFIG.PROP.FORM_TASK);
    if (tf) {
      const T = CONFIG.FORMS.TASK;
      Forms.setChoices_(tf, T.PROJECT, Forms.activeNames_(Masters.projects(S)));
      Forms.setChoices_(tf, T.PIC, Forms.activeNames_(Masters.staff(S)));
      Forms.setChoices_(tf, T.PRIORITY, S.priorities);
      out.task = true;
    }
    const pf = Forms.open_(CONFIG.PROP.FORM_PROGRESS);
    if (pf) {
      const P = CONFIG.FORMS.PROGRESS;
      Forms.setChoices_(pf, P.TASK, Forms.openTaskChoices_());
      Forms.setChoices_(pf, P.STATUS, S.statuses.filter(function (x) { return x !== CONFIG.STATUS.NEW; }));
      out.progress = true;
    }
    return out;
  },

  /** Seperti refresh() tetapi tidak pernah melempar error dan tidak berbuat apa-apa bila form belum dibuat. */
  refreshQuiet() {
    try {
      const p = Forms.props();
      if (!p.getProperty(CONFIG.PROP.FORM_TASK) && !p.getProperty(CONFIG.PROP.FORM_PROGRESS)) return;
      Forms.refresh();
    } catch (e) { Log.audit('WARN', 'FORMS', 'Refresh pilihan form gagal: ' + String(e.message || e)); }
  },

  links() {
    const o = {}; const tf = Forms.open_(CONFIG.PROP.FORM_TASK), pf = Forms.open_(CONFIG.PROP.FORM_PROGRESS);
    if (tf) { o.task = tf.getPublishedUrl(); o.taskEdit = tf.getEditUrl(); }
    if (pf) { o.progress = pf.getPublishedUrl(); o.progressEdit = pf.getEditUrl(); }
    return o;
  },

  answers_(response) {
    const m = {};
    response.getItemResponses().forEach(function (ir) {
      let v = ir.getResponse(); if (Array.isArray(v)) v = v.join(', ');
      m[ir.getItem().getTitle()] = v;
    });
    return m;
  }
};

// ============================================================================
//  Guard  (cegah edit baris yang sudah SUBMITTED / sudah punya Log ID)
// ============================================================================
const Guard = {
  onEdit(e) {
    if (!e || !e.range) return;
    const sh = e.range.getSheet(); const name = sh.getName();
    if (name === CONFIG.SHEETS.INPUT) Guard.input_(e, sh);
    else if (name === CONFIG.SHEETS.PROGRESS || name === CONFIG.SHEETS.EVIDENCE) Guard.log_(e, sh, name);
  },

  revert_(e, msg) {
    const single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
    if (single) {
      if (e.oldValue === undefined) e.range.clearContent(); else e.range.setValue(e.oldValue);
      Util.toast(msg + ' Perubahan dibatalkan.', 'Task System');
    } else {
      Util.toast(msg + ' Edit/paste ini mengenai baris terkunci - tekan Ctrl+Z untuk membatalkan.', 'Task System');
    }
  },

  input_(e, sh) {
    const IN = CONFIG.INPUT, C = IN.COL;
    const r1 = e.range.getRow(), r2 = r1 + e.range.getNumRows() - 1;
    const c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
    if (r2 < IN.FIRST_ROW || r1 > IN.LAST_ROW || c1 > C.STATUS) return;
    const from = Math.max(r1, IN.FIRST_ROW), to = Math.min(r2, IN.LAST_ROW);
    const st = sh.getRange(from, C.STATUS, to - from + 1, 1).getValues();
    const single = from === to && c1 === c2;
    for (let i = 0; i < st.length; i++) {
      let locked = Util.up(st[i][0]) === IN.SUBMITTED;
      if (single && c1 === C.STATUS) locked = Util.up(e.oldValue) === IN.SUBMITTED && Util.up(e.value) !== IN.SUBMITTED;
      if (single && c1 === C.STATUS && locked) { e.range.setValue(e.oldValue); Util.toast('Status SUBMITTED tidak boleh diubah.', 'Task System'); return; }
      if (locked) { Guard.revert_(e, 'Baris ' + (from + i) + ' sudah SUBMITTED dan terkunci.'); return; }
    }
  },

  log_(e, sh, name) {
    const isP = name === CONFIG.SHEETS.PROGRESS;
    const r1 = e.range.getRow(), nr = e.range.getNumRows();
    if (r1 + nr - 1 < 2) return;
    const from = Math.max(r1, 2);
    const idCol = 1;  // Log ID / Evidence ID selalu kolom A
    const ids = sh.getRange(from, idCol, r1 + nr - from, 1).getValues();
    const single = nr === 1 && e.range.getNumColumns() === 1;
    for (let i = 0; i < ids.length; i++) {
      let had = Util.str(ids[i][0]) !== '';
      if (single && e.range.getColumn() === idCol) had = Util.str(e.oldValue) !== '';
      if (had) { Guard.revert_(e, 'Baris ' + (from + i) + ' sudah tercatat (' + (isP ? 'Log ID' : 'Evidence ID') + ') dan tidak boleh diubah.'); return; }
    }
  }
};

// ============================================================================
//  Triggers
// ============================================================================
const Triggers = {
  install() {
    const S = Settings.load();
    const hour = Math.min(23, Math.max(0, Math.floor(Settings.num(S, 'REMINDER_HOUR'))));
    Triggers.removeByHandler_('dailyJob');
    ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(hour).inTimezone(Util.tz()).create();
    return 'Trigger harian dipasang: setiap hari sekitar jam ' + Util.pad(hour, 2) + ':00 (' + Util.tz() + ').';
  },
  installForm() {
    Triggers.removeByHandler_('onFormSubmitRouter');
    ScriptApp.newTrigger('onFormSubmitRouter').forSpreadsheet(Util.ss()).onFormSubmit().create();
    return 'Trigger Google Form dipasang (onFormSubmitRouter).';
  },
  remove() {
    let n = 0;
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (['dailyJob', 'onFormSubmitRouter', 'onTaskFormSubmit', 'onProgressFormSubmit'].indexOf(t.getHandlerFunction()) >= 0) { ScriptApp.deleteTrigger(t); n++; }
    });
    return n + ' trigger dihapus.';
  },
  removeByHandler_(name) {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === name) ScriptApp.deleteTrigger(t); });
  },
  has(name) { return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === name; }); }
};

// ============================================================================
//  Protections (opsional): peringatan bila kolom otomatis diedit
// ============================================================================
const Protections = {
  apply() {
    const ss = Util.ss(); const tag = CONFIG.PROTECT_TAG; let n = 0;
    ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).concat(ss.getProtections(SpreadsheetApp.ProtectionType.SHEET))
      .forEach(function (p) { if (Util.str(p.getDescription()).indexOf(tag) === 0) p.remove(); });
    const warn = function (range, desc) { const p = range.protect().setDescription(tag + ' ' + desc); p.setWarningOnly(true); n++; };
    const IN = CONFIG.INPUT;
    const inp = Util.sheet(CONFIG.SHEETS.INPUT);
    [1, 2, CONFIG.INPUT.COL.STATUS, CONFIG.INPUT.COL.MESSAGE].forEach(function (c) { warn(inp.getRange(IN.FIRST_ROW, c, IN.LAST_ROW - IN.FIRST_ROW + 1, 1), 'TASK INPUT kolom otomatis'); });
    const db = Db.open(); const H = CONFIG.DB_H;
    const p = db.sheet.protect().setDescription(tag + ' TASK DATABASE'); p.setWarningOnly(true); n++;
    p.setUnprotectedRanges([db.sheet.getRange(2, db.h[H.DELAY_CATEGORY], CONFIG.DB_MAX_ROWS, 2)]);
    const pl = Util.sheet(CONFIG.SHEETS.PROGRESS);
    [1, 2, 5, 11].forEach(function (c) { warn(pl.getRange(2, c, 2000, 1), 'PROGRESS LOG kolom otomatis'); });
    const ev = Util.sheet(CONFIG.SHEETS.EVIDENCE);
    [1, 3, 7].forEach(function (c) { warn(ev.getRange(2, c, 2000, 1), 'EVIDENCE LOG kolom otomatis'); });
    const rep = function (name, unprot) {
      const sh = Util.sheet(name); const pp = sh.protect().setDescription(tag + ' ' + name); pp.setWarningOnly(true); n++;
      pp.setUnprotectedRanges(unprot.map(function (a) { return sh.getRange(a); }));
    };
    rep(CONFIG.SHEETS.DASHBOARD, ['C4:C8']); rep(CONFIG.SHEETS.STAFF_PERF, ['C4:C5']); rep(CONFIG.SHEETS.PROJECT_REPORT, ['C27']);
    return n + ' proteksi (mode peringatan) dipasang.';
  }
};

// ============================================================================
//  SelfTest  (jangan mengasumsikan script berfungsi - jalankan ini)
// ============================================================================
const SelfTest = {
  run() {
    const out = []; const add = function (name, ok, detail) { out.push({ name: name, ok: !!ok, detail: detail || '' }); };
    const ss = Util.ss();
    Object.keys(CONFIG.SHEETS).forEach(function (k) { add('Sheet ' + CONFIG.SHEETS[k], !!ss.getSheetByName(CONFIG.SHEETS[k])); });
    const order = ss.getSheets().map(function (s) { return s.getName(); });
    add('Sheet pertama = TASK INPUT', order[0] === CONFIG.SHEETS.INPUT, order[0]);
    try { Util.assertTimezone(); add('Zona waktu spreadsheet = script', true, Util.tz()); } catch (e) { add('Zona waktu spreadsheet = script', false, e.message); }
    let S = null;
    try {
      S = Settings.load();
      add('SETTINGS: priority terisi', S.priorities.length >= 1, S.priorities.join(', '));
      add('SETTINGS: status inti lengkap', ['NEW', 'IN PROGRESS', 'PENDING', 'COMPLETED', 'CANCELLED'].every(function (x) { return S.statuses.indexOf(x) >= 0; }), S.statuses.join(', '));
      add('SETTINGS: admin terisi', S.admins.length >= 1, S.admins.map(function (a) { return a.name + ' <' + a.email + '>'; }).join('; '));
      add('SETTINGS: delay category terisi', S.delayCategories.length >= 1);
      add('SETTINGS: EMAIL_ENABLED', true, S.cfg.EMAIL_ENABLED);
    } catch (e) { add('SETTINGS terbaca', false, e.message); }
    try {
      const db = Db.open(); add('TASK DATABASE: header lengkap', true);
      const d = Db.readAll(db);
      const seen = {}; let dup = 0; d.rows.forEach(function (o) { const k = Util.up(o[CONFIG.DB_H.TASK_ID]); if (seen[k]) dup++; seen[k] = 1; });
      add('TASK DATABASE: Task ID unik', dup === 0, d.rows.length + ' tugas, duplikat: ' + dup);
      add('TASK DATABASE: rumus kolom otomatis ada', !!Util.str(db.sheet.getRange(2, db.autoStart).getFormula()));
    } catch (e) { add('TASK DATABASE', false, e.message); }
    ['PROGRESS', 'EVIDENCE'].forEach(function (k) {
      try {
        const sh = Util.sheet(CONFIG.SHEETS[k]); const h = Util.headerMap(sh, 1);
        const H2 = k === 'PROGRESS' ? CONFIG.PL_H : CONFIG.EV_H;
        Util.requireHeaders(h, Object.keys(H2).map(function (x) { return H2[x]; }), CONFIG.SHEETS[k]); add(CONFIG.SHEETS[k] + ': header lengkap', true);
      } catch (e) { add(CONFIG.SHEETS[k] + ': header', false, e.message); }
    });
    try {
      const sh = Util.sheet(CONFIG.SHEETS.INPUT);
      add('TASK INPUT: header baris 4', Util.str(sh.getRange(CONFIG.INPUT.HEADER_ROW, CONFIG.INPUT.COL.TITLE).getValue()) === 'Task Title' &&
        Util.str(sh.getRange(CONFIG.INPUT.HEADER_ROW, CONFIG.INPUT.COL.STATUS).getValue()) === 'Submission Status');
    } catch (e) { add('TASK INPUT', false, e.message); }
    if (S) {
      try {
        const st = Masters.staff(S), pj = Masters.projects(S);
        add('STAFF MASTER: ada staf Active', st.list.some(function (x) { return x.active; }), st.list.length + ' staf');
        add('PROJECT MASTER: ada proyek Active', pj.list.some(function (x) { return x.active; }), pj.list.length + ' proyek');
        const bad = st.list.filter(function (x) { return x.active && !Notify.validEmail(x.email); }).map(function (x) { return x.name; });
        add('Email staf Active valid', bad.length === 0, bad.length ? 'tidak valid: ' + bad.join(', ') : '');
        const badA = S.admins.filter(function (a) { return !Notify.validEmail(a.email); }).map(function (a) { return a.name; });
        add('Email admin valid', badA.length === 0, badA.length ? 'tidak valid: ' + badA.join(', ') : '');
      } catch (e) { add('Master data', false, e.message); }
    }
    if (S) {
      add('EMAIL_ENABLED = YES (email aktif)', Settings.yes(S, 'EMAIL_ENABLED'), S.cfg.EMAIL_ENABLED);
      try {
        const ex = Masters.staff(S).list.filter(function (x) { return x.active; }).map(function (x) { return x.email; })
          .concat(S.admins.map(function (a) { return a.email; })).filter(function (e) { return /@example\.(com|org|net)$/i.test(e); });
        add('Tidak ada email contoh (@example.com)', ex.length === 0, ex.length ? 'ganti: ' + ex.join(', ') : '');
      } catch (e) { add('Cek email contoh', false, e.message); }
      try {
        const p = PropertiesService.getScriptProperties();
        add('Form HP tugas & progres dibuat', !!(p.getProperty(CONFIG.PROP.FORM_TASK) && p.getProperty(CONFIG.PROP.FORM_PROGRESS)), 'menu Mobile Forms > Setup / Buat Form');
        add('Trigger form HP terpasang', Triggers.has('onTaskFormSubmit') && Triggers.has('onProgressFormSubmit'), '');
      } catch (e) { add('Form HP', false, e.message); }
    }
    add('Email akun aktif terbaca', !!Util.activeEmail(), Util.activeEmail() || 'kosong (akun mungkin tidak mengizinkan)');
    try { add('Trigger harian terpasang', Triggers.has('dailyJob'), Triggers.has('dailyJob') ? '' : 'jalankan menu Install Triggers'); } catch (e) { add('Trigger harian', false, 'izin belum diberikan: ' + e.message); }
    try { add('Kuota email harian', MailApp.getRemainingDailyQuota() > 0, String(MailApp.getRemainingDailyQuota())); } catch (e) { add('Kuota email', false, e.message); }
    try {
      const pend = Patch.pending();
      add('Patch terbaru (' + Patch.VERSION + ')', pend.length === 0, pend.length ? 'belum: ' + pend.map(function (x) { return x.id; }).join(', ') + ' -> menu Apply Patch' : '');
      if (S && !pend.length) {
        add('Web app aktif (WEB_ENABLED)', Settings.yes(S, 'WEB_ENABLED'), S.cfg.WEB_ENABLED);
        const st = Masters.staff(S).list.filter(function (x) { return x.active; });
        const noCode = st.filter(function (x) { return !x.code; }).map(function (x) { return x.name; })
          .concat(S.admins.filter(function (a) { return !a.code; }).map(function (a) { return a.name; }));
        add('Semua staf aktif & admin punya kode akses', noCode.length === 0, noCode.length ? 'belum: ' + noCode.join(', ') + ' -> menu Web App > Generate Access Codes' : '');
        const seen = {}; let dupC = 0;
        st.map(function (x) { return x.code; }).concat(S.admins.map(function (a) { return a.code; })).forEach(function (c) { const k = Auth.norm(c); if (k) { if (seen[k]) dupC++; seen[k] = 1; } });
        add('Kode akses unik', dupC === 0, dupC ? 'kode ganda: ' + dupC : '');
        add('Zona waktu web app = spreadsheet', Util.tz() === Session.getScriptTimeZone(), '');
      }
    } catch (e) { add('Pemeriksaan patch/web', false, e.message); }
    return out;
  },
  text(r) {
    const bad = r.filter(function (x) { return !x.ok; }).length;
    return (bad ? 'ADA ' + bad + ' MASALAH' : 'SEMUA PEMERIKSAAN LULUS') + '\n\n' +
      r.map(function (x) { return (x.ok ? '[OK]   ' : '[GAGAL] ') + x.name + (x.detail ? ' - ' + x.detail : ''); }).join('\n');
  }
};

// ============================================================================
//  Patch  (migrasi in-place: Google Sheet yang sudah ada TIDAK diganti, link tetap sama)
//  Idempotent: aman dijalankan berkali-kali. Selalu membuat salinan cadangan lebih dulu.
// ============================================================================
const Patch = {
  VERSION: '3.1.0',

  cfgRow_(sh, key) {
    const L = CONFIG.SETTINGS_LAYOUT;
    const keys = sh.getRange(L.CFG_FIRST, L.CFG_KEY_COL, L.CFG_LAST - L.CFG_FIRST + 1, 1).getDisplayValues();
    for (let i = 0; i < keys.length; i++) if (Util.str(keys[i][0]) === key) return L.CFG_FIRST + i;
    return 0;
  },

  /** Kolom "Deadline Time" (layout lama) di TASK INPUT, 0 bila tidak ada. */
  deadlineTimeCol_() {
    const sh = Util.ss().getSheetByName(CONFIG.SHEETS.INPUT);
    if (!sh) return 0;
    const lc = sh.getLastColumn();
    const v = sh.getRange(CONFIG.INPUT.HEADER_ROW, 1, 1, lc).getDisplayValues()[0];
    for (let i = 0; i < v.length; i++) if (/^deadline\s*time/i.test(Util.str(v[i]))) return i + 1;
    return 0;
  },

  webKeys_() {
    return [
      ['WEB_ENABLED', 'YES', 'YES = web app (HP) boleh dipakai. NO = tutup akses web.', true],
      ['WEB_APP_URL', '', 'Opsional: isi link web app (hasil Deploy) untuk dicatat/dibagikan.', false],
      ['EVIDENCE_PUBLIC_LINK', 'YES', 'YES = foto bukti bisa dibuka lewat link (tanpa login Google). NO = hanya email admin yang diberi akses.', true],
      ['REQUIRE_EVIDENCE_ON_COMPLETE', 'NO', 'YES = staf wajib melampirkan foto bukti saat menyelesaikan tugas (COMPLETED).', true],
      ['MAX_PHOTOS_PER_TASK', 20, 'Batas jumlah foto bukti per tugas.', false],
      ['REQUIRE_BEFORE_AFTER', 'NO', 'YES = tugas baru boleh COMPLETED (lewat web) hanya bila ada minimal 1 foto SEBELUM dan 1 foto SESUDAH.', true]
    ];
  },

  inputMsgFormula_(r, allowPastRef) {
    const DB = "'" + CONFIG.SHEETS.DB + "'!$A$2:$A$" + CONFIG.DB_MAX_ROWS;
    const parts = [
      'IF($D' + r + '="","Created By","")', 'IF($E' + r + '="","Project","")', 'IF($F' + r + '="","PIC","")',
      'IF($H' + r + '="","Task Title","")', 'IF($J' + r + '="","Priority","")',
      'IF($L' + r + '="","Deadline Date",IF(NOT(ISNUMBER($L' + r + ')),"Format Deadline Date",IF(AND(' + allowPastRef + '<>"YES",$L' + r + '<TODAY()),"Deadline sudah lewat","")))',
      'IF(AND($K' + r + '<>"",NOT(ISNUMBER($K' + r + '))),"Format Start Date","")',
      'IF(AND(ISNUMBER($K' + r + '),ISNUMBER($L' + r + ')),IF($L' + r + '<$K' + r + ',"Deadline < Start Date",""),"")',
      'IF(AND($G' + r + '<>"",COUNTIF(' + DB + ',$G' + r + ')=0),"Parent Task ID tidak ditemukan","")'
    ];
    const tj = 'TEXTJOIN(", ",TRUE,' + parts.join(',') + ')';
    return '=IF(COUNTA($C' + r + ':$M' + r + ')=0,"",IF($N' + r + '="SUBMITTED","TERKIRIM",IF(LEN(' + tj + ')=0,"SIAP SUBMIT","PERBAIKI: "&' + tj + ')))';
  },

  steps_() {
    const ss = Util.ss();
    return [
      {
        id: 'input-deadline-time', title: 'TASK INPUT: hapus kolom Deadline Time (deadline = akhir hari)',
        needed: function () { return Patch.deadlineTimeCol_() > 0; },
        apply: function () {
          const sh = Util.sheet(CONFIG.SHEETS.INPUT); const IN = CONFIG.INPUT;
          const col = Patch.deadlineTimeCol_();
          const hdr = function (c) { return Util.str(sh.getRange(IN.HEADER_ROW, c).getDisplayValue()); };
          if (col !== 13 || !/^deadline date/i.test(hdr(12)) || !/^notes/i.test(hdr(14))) {
            throw new Error('Layout TASK INPUT tidak dikenali (Deadline Time di kolom ' + col + '). Tidak diubah. Hubungi pembuat sistem.');
          }
          sh.deleteColumn(col);
          if (!/^notes/i.test(hdr(13)) || !/^submission/i.test(hdr(14))) throw new Error('Pemeriksaan sesudah hapus kolom gagal. Pulihkan dari salinan cadangan.');
          const sset = Util.sheet(CONFIG.SHEETS.SETTINGS);
          const r = Patch.cfgRow_(sset, 'ALLOW_PAST_DEADLINE');
          const ref = r ? 'SETTINGS!$N$' + r : '"NO"';
          const n = IN.LAST_ROW - IN.FIRST_ROW + 1;
          const fa = [], fo = [];
          for (let i = 0; i < n; i++) {
            const row = IN.FIRST_ROW + i;
            fa.push(['=IF(COUNTA($C' + row + ':$M' + row + ')=0,"",ROW()-' + (IN.FIRST_ROW - 1) + ')']);
            fo.push([Patch.inputMsgFormula_(row, ref)]);
          }
          sh.getRange(IN.FIRST_ROW, 1, n, 1).setFormulas(fa);
          sh.getRange(IN.FIRST_ROW, 15, n, 1).setFormulas(fo);
          sh.getRange('H2').setValue('Sheet ini untuk input massal di laptop (menu Task System > Submit Tasks). Dari HP: pakai web app / Google Form.');
          return 'kolom dihapus, rumus No & Validation Message dibangun ulang (' + n + ' baris)';
        }
      },
      {
        id: 'system-log', title: 'Buat sheet SYSTEM LOG',
        needed: function () { return !ss.getSheetByName(CONFIG.SHEETS.SYSTEM_LOG); },
        apply: function () {
          const sh = ss.insertSheet(CONFIG.SHEETS.SYSTEM_LOG);
          sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Source', 'Level', 'Message']])
            .setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
          [[1, 120], [2, 110], [3, 60], [4, 700]].forEach(function (w) { sh.setColumnWidth(w[0], w[1]); });
          sh.setFrozenRows(1);
          sh.getRange('F1').setValue('Dicatat otomatis oleh script (form HP, web app, email, error). Hanya-tambah; boleh dihapus manual bila terlalu panjang.').setFontStyle('italic').setFontColor('#595959');
          try { sh.setTabColor('#A6A6A6'); } catch (e) { /* noop */ }
          return 'sheet dibuat';
        }
      },
      {
        id: 'settings-web', title: 'SETTINGS: kolom Admin Access Code + konfigurasi web app',
        needed: function () {
          const sh = ss.getSheetByName(CONFIG.SHEETS.SETTINGS); if (!sh) return false;
          if (Util.str(sh.getRange(4, CONFIG.SETTINGS_LAYOUT.COL.ADMIN_CODE).getDisplayValue()) === '') return true;
          return Patch.webKeys_().some(function (k) { return !Patch.cfgRow_(sh, k[0]); });
        },
        apply: function () {
          const sh = Util.sheet(CONFIG.SHEETS.SETTINGS); const L = CONFIG.SETTINGS_LAYOUT; const notes = [];
          // kolom kode akses admin (L)
          sh.getRange(4, 11).copyTo(sh.getRange(4, L.COL.ADMIN_CODE), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          sh.getRange(5, 11, 10, 1).copyTo(sh.getRange(5, L.COL.ADMIN_CODE, 10, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          sh.getRange(4, L.COL.ADMIN_CODE).setValue('Admin Access Code');
          sh.setColumnWidth(L.COL.ADMIN_CODE, 130);
          // blok konfigurasi web (baris 31+, di bawah blok bobot)
          let blk = sh.getRange(L.WEB_BLOCK_ROW, L.CFG_KEY_COL);
          if (Util.str(blk.getDisplayValue()) === '') blk.setValue('WEB APP & FOTO BUKTI').setFontWeight('bold');
          Patch.webKeys_().forEach(function (k) {
            if (Patch.cfgRow_(sh, k[0])) return;
            let r = L.WEB_BLOCK_ROW + 1;
            while (Util.str(sh.getRange(r, L.CFG_KEY_COL).getDisplayValue()) !== '') r++;
            sh.getRange(r, L.CFG_KEY_COL, 1, 3).setValues([[k[0], k[1], k[2]]]);
            sh.getRange(r, L.CFG_KEY_COL).setFontWeight('bold');
            sh.getRange(r, L.CFG_VAL_COL).setFontColor('#0000FF').setHorizontalAlignment('center');
            sh.getRange(r, L.CFG_VAL_COL + 1).setFontColor('#595959').setFontSize(9);
            if (k[3]) sh.getRange(r, L.CFG_VAL_COL).setDataValidation(SpreadsheetApp.newDataValidation()
              .requireValueInRange(sh.getRange('K5:K6'), true).setAllowInvalid(false).build());
            notes.push(k[0]);
          });
          // sisa layout lama: DEFAULT_DEADLINE_TIME tidak dipakai lagi; email default aktif
          const rd = Patch.cfgRow_(sh, 'DEFAULT_DEADLINE_TIME');
          if (rd) {
            sh.getRange(rd, L.CFG_KEY_COL + 2).setValue('(tidak dipakai lagi - deadline = akhir hari 23:59)');
            const re = Patch.cfgRow_(sh, 'EMAIL_ENABLED');
            if (re && Util.up(sh.getRange(re, L.CFG_VAL_COL).getDisplayValue()) === 'NO') {
              sh.getRange(re, L.CFG_VAL_COL).setValue('YES'); notes.push('EMAIL_ENABLED NO -> YES (layout lama)');
            }
          }
          return 'kunci ditambah: ' + (notes.join(', ') || '-');
        }
      },
      {
        id: 'staff-code', title: 'STAFF MASTER: kolom Access Code',
        needed: function () {
          const sh = ss.getSheetByName(CONFIG.SHEETS.STAFF);
          return !!sh && Util.str(sh.getRange(1, 6).getDisplayValue()) !== 'Access Code';
        },
        apply: function () {
          const sh = Util.sheet(CONFIG.SHEETS.STAFF);
          if (Util.str(sh.getRange(1, 6).getDisplayValue()) !== '') throw new Error('STAFF MASTER F1 sudah berisi "' + sh.getRange(1, 6).getDisplayValue() + '". Kosongkan dulu kolom F.');
          sh.getRange(1, 5).copyTo(sh.getRange(1, 6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          sh.getRange(2, 5, 100, 1).copyTo(sh.getRange(2, 6, 100, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          sh.getRange(1, 6).setValue('Access Code');
          sh.setColumnWidth(6, 110);
          return 'kolom dibuat';
        }
      }
    ];
  },

  pending() { return Patch.steps_().filter(function (s) { return s.needed(); }); },

  /** Untuk onOpen: apakah ada patch yang belum diterapkan? (murah, tidak melempar error) */
  isCurrent() {
    try { return PropertiesService.getScriptProperties().getProperty(CONFIG.PROP.PATCH_VERSION) === Patch.VERSION; } catch (e) { return true; }
  },

  backup_() {
    const ss = Util.ss();
    const stamp = Utilities.formatDate(new Date(), Util.tz(), 'yyyy-MM-dd HH.mm');
    const copy = DriveApp.getFileById(ss.getId()).makeCopy('BACKUP sebelum patch ' + stamp + ' - ' + ss.getName());
    return copy.getUrl();
  },

  run() {
    return Util.withLock(function () {
      const pend = Patch.pending();
      const out = { version: Patch.VERSION, backup: '', applied: [], failed: null, already: pend.length === 0 };
      if (!pend.length) { PropertiesService.getScriptProperties().setProperty(CONFIG.PROP.PATCH_VERSION, Patch.VERSION); return out; }
      out.backup = Patch.backup_();   // gagal -> exception, tidak ada yang diubah
      const lines = [['INFO', 'Mulai patch ' + Patch.VERSION + '. Cadangan: ' + out.backup]];   // dicatat di akhir (SYSTEM LOG mungkin baru dibuat patch ini)
      for (let i = 0; i < pend.length; i++) {
        const st = pend[i];
        try {
          const d = st.apply();
          out.applied.push({ id: st.id, title: st.title, detail: d });
          lines.push(['INFO', st.title + ' -> ' + d]);
        } catch (e) {
          out.failed = { id: st.id, title: st.title, error: String(e && e.message ? e.message : e) };
          lines.push(['ERROR', st.title + ' GAGAL: ' + out.failed.error]);
          break;
        }
      }
      lines.forEach(function (l) { Log.audit(l[0], 'PATCH', l[1]); });
      if (!out.failed) PropertiesService.getScriptProperties().setProperty(CONFIG.PROP.PATCH_VERSION, Patch.VERSION);
      return out;
    });
  },

  text(r) {
    if (r.already) return 'Sudah versi terbaru (' + r.version + '). Tidak ada yang diubah.';
    const L = ['Patch ' + r.version, 'Cadangan: ' + r.backup, ''];
    r.applied.forEach(function (a) { L.push('[OK] ' + a.title + ' - ' + a.detail); });
    if (r.failed) L.push('[GAGAL] ' + r.failed.title + ' - ' + r.failed.error, '', 'Langkah berikutnya tidak dijalankan. Data aman; cadangan tersedia di atas.');
    else L.push('', 'Selesai. Lanjut: Task System > Web App > Generate Access Codes.');
    return L.join('\n');
  }
};

class UserError extends Error {}

// ============================================================================
//  Auth  (login web app dengan kode akses; sheet tetap sumber kebenaran)
// ============================================================================
const Auth = {
  ALPHABET: 'ABCDEFGHJKMNPQRSTUVWXYZ23456789',
  MAX_FAILS: 20, FAIL_WINDOW_S: 600,

  norm(c) { return Util.up(c).replace(/[^A-Z0-9]/g, ''); },

  newCode() {
    const hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    let s = '';
    for (let i = 0; i < 8; i++) s += Auth.ALPHABET.charAt(parseInt(hex.substr(i * 4, 4), 16) % Auth.ALPHABET.length);
    return s.substr(0, 4) + '-' + s.substr(4, 4);
  },

  failCount_() { try { return Number(CacheService.getScriptCache().get('AUTH_FAILS') || 0); } catch (e) { return 0; } },
  fail_() {
    try { CacheService.getScriptCache().put('AUTH_FAILS', String(Auth.failCount_() + 1), Auth.FAIL_WINDOW_S); } catch (e) { /* noop */ }
  },

  /** @return {role:'admin'|'staff', name, email, id} ; melempar UserError bila kode salah. */
  resolve(code, S) {
    const c = Auth.norm(code);
    const hits = [];
    if (c.length >= 6) {
      S.admins.forEach(function (a) { if (a.code && Auth.norm(a.code) === c) hits.push({ role: 'admin', name: a.name, email: a.email, id: '' }); });
      Masters.staff(S).list.forEach(function (x) {
        if (x.active && x.code && Auth.norm(x.code) === c) hits.push({ role: 'staff', name: x.name, email: x.email, id: x.id });
      });
    }
    if (hits.length === 1) return hits[0];     // kode benar selalu lolos: penyerang anonim tidak bisa mengunci pengguna sah
    Auth.fail_();
    if (Auth.failCount_() >= Auth.MAX_FAILS) Utilities.sleep(2500);   // tebakan salah beruntun diperlambat, bukan memblokir semua orang
    if (hits.length > 1) {
      Log.audit('WARN', 'WEB-LOGIN', 'Kode akses ganda dipakai oleh ' + hits.map(function (h) { return h.name; }).join(' & '));
      throw new UserError('Kode akses ganda. Hubungi admin.');
    }
    if (Auth.failCount_() <= Auth.MAX_FAILS) Log.audit('WARN', 'WEB-LOGIN', 'Kode akses salah/tidak aktif' + (Auth.failCount_() === Auth.MAX_FAILS ? ' (batas ' + Auth.MAX_FAILS + ' kegagalan/10 menit tercapai; kegagalan berikutnya tidak dicatat lagi)' : ''));
    throw new UserError('Kode akses salah atau akun tidak aktif.');
  },

  /** Isi kode akses yang masih kosong (tidak pernah menimpa kode yang ada). */
  generate() {
    return Util.withLock(function () {
      const S = Settings.load(); const L = CONFIG.SETTINGS_LAYOUT;
      const used = {}; let made = 0;
      const sset = Util.sheet(CONFIG.SHEETS.SETTINGS), sst = Util.sheet(CONFIG.SHEETS.STAFF);
      if (Util.str(sst.getRange(1, 6).getDisplayValue()) !== 'Access Code') throw new Error('Kolom Access Code belum ada. Jalankan Task System > Apply Patch dulu.');
      const stLast = Util.lastNonEmptyRow(sst, 2, 2);
      const sv = stLast >= 2 ? sst.getRange(2, 2, stLast - 1, 5).getDisplayValues() : [];   // B..F
      sv.forEach(function (r) { if (Util.str(r[4])) used[Auth.norm(r[4])] = 1; });
      const an = sset.getRange(L.LIST_FIRST, L.COL.ADMIN_NAME, L.LIST_LAST - L.LIST_FIRST + 1, 1).getDisplayValues();
      const ac = sset.getRange(L.LIST_FIRST, L.COL.ADMIN_CODE, L.LIST_LAST - L.LIST_FIRST + 1, 1).getDisplayValues();
      ac.forEach(function (r) { if (Util.str(r[0])) used[Auth.norm(r[0])] = 1; });
      const fresh = function () { let c; do { c = Auth.newCode(); } while (used[Auth.norm(c)]); used[Auth.norm(c)] = 1; return c; };
      const outS = sv.map(function (r) {
        if (Util.str(r[0]) && !Util.str(r[4])) { made++; return [fresh()]; }
        return [Util.str(r[4])];
      });
      if (outS.length) sst.getRange(2, 6, outS.length, 1).setNumberFormat('@').setValues(outS);
      const outA = an.map(function (r, i) {
        if (Util.str(r[0]) && !Util.str(ac[i][0])) { made++; return [fresh()]; }
        return [Util.str(ac[i][0])];
      });
      sset.getRange(L.LIST_FIRST, L.COL.ADMIN_CODE, outA.length, 1).setNumberFormat('@').setValues(outA);
      Log.audit('INFO', 'WEB-CODE', 'Kode akses baru dibuat: ' + made);
      return { made: made };
    });
  }
};

// ============================================================================
//  Analytics  (angka sama dengan sheet STAFF PERFORMANCE / DASHBOARD / PROJECT REPORT)
// ============================================================================
const Analytics = {
  derive(o, S, now) {
    const H = CONFIG.DB_H, ST = CONFIG.STATUS;
    const cs = Util.up(o[H.STATUS]); const dl = o[H.DEADLINE]; const cd = o[H.COMPLETION]; const sd = o[H.START];
    let eff = cs;
    if (cs !== ST.COMPLETED && cs !== ST.CANCELLED && Util.isDate(dl) && dl.getTime() < now.getTime()) eff = 'OVERDUE';
    const dc = Util.str(o[H.DELAY_CATEGORY]);
    let excused = 0;
    if (dc) {
      const c = S.delayCategories.filter(function (x) { return Util.low(x.name) === Util.low(dc); })[0];
      if (c && !c.countsAgainstStaff) excused = 1;
    }
    let tl = '';
    if (cs === ST.COMPLETED && Util.isDate(cd) && Util.isDate(dl)) tl = cd.getTime() <= dl.getTime() ? 'ON TIME' : (excused ? 'LATE (EXCUSED)' : 'LATE');
    const ql = Util.str(o[H.QUALITY]); let qs = null;
    if (ql) { const q = S.quality.filter(function (x) { return Util.low(x.label) === Util.low(ql); })[0]; if (q) qs = q.score; }
    let days = null;
    if (cs === ST.COMPLETED && Util.isDate(cd) && Util.isDate(sd)) days = Math.max(0, Math.round((cd.getTime() - sd.getTime()) / 8640000) / 10);
    return { eff: eff, tl: tl, excused: excused, qs: qs, days: days };
  },

  inPeriod(o, from, to) {
    const dl = o[CONFIG.DB_H.DEADLINE];
    if (!Util.isDate(dl)) return !from && !to;
    const d = Util.startOfDay(dl).getTime();
    return (!from || d >= from.getTime()) && (!to || d <= to.getTime());
  },

  parseDay(s) {
    return Util.ymd(Util.str(s).trim());
  },

  avg_(a) { return a.length ? a.reduce(function (x, y) { return x + y; }, 0) / a.length : null; },

  /** Metrik satu staf (rumus identik dengan STAFF PERFORMANCE). rows = [{o, d}] sudah difilter periode. */
  staffStats(rows, S, now, name) {
    const H = CONFIG.DB_H, ST = CONFIG.STATUS;
    const mine = rows.filter(function (r) { return Util.low(r.o[H.ASSIGNED_TO]) === Util.low(name); });
    const n = function (f) { return mine.filter(f).length; };
    const m = { name: name };
    m.total = n(function (r) { return r.d.eff !== ST.CANCELLED; });
    m.newCount = n(function (r) { return r.d.eff === ST.NEW; });
    m.inProgress = n(function (r) { return r.d.eff === ST.IN_PROGRESS || r.d.eff === 'IN PROGRESS'; });
    m.completed = n(function (r) { return r.d.eff === ST.COMPLETED; });
    m.pending = n(function (r) { return r.d.eff === ST.PENDING; });
    m.overdue = n(function (r) { return r.d.eff === 'OVERDUE'; });
    m.overdueStaff = n(function (r) { return r.d.eff === 'OVERDUE' && r.d.excused === 0; });
    m.onTime = n(function (r) { return r.d.tl === 'ON TIME'; });
    m.late = n(function (r) { return r.d.tl === 'LATE'; });
    m.lateExcused = n(function (r) { return r.d.tl === 'LATE (EXCUSED)'; });
    m.completionRate = m.total === 0 ? null : m.completed / m.total;
    m.onTimeRate = m.completed === 0 ? null : m.onTime / m.completed;
    m.onTimeAdj = m.completed === 0 ? null : (m.onTime + m.lateExcused) / m.completed;
    const done = mine.filter(function (r) { return r.d.eff === ST.COMPLETED; });
    const dd = done.filter(function (r) { return r.d.days !== null; }).map(function (r) { return r.d.days; });
    const qq = done.filter(function (r) { return r.d.qs !== null; }).map(function (r) { return r.d.qs; });
    m.avgDays = Analytics.avg_(dd); m.avgQuality = Analytics.avg_(qq);
    const stale = Util.startOfDay(now).getTime() - Settings.num(S, 'UPDATE_STALE_DAYS') * 86400000;
    const fresh = n(function (r) {
      return r.d.eff && r.d.eff !== ST.COMPLETED && r.d.eff !== ST.CANCELLED && Util.isDate(r.o[H.LAST_UPDATED]) && r.o[H.LAST_UPDATED].getTime() >= stale;
    });
    m.consistency = m.total === 0 ? null : (m.completed + fresh) / m.total;
    // skor
    m.score = null;
    if (m.total > 0) {
      const w = S.weights; const maxQ = Math.max.apply(null, S.quality.map(function (x) { return x.score; }).concat([0]));
      const hasO = m.onTimeAdj !== null, hasQ = m.avgQuality !== null && maxQ > 0;
      const num = w[0] * m.completionRate + (hasO ? w[1] * m.onTimeAdj : 0) + (hasQ ? w[2] * m.avgQuality / maxQ : 0) + w[3] * m.consistency;
      const den = w[0] + (hasO ? w[1] : 0) + (hasQ ? w[2] : 0) + w[3];
      if (den > 0) m.score = Math.round(1000 * num / den) / 10;
    }
    return m;
  },

  /** Semua data analitik untuk admin / satu staf. opts {from,to,staff} */
  build(S, now, opts) {
    const H = CONFIG.DB_H, ST = CONFIG.STATUS;
    const db = Db.readAll(Db.open());
    const from = Analytics.parseDay(opts.from), to = Analytics.parseDay(opts.to);
    const all = db.rows.map(function (o) { return { o: o, d: Analytics.derive(o, S, now) }; });
    const rows = all.filter(function (r) { return Analytics.inPeriod(r.o, from, to); });
    const staff = Masters.staff(S).list;
    const names = opts.staff ? [opts.staff] : staff.map(function (x) { return x.name; });
    const out = { staff: names.map(function (nm) { return Analytics.staffStats(rows, S, now, nm); }) };
    if (opts.staff) { out.all = all; out.rows = rows; return out; }

    const cnt = function (st) { return rows.filter(function (r) { return r.d.eff === st; }).length; };
    const cancelled = cnt(ST.CANCELLED);
    const total = rows.length; const completed = cnt(ST.COMPLETED);
    const onTime = rows.filter(function (r) { return r.d.tl === 'ON TIME'; }).length;
    out.kpi = { total: total, cancelled: cancelled, newCount: cnt(ST.NEW), inProgress: cnt('IN PROGRESS'), pending: cnt(ST.PENDING),
                completed: completed, overdue: cnt('OVERDUE'),
                completionRate: (total - cancelled) === 0 ? 0 : completed / (total - cancelled),
                onTimeRate: completed === 0 ? 0 : onTime / completed };
    // proyek (seluruh periode, seperti PROJECT REPORT)
    out.projects = Masters.projects(S).list.map(function (p) {
      const pr = all.filter(function (r) { return Util.low(r.o[H.PROJECT_NAME]) === Util.low(p.name); });
      const c = function (st) { return pr.filter(function (r) { return r.d.eff === st; }).length; };
      const tot = pr.length - c(ST.CANCELLED);
      const pics = []; pr.forEach(function (r) { const a = Util.str(r.o[H.ASSIGNED_TO]); if (a && pics.indexOf(a) < 0) pics.push(a); });
      return { id: p.id, name: p.name, location: p.location, manager: p.manager, active: p.active, total: tot, completed: c(ST.COMPLETED),
               inProgress: c('IN PROGRESS'), pending: c(ST.PENDING), newCount: c(ST.NEW), overdue: c('OVERDUE'), open: tot - c(ST.COMPLETED),
               rate: tot === 0 ? null : c(ST.COMPLETED) / tot, pics: pics };
    });
    out.all = all;
    return out;
  }
};

// ============================================================================
//  Photos  (foto bukti dari web app -> Google Drive + EVIDENCE LOG)
// ============================================================================
const Web = {
  /** URL dasar web app: SETTINGS > WEB_APP_URL, bila kosong URL deployment aktif. '' bila tidak ada. */
  url(S) {
    let u = Util.str(S && S.cfg && S.cfg.WEB_APP_URL).trim();
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(u)) u = '';
    if (!u) { try { u = ScriptApp.getService().getUrl() || ''; } catch (e) { u = ''; } }
    return /^https:\/\//.test(u) ? u : '';
  },
  /** Link dalam web app, mis. link(S, {tab:'report', staff:'Staf 01'}). '' bila web app belum di-deploy / WEB_ENABLED = NO. */
  link(S, params) {
    if (!Settings.yes(S, 'WEB_ENABLED')) return '';
    const base = Web.url(S); if (!base) return '';
    const q = Object.keys(params || {}).filter(function (k) { return params[k] !== '' && params[k] !== undefined && params[k] !== null; })
      .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
    return base + (q ? (base.indexOf('?') >= 0 ? '&' : '?') + q : '');
  },
  a(S, label, params) {
    const u = Web.link(S, params);
    return u ? '<a href="' + Util.esc(u) + '" style="color:#1a56db;font-weight:bold">' + Util.esc(label) + '</a>' : '';
  },
  /** Sanitasi parameter deep-link dari URL (hanya navigasi; login tetap wajib). */
  deep(e) {
    const p = (e && e.parameter) || {}; const out = {};
    if (['monitor', 'tasks', 'input', 'report'].indexOf(String(p.tab)) >= 0) out.tab = String(p.tab);
    if (/^[A-Za-z0-9_-]{3,30}$/.test(String(p.task || ''))) out.task = String(p.task).toUpperCase();
    if (p.staff && String(p.staff).length <= 80) out.staff = String(p.staff);
    return out;
  }
};

const Photos = {
  KINDS: { BEFORE: 'SEBELUM', AFTER: 'SESUDAH', OTHER: 'LAINNYA' },
  tagOf_(desc) { const m = /^\[(BEFORE|AFTER)\]\s?/.exec(String(desc || '')); return m ? { kind: m[1], desc: String(desc).substring(m[0].length) } : { kind: 'OTHER', desc: String(desc || '') }; },
  MIME: { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' },
  MAX_BYTES: 6 * 1024 * 1024,

  props_() { return PropertiesService.getScriptProperties(); },

  /** Folder induk (dibuat otomatis, id disimpan). Wajib dipanggil di dalam lock. */
  root_(S) {
    const p = Photos.props_(); let folder = null;
    const id = p.getProperty(CONFIG.PROP.EVIDENCE_FOLDER);
    if (id) { try { folder = DriveApp.getFolderById(id); if (folder.isTrashed()) folder = null; } catch (e) { folder = null; } }
    if (!folder) {
      folder = DriveApp.createFolder('Evidence - ' + S.cfg.SYSTEM_NAME + ' (' + Util.ss().getName() + ')');
      p.setProperty(CONFIG.PROP.EVIDENCE_FOLDER, folder.getId());
      Photos.share_(folder, S);
    }
    return folder;
  },

  share_(item, S) {
    const pub = Settings.yes(S, 'EVIDENCE_PUBLIC_LINK');
    try {
      if (pub) item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { Log.audit('WARN', 'WEB-PHOTO', 'Sharing link publik ditolak kebijakan Drive: ' + e.message); }
    S.admins.forEach(function (a) { if (Notify.validEmail(a.email)) { try { item.addViewer(a.email); } catch (e) { /* noop */ } } });
  },

  taskFolder_(root, taskId) {
    const it = root.getFoldersByName(taskId);
    return it.hasNext() ? it.next() : root.createFolder(taskId);
  },

  fileId_(url) {
    const m = /\/d\/([A-Za-z0-9_-]+)/.exec(url) || /[?&]id=([A-Za-z0-9_-]+)/.exec(url);
    return m ? m[1] : '';
  },

  sniff_(bytes) {
    const b = function (i) { return (bytes[i] || 0) & 255; };
    if (b(0) === 0xFF && b(1) === 0xD8 && b(2) === 0xFF) return 'image/jpeg';
    if (b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4E && b(3) === 0x47) return 'image/png';
    if (b(0) === 0x52 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x46 && b(8) === 0x57 && b(9) === 0x45 && b(10) === 0x42 && b(11) === 0x50) return 'image/webp';
    return '';
  },

  evidenceRows_() {
    const E_ = CONFIG.EV_H; const sh = Util.sheet(CONFIG.SHEETS.EVIDENCE); const h = Util.headerMap(sh, 1);
    const last = Math.max(Util.lastNonEmptyRow(sh, h[E_.ID], 2), Util.lastNonEmptyRow(sh, h[E_.URL], 2));
    if (last < 2) return [];
    const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const out = [];
    v.forEach(function (r) {
      if (Util.str(r[h[E_.ID] - 1]) === '') return;
      out.push({ id: Util.str(r[h[E_.ID] - 1]), taskId: Util.up(r[h[E_.TASK_ID] - 1]), date: r[h[E_.DATE] - 1], by: Util.str(r[h[E_.BY] - 1]),
                 desc: Photos.tagOf_(Util.str(r[h[E_.DESC] - 1])).desc, kind: Photos.tagOf_(Util.str(r[h[E_.DESC] - 1])).kind, url: Util.str(r[h[E_.URL] - 1]) });
    });
    return out;
  },

  /** Simpan satu foto. p = {taskId, name, mime, data(base64), desc}. t = baris tugas dari Db. */
  save(user, S, t, p) {
    const H = CONFIG.DB_H;
    const mime = Util.low(p.mime);
    if (!Photos.MIME[mime]) throw new UserError('Format foto tidak didukung (gunakan JPG/PNG/WebP).');
    const b64 = Util.str(p.data);
    if (!b64 || b64.length > Photos.MAX_BYTES * 1.4) throw new UserError('Ukuran foto terlalu besar atau kosong.');
    const bytes = Utilities.base64Decode(b64);
    if (bytes.length > Photos.MAX_BYTES) throw new UserError('Ukuran foto terlalu besar (maks 6 MB).');
    if (Photos.sniff_(bytes) !== mime) throw new UserError('Isi file bukan foto yang valid.');
    const taskId = t[H.TASK_ID];
    const kind = Photos.KINDS[Util.up(p.kind)] ? Util.up(p.kind) : 'OTHER';
    const maxN = Math.max(1, Settings.num(S, 'MAX_PHOTOS_PER_TASK'));
    return Util.withLock(function () {
      const have = Photos.evidenceRows_().filter(function (e) { return e.taskId === Util.up(taskId); }).length;
      if (have >= maxN) throw new UserError('Batas foto per tugas tercapai (' + maxN + ').');
      const now = new Date();
      const root = Photos.root_(S);
      const folder = Photos.taskFolder_(root, taskId);
      const stamp = Utilities.formatDate(now, Util.tz(), 'yyyyMMdd-HHmmss');
      const fname = taskId + '_' + (kind === 'OTHER' ? '' : Photos.KINDS[kind] + '_') + stamp + '_' + (have + 1) + '.' + Photos.MIME[mime];
      const file = folder.createFile(Utilities.newBlob(bytes, mime, fname));
      Photos.share_(file, S);
      const url = file.getUrl();
      // EVIDENCE LOG: tambah baris lalu proses dengan jalur yang sama dengan input manual
      const E_ = CONFIG.EV_H; const sh = Util.sheet(CONFIG.SHEETS.EVIDENCE); const h = Util.headerMap(sh, 1);
      const row = Math.max(Util.lastNonEmptyRow(sh, h[E_.ID], 2), Util.lastNonEmptyRow(sh, h[E_.TASK_ID], 2), Util.lastNonEmptyRow(sh, h[E_.URL], 2)) + 1;
      sh.getRange(row, h[E_.TASK_ID]).setValue(taskId);
      sh.getRange(row, h[E_.BY]).setValue(user.name);
      sh.getRange(row, h[E_.DESC]).setNumberFormat('@').setValue((kind === 'OTHER' ? '' : '[' + kind + '] ') + (Util.str(p.desc).substring(0, 300) || 'Foto ' + Photos.KINDS[kind].toLowerCase() + ' (web)'));
      sh.getRange(row, h[E_.URL]).setValue(url);
      const r = Evidence.process_(now);
      const mine = r.errors.filter(function (e) { return e.row === row; })[0];
      if (mine) { sh.getRange(row, 1, 1, sh.getLastColumn()).clearContent(); try { file.setTrashed(true); } catch (e) { /* noop */ } throw new UserError(mine.messages.join('; ')); }
      // Evidence Link di database = folder tugas (memuat semua foto)
      const db = Db.open();
      db.sheet.getRange(t._row, db.h[H.EVIDENCE]).setValue(folder.getUrl());
      Log.audit('INFO', 'WEB-PHOTO', taskId + ' foto ' + fname + ' oleh ' + user.name);
      return { evidenceId: r.processed[0] ? r.processed[0].id : '', url: url, folder: folder.getUrl(), kind: kind };
    });
  },

  image(url) {
    const id = Photos.fileId_(url);
    if (!id) throw new UserError('Link bukan file Drive.');
    const blob = DriveApp.getFileById(id).getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > 4 * 1024 * 1024) throw new UserError('Foto terlalu besar untuk ditampilkan; buka lewat link.');
    return { mime: blob.getContentType(), data: Utilities.base64Encode(bytes) };
  }
};

// ============================================================================
//  Api  (dipanggil browser lewat google.script.run -> webApi)
// ============================================================================
const Api = {
  fmtDt_(d) { return Util.isDate(d) ? Utilities.formatDate(d, Util.tz(), 'yyyy-MM-dd HH:mm') : ''; },
  fmtD_(d) { return Util.isDate(d) ? Utilities.formatDate(d, Util.tz(), 'yyyy-MM-dd') : ''; },

  taskDto_(o, d, now, evCount) {
    const H = CONFIG.DB_H; const dl = o[H.DEADLINE];
    return {
      id: o[H.TASK_ID], title: Util.str(o[H.TITLE]), desc: Util.str(o[H.DESC]), project: Util.str(o[H.PROJECT_NAME]), pic: Util.str(o[H.ASSIGNED_TO]),
      priority: Util.str(o[H.PRIORITY]), status: Util.up(o[H.STATUS]), eff: d.eff, progress: Number(o[H.PROGRESS]) || 0,
      start: Api.fmtD_(o[H.START]), deadline: Api.fmtD_(dl), createdBy: Util.str(o[H.CREATED_BY]), createdAt: Api.fmtDt_(o[H.CREATED_AT]),
      lastUpdated: Api.fmtDt_(o[H.LAST_UPDATED]), completion: Api.fmtDt_(o[H.COMPLETION]), parent: Util.str(o[H.PARENT]),
      delayReason: Util.str(o[H.DELAY_REASON]), resultNotes: Util.str(o[H.RESULT_NOTES]), delayCategory: Util.str(o[H.DELAY_CATEGORY]),
      quality: Util.str(o[H.QUALITY]), timeliness: d.tl, evidence: evCount || 0,
      daysLeft: Util.isDate(dl) ? Math.round((dl.getTime() - now.getTime()) / 8640000) / 10 : null
    };
  },

  task_(db, id) {
    const t = db.byId[Util.up(id)];
    if (!t) throw new UserError('Tugas "' + id + '" tidak ditemukan.');
    return t;
  },
  own_(user, t) {
    if (user.role === 'staff' && Util.low(t[CONFIG.DB_H.ASSIGNED_TO]) !== Util.low(user.name)) throw new UserError('Tugas ini bukan milik Anda.');
  },
  admin_(user) { if (user.role !== 'admin') throw new UserError('Hanya admin yang boleh melakukan ini.'); },

  meta_(S, user) {
    const out = {
      systemName: S.cfg.SYSTEM_NAME, user: { role: user.role, name: user.name }, priorities: S.priorities,
      statuses: S.statuses, delayCategories: S.delayCategories.map(function (x) { return x.name; }), quality: S.quality.map(function (x) { return x.label; }),
      requireDelay: Settings.yes(S, 'REQUIRE_DELAY_REASON'), requireEvidence: Settings.yes(S, 'REQUIRE_EVIDENCE_ON_COMPLETE'), requireBeforeAfter: Settings.yes(S, 'REQUIRE_BEFORE_AFTER'),
      dueSoonDays: Settings.num(S, 'DUE_SOON_DAYS'), today: Utilities.formatDate(new Date(), Util.tz(), 'yyyy-MM-dd')
    };
    if (user.role === 'admin') {
      out.projects = Masters.projects(S).list.filter(function (x) { return x.active; }).map(function (x) { return x.name; });
      out.staff = Masters.staff(S).list.filter(function (x) { return x.active; }).map(function (x) { return x.name; });
      out.allStaff = Masters.staff(S).list.map(function (x) { return x.name; });
    }
    return out;
  },

  // ---------------------------------------------------------------- actions
  actions: {
    login(user, p, S) { return Api.meta_(S, user); },

    tasks(user, p, S) {
      const now = new Date(); const H = CONFIG.DB_H, ST = CONFIG.STATUS;
      const db = Db.readAll(Db.open()); const ev = Photos.evidenceRows_(); const evc = {};
      ev.forEach(function (e) { evc[e.taskId] = (evc[e.taskId] || 0) + 1; });
      const scope = p.scope || 'open'; const q = Util.low(p.q);
      let list = db.rows.map(function (o) { return { o: o, d: Analytics.derive(o, S, now) }; });
      if (user.role === 'staff') list = list.filter(function (r) { return Util.low(r.o[H.ASSIGNED_TO]) === Util.low(user.name); });
      else {
        if (p.pic) list = list.filter(function (r) { return Util.low(r.o[H.ASSIGNED_TO]) === Util.low(p.pic); });
        if (p.project) list = list.filter(function (r) { return Util.low(r.o[H.PROJECT_NAME]) === Util.low(p.project); });
      }
      const closed = function (r) { return r.d.eff === ST.COMPLETED || r.d.eff === ST.CANCELLED; };
      if (scope === 'open') list = list.filter(function (r) { return !closed(r); });
      else if (scope === 'done') list = list.filter(closed);
      if (q) list = list.filter(function (r) { return (Util.low(r.o[H.TITLE]) + ' ' + Util.low(r.o[H.TASK_ID]) + ' ' + Util.low(r.o[H.PROJECT_NAME]) + ' ' + Util.low(r.o[H.ASSIGNED_TO])).indexOf(q) >= 0; });
      const dlv = function (r) { return Util.isDate(r.o[H.DEADLINE]) ? r.o[H.DEADLINE].getTime() : 9e15; };
      list.sort(scope === 'open' ? function (a, b) { return dlv(a) - dlv(b); } : function (a, b) { return dlv(b) - dlv(a); });
      const LIM = 200; const total = list.length;
      return { total: total, truncated: total > LIM,
               tasks: list.slice(0, LIM).map(function (r) { return Api.taskDto_(r.o, r.d, now, evc[Util.up(r.o[H.TASK_ID])]); }) };
    },

    task(user, p, S) {
      const now = new Date(); const H = CONFIG.DB_H;
      const db = Db.readAll(Db.open()); const t = Api.task_(db, p.taskId); Api.own_(user, t);
      const id = Util.up(t[H.TASK_ID]);
      const ev = Photos.evidenceRows_().filter(function (e) { return e.taskId === id; });
      // riwayat progres
      const P = CONFIG.PL_H; const sh = Util.sheet(CONFIG.SHEETS.PROGRESS); const h = Util.headerMap(sh, 1);
      const last = Util.lastNonEmptyRow(sh, h[P.LOG_ID], 2); const hist = [];
      if (last >= 2) {
        sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues().forEach(function (r) {
          if (Util.str(r[h[P.LOG_ID] - 1]) === '' || Util.up(r[h[P.TASK_ID] - 1]) !== id) return;
          hist.push({ at: Api.fmtDt_(r[h[P.TS] - 1]), by: Util.str(r[h[P.UPDATED_BY] - 1]), from: Util.str(r[h[P.PREV] - 1]), to: Util.str(r[h[P.NEW] - 1]),
                      progress: Number(r[h[P.PROGRESS] - 1]) || 0, notes: Util.str(r[h[P.NOTES] - 1]), delay: Util.str(r[h[P.DELAY] - 1]) });
        });
      }
      hist.reverse();
      return { task: Api.taskDto_(t, Analytics.derive(t, S, now), now, ev.length), history: hist,
               evidence: ev.map(function (e) { return { id: e.id, at: Api.fmtDt_(e.date), by: e.by, desc: e.desc, kind: e.kind, url: e.url }; }) };
    },

    evidenceImage(user, p, S) {
      const e = Photos.evidenceRows_().filter(function (x) { return x.id === Util.up(p.evidenceId); })[0];
      if (!e) throw new UserError('Foto tidak ditemukan.');
      const t = Api.task_(Db.readAll(Db.open()), e.taskId); Api.own_(user, t);
      return Photos.image(e.url);
    },

    uploadPhoto(user, p, S) {
      const db = Db.readAll(Db.open()); const t = Api.task_(db, p.taskId); Api.own_(user, t);
      return Photos.save(user, S, t, p);
    },

    submitUpdate(user, p, S) {
      Util.assertTimezone();
      const H = CONFIG.DB_H, ST = CONFIG.STATUS;
      const db = Db.readAll(Db.open()); const t = Api.task_(db, p.taskId); Api.own_(user, t);
      const ns = Util.up(p.status);
      if (!ns) throw new UserError('Status baru wajib dipilih.');
      if (ns === ST.CANCELLED && user.role !== 'admin') throw new UserError('Status CANCELLED hanya boleh oleh admin.');
      let pn = '';
      if (Util.str(p.progress) !== '') {
        pn = Number(p.progress);
        if (isNaN(pn) || pn < 0 || pn > 100) throw new UserError('Progress harus 0-100.');
        pn = pn / 100;
      }
      if (ns === ST.COMPLETED && Settings.yes(S, 'REQUIRE_EVIDENCE_ON_COMPLETE') &&
          Photos.evidenceRows_().filter(function (e) { return e.taskId === Util.up(t[H.TASK_ID]); }).length === 0) {
        throw new UserError('Lampirkan minimal 1 foto bukti sebelum menyelesaikan tugas.');
      }
      if (Util.str(p.notes).length > CONFIG.LIMITS.TEXT || Util.str(p.delay).length > CONFIG.LIMITS.TEXT) throw new UserError('Catatan terlalu panjang (maks ' + CONFIG.LIMITS.TEXT + ' karakter).');
      if (ns === ST.COMPLETED && Settings.yes(S, 'REQUIRE_BEFORE_AFTER')) {
        const evs = Photos.evidenceRows_().filter(function (e) { return e.taskId === Util.up(t[H.TASK_ID]); });
        if (!evs.some(function (e) { return e.kind === 'BEFORE'; }) || !evs.some(function (e) { return e.kind === 'AFTER'; })) {
          throw new UserError('Lampirkan foto SEBELUM dan foto SESUDAH sebelum menyelesaikan tugas.');
        }
      }
      const revised = Util.str(p.revised) ? Analytics.parseDay(p.revised) : '';
      if (Util.str(p.revised) && !revised) throw new UserError('Format tanggal revisi tidak valid.');
      const ar = Progress.appendAndRun_({ taskId: t[H.TASK_ID], by: user.name, ns: ns, pn: pn, notes: p.notes, delay: p.delay, revised: revised });
      if (ar.mine) throw new UserError(ar.mine.messages.join('; '));
      Log.audit('INFO', 'WEB-UPDATE', t[H.TASK_ID] + ' -> ' + ns + ' oleh ' + user.name);
      const t2 = Db.readAll(Db.open()).byId[Util.up(p.taskId)]; const now = new Date();
      return { task: Api.taskDto_(t2, Analytics.derive(t2, S, now), now, 0) };
    },

    createTasks(user, p, S) {
      Api.admin_(user);
      if (!Array.isArray(p.items) || !p.items.length) throw new UserError('Tidak ada tugas untuk dikirim.');
      if (p.items.length > 30) throw new UserError('Maksimal 30 tugas per pengiriman.');
      const res = Submit.fromWeb(user, p.items);
      return res;
    },

    rateTask(user, p, S) {
      Api.admin_(user);
      const H = CONFIG.DB_H;
      return Util.withLock(function () {
        const db = Db.open(); const t = Api.task_(Db.readAll(db), p.taskId);
        const dc = Util.str(p.delayCategory), q = Util.str(p.quality);
        if (dc && !S.delayCategories.some(function (x) { return x.name === dc; })) throw new UserError('Delay Category tidak valid.');
        if (q && !S.quality.some(function (x) { return x.label === q; })) throw new UserError('Quality Rating tidak valid.');
        if (q && Util.up(t[H.STATUS]) !== CONFIG.STATUS.COMPLETED) throw new UserError('Quality Rating hanya untuk tugas COMPLETED.');
        const before = Util.str(t[H.DELAY_CATEGORY]) + ' / ' + Util.str(t[H.QUALITY]);
        db.sheet.getRange(t._row, db.h[H.DELAY_CATEGORY]).setValue(dc);
        db.sheet.getRange(t._row, db.h[H.QUALITY]).setValue(q);
        Log.audit('INFO', 'WEB-RATE', t[H.TASK_ID] + ' oleh ' + user.name + ': [' + before + '] -> [' + dc + ' / ' + q + ']');
        return { ok: true };
      });
    },

    dashboard(user, p, S) {
      Api.admin_(user);
      const now = new Date(); const H = CONFIG.DB_H;
      const a = Analytics.build(S, now, { from: p.from, to: p.to });
      const evc = {}; Photos.evidenceRows_().forEach(function (e) { evc[e.taskId] = (evc[e.taskId] || 0) + 1; });
      const days = Settings.num(S, 'DUE_SOON_DAYS'); const ST = CONFIG.STATUS;
      const mk = function (r) { return Api.taskDto_(r.o, r.d, now, evc[Util.up(r.o[H.TASK_ID])]); };
      const dlv = function (r) { return r.o[H.DEADLINE].getTime(); };
      const od = a.all.filter(function (r) { return r.d.eff === 'OVERDUE'; }).sort(function (x, y) { return dlv(x) - dlv(y); });
      const ds = a.all.filter(function (r) {
        return Util.isDate(r.o[H.DEADLINE]) && r.d.eff !== 'OVERDUE' && r.d.eff !== ST.COMPLETED && r.d.eff !== ST.CANCELLED &&
               r.o[H.DEADLINE].getTime() - now.getTime() <= days * 86400000;
      }).sort(function (x, y) { return dlv(x) - dlv(y); });
      return { kpi: a.kpi, staff: a.staff, projects: a.projects, overdue: od.slice(0, 50).map(mk), overdueTotal: od.length,
               dueSoon: ds.slice(0, 50).map(mk), dueSoonTotal: ds.length, at: Api.fmtDt_(now) };
    },

    report(user, p, S) {
      const now = new Date(); const H = CONFIG.DB_H;
      let name = user.name;
      if (user.role === 'admin') {
        name = Util.str(p.staff);
        if (!name || !Masters.staff(S).byName[Util.low(name)]) throw new UserError('Pilih staf.');
        name = Masters.staff(S).byName[Util.low(name)].name;
      }
      const a = Analytics.build(S, now, { from: p.from, to: p.to, staff: name });
      const mine = a.rows.filter(function (r) { return Util.low(r.o[H.ASSIGNED_TO]) === Util.low(name); });
      const fin = mine.filter(function (r) { return r.d.eff === 'COMPLETED'; })
        .sort(function (x, y) { return (Util.isDate(y.o[H.COMPLETION]) ? y.o[H.COMPLETION].getTime() : 0) - (Util.isDate(x.o[H.COMPLETION]) ? x.o[H.COMPLETION].getTime() : 0); });
      return { name: name, stats: a.staff[0], recent: fin.slice(0, 15).map(function (r) { return Api.taskDto_(r.o, r.d, now, 0); }),
               weights: S.weights };
    }
  },

  /** Penanganan umum: kunci idempotensi (reqId), otorisasi, JSON. */
  handle(code, action, payloadJson) {
    try {
      let p = {};
      try { p = payloadJson ? JSON.parse(payloadJson) : {}; } catch (e) { throw new UserError('Permintaan tidak valid.'); }
      if (p === null || typeof p !== 'object' || Array.isArray(p)) throw new UserError('Permintaan tidak valid.');
      if (!Api.actions.hasOwnProperty(action)) throw new UserError('Aksi tidak dikenal.');
      const S = Settings.load();
      if (!Settings.yes(S, 'WEB_ENABLED')) throw new UserError('Web app sedang dinonaktifkan oleh admin.');
      const user = Auth.resolve(code, S);
      const mutating = ['uploadPhoto', 'submitUpdate', 'createTasks', 'rateTask'].indexOf(action) >= 0;
      const key = mutating && p.reqId ? 'REQ_' + Auth.norm(code).substring(0, 4) + '_' + Util.str(p.reqId).replace(/[^A-Za-z0-9_-]/g, '').substring(0, 40) : '';
      const cache = CacheService.getScriptCache();
      if (key) { const hit = cache.get(key); if (hit) return hit; }
      const data = Api.actions[action](user, p, S);
      const out = JSON.stringify({ ok: true, data: data });
      if (key && out.length < 90000) cache.put(key, out, 600);
      return out;
    } catch (e) {
      if (!(e instanceof UserError)) Log.error('webApi:' + action, e);
      return JSON.stringify({ ok: false, error: e instanceof UserError ? e.message : 'Kesalahan sistem: ' + (e && e.message ? e.message : e) });
    }
  }
};

// ============================================================================
//  Fungsi publik (menu / tombol / trigger)
// ============================================================================

/** Menu otomatis saat spreadsheet dibuka. */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Task System')
      .addItem('Submit Tasks', 'submitTasks')
      .addItem('Process Progress Updates', 'processProgressUpdates')
      .addItem('Process Evidence Log', 'processEvidenceLog')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Web App (HP: admin & staf)')
        .addItem('Generate Access Codes', 'generateAccessCodes')
        .addItem('Tampilkan Link Web App', 'showWebAppUrl')
        .addItem('Cek Folder Foto Bukti', 'checkEvidenceFolder'))
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Mobile Forms (input via HP)')
        .addItem('Setup / Buat Form', 'setupMobileForms')
        .addItem('Refresh Pilihan Form', 'refreshFormChoices')
        .addItem('Tampilkan Link Form', 'showFormLinks'))
      .addSeparator()
      .addItem('Clear Submitted Rows (TASK INPUT)', 'clearSubmittedRows')
      .addItem('Detect Overdue (laporan)', 'detectOverdueReport')
      .addItem('Run Daily Reminder Now', 'runDailyReminderNow')
      .addSeparator()
      .addItem('Install Triggers (reminder harian)', 'installTriggers')
      .addItem('Install Form Trigger', 'installFormTrigger')
      .addItem('Remove Triggers', 'removeTriggers')
      .addItem('Apply Protections (peringatan)', 'applyProtections')
      .addSeparator()
      .addItem('Apply Patch (update sistem)', 'applyPatch')
      .addItem('Self Test (cek instalasi)', 'selfTest')
      .addItem('Send Test Email', 'sendTestEmail')
      .addToUi();
    if (!Patch.isCurrent()) Util.toast('Ada pembaruan sistem. Jalankan menu Task System > Apply Patch.', 'Update tersedia');
  } catch (e) {
    // Terjadi bila onOpen dijalankan dari editor Apps Script (tombol Run): tidak ada UI spreadsheet. Bukan error.
    Log.info('onOpen', 'Menu dilewati (tidak ada UI di konteks ini). Buka/refresh spreadsheet untuk memunculkan menu Task System.');
  }
}

/** TOMBOL SUBMIT TASK: kirim baris TASK INPUT ke TASK DATABASE. */
function submitTasks() {
  return Util.run('Submit Tasks', function () {
    const r = Submit.run();
    Util.alert('Submit Tasks', Submit.summaryText(r));
    return r;
  });
}

function processProgressUpdates() {
  return Util.run('Process Progress Updates', function () {
    const r = Progress.run(); Util.alert('Process Progress Updates', Progress.summaryText(r)); return r;
  });
}

function processEvidenceLog() {
  return Util.run('Process Evidence Log', function () {
    const r = Evidence.run(); Util.alert('Process Evidence Log', Evidence.summaryText(r)); return r;
  });
}

/** Hapus isi baris TASK INPUT yang sudah SUBMITTED (data tetap aman di TASK DATABASE). */
function clearSubmittedRows() {
  return Util.run('Clear Submitted Rows', function () {
    const ui = Util.ui(); const IN = CONFIG.INPUT, C = IN.COL;
    if (ui && ui.alert('Clear Submitted Rows', 'Hapus isi baris TASK INPUT yang berstatus SUBMITTED?\nData di TASK DATABASE TIDAK terpengaruh.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return null;
    return Util.withLock(function () {
      const sh = Util.sheet(CONFIG.SHEETS.INPUT);
      const n = IN.LAST_ROW - IN.FIRST_ROW + 1;
      const st = sh.getRange(IN.FIRST_ROW, C.STATUS, n, 1).getValues();
      const ids = sh.getRange(IN.FIRST_ROW, C.TASK_ID, n, 1).getValues();
      const db = Db.readAll(Db.open());
      let cleared = 0, skipped = 0, i = 0;
      while (i < n) {
        if (Util.up(st[i][0]) !== IN.SUBMITTED) { i++; continue; }
        let j = i;
        while (j + 1 < n && Util.up(st[j + 1][0]) === IN.SUBMITTED) j++;
        // keamanan: hanya hapus bila Task ID benar-benar ada di database
        let ok = true;
        for (let k = i; k <= j; k++) if (!db.byId[Util.up(ids[k][0])]) ok = false;
        if (ok) { sh.getRange(IN.FIRST_ROW + i, C.TASK_ID, j - i + 1, C.STATUS - C.TASK_ID + 1).clearContent(); cleared += j - i + 1; }
        else skipped += j - i + 1;
        i = j + 1;
      }
      Util.alert('Clear Submitted Rows', 'Baris dibersihkan: ' + cleared + (skipped ? '\nDilewati (Task ID tidak ditemukan di database): ' + skipped : ''));
      return cleared;
    });
  });
}

function detectOverdueReport() {
  return Util.run('Detect Overdue', function () {
    const det = Overdue.detect(new Date());
    const fmt = function (a) { return a.slice(0, 15).map(function (t) { return '  ' + t.id + ' | ' + t.pic + ' | ' + Util.fmt(t.deadline) + ' | ' + t.title; }).join('\n'); };
    Util.alert('Detect Overdue', 'Overdue: ' + det.overdue.length + '\n' + fmt(det.overdue) + '\n\nMendekati deadline: ' + det.dueSoon.length + '\n' + fmt(det.dueSoon));
    return det;
  });
}

/** Handler trigger harian (time-driven). Aman dipanggil berkali-kali per hari. */
function dailyJob() {
  try { Forms.refreshQuiet(); return Reminders.run(false); } catch (e) { Log.error('dailyJob', e); return null; }
}

function runDailyReminderNow() {
  return Util.run('Run Daily Reminder', function () {
    const r = Reminders.run(true);
    Util.alert('Daily Reminder', 'Overdue: ' + r.overdue + '\nMendekati deadline: ' + r.dueSoon + '\n' +
      (r.mail ? 'Email terkirim: ' + r.mail.sent + (r.mail.failed.length ? '\nGagal: ' + r.mail.failed.join('; ') : '') : r.reason));
    return r;
  });
}

function installTriggers() { return Util.run('Install Triggers', function () { const m = Triggers.install(); Util.alert('Install Triggers', m); return m; }); }
function installFormTrigger() { return Util.run('Install Form Trigger', function () { const m = Triggers.installForm(); Util.alert('Install Form Trigger', m); return m; }); }
function removeTriggers() { return Util.run('Remove Triggers', function () { const m = Triggers.remove(); Util.alert('Remove Triggers', m); return m; }); }
function applyProtections() { return Util.run('Apply Protections', function () { const m = Protections.apply(); Util.alert('Apply Protections', m); return m; }); }

function selfTest() {
  return Util.run('Self Test', function () { const r = SelfTest.run(); Util.alert('Self Test', SelfTest.text(r)); return r; });
}

function sendTestEmail() {
  return Util.run('Send Test Email', function () {
    const S = Settings.load(); const to = Util.activeEmail();
    if (!to) throw new Error('Email akun aktif tidak terbaca.');
    const r = Notify.send(S, to, '', 'Tes email', '<p>Email tes dari ' + Util.esc(S.cfg.SYSTEM_NAME) + ' berhasil.</p>');
    Util.alert('Send Test Email', r.ok ? 'Terkirim ke ' + to : 'GAGAL: ' + r.error);
    return r;
  });
}

/** Simple trigger: dipanggil otomatis oleh Google Sheets pada setiap edit manual. */
function onEdit(e) { try { Guard.onEdit(e); } catch (err) { Log.error('onEdit', err); } }

/** Buat Google Form HP (tugas & progres), isi dropdown, pasang trigger. */
function setupMobileForms() {
  return Util.run('Setup Mobile Forms', function () {
    const r = Forms.setup();
    Util.alert('Mobile Forms', (r.created.length ? 'Form dibuat: ' + r.created.join(', ') : 'Form sudah ada - pilihan dropdown diperbarui.') + '\n\n' + Forms.linksText_(r.links));
    return r;
  });
}
function refreshFormChoices() {
  return Util.run('Refresh Form Choices', function () { const r = Forms.refresh(); Util.alert('Refresh Pilihan Form', 'Form tugas: ' + (r.task ? 'OK' : 'belum dibuat') + '\nForm progres: ' + (r.progress ? 'OK' : 'belum dibuat')); return r; });
}
function showFormLinks() {
  return Util.run('Form Links', function () { const l = Forms.links(); Util.alert('Link Form HP', Forms.linksText_(l)); return l; });
}
Forms.linksText_ = function (l) {
  if (!l.task && !l.progress) return 'Form belum dibuat. Jalankan Mobile Forms > Setup / Buat Form.';
  return 'LINK UNTUK HP (bagikan ke yang berhak):\n\nForm TUGAS (admin):\n' + (l.task || '-') + '\n\nForm PROGRES (staf & admin):\n' + (l.progress || '-') +
    '\n\nTips: buka link di HP > menu browser > "Tambahkan ke layar utama".\nEdit form (hanya pemilik):\n' + (l.taskEdit || '-') + '\n' + (l.progressEdit || '-');
};

/** Handler trigger Google Form TUGAS (dipasang oleh Setup / Buat Form). */
function onTaskFormSubmit(e) {
  try {
    const email = e.response.getRespondentEmail();
    Submit.fromForm(Forms.answers_(e.response), email);
  } catch (err) {
    Log.error('onTaskFormSubmit', err);
    try { Notify.rejectToSender(e.response.getRespondentEmail(), 'Tugas TIDAK tercatat (error sistem)', [String(err.message || err)]); } catch (e2) { /* noop */ }
  }
}

/** Handler trigger Google Form PROGRES. */
function onProgressFormSubmit(e) {
  try {
    const email = e.response.getRespondentEmail();
    Progress.fromFormResponse(Forms.answers_(e.response), email);
  } catch (err) {
    Log.error('onProgressFormSubmit', err);
    try { Notify.rejectToSender(e.response.getRespondentEmail(), 'Update progres TIDAK tercatat (error sistem)', [String(err.message || err)]); } catch (e2) { /* noop */ }
  }
}

/** Trigger installable Google Form EVIDENCE (Install Form Trigger; form dibuat manual & ditautkan ke spreadsheet). */
function onFormSubmitRouter(e) {
  try {
    const nv = e && e.namedValues ? e.namedValues : {};
    if (nv[CONFIG.FORMS.EVIDENCE.FILES]) return Evidence.fromForm(nv);
    Log.info('onFormSubmitRouter', 'form tidak dikenali: ' + Object.keys(nv).join(', '));
  } catch (err) { Log.error('onFormSubmitRouter', err); }
  return null;
}

// ---------------------------------------------------------------- Web App & Patch (menu + endpoint)

/** Menu: terapkan patch ke Google Sheet yang sedang dipakai (link tetap sama, ada cadangan otomatis). */
function applyPatch() {
  return Util.run('Apply Patch', function () {
    const ui = Util.ui();
    if (ui) {
      const pend = Patch.pending();
      if (!pend.length) { const r0 = Patch.run(); Util.alert('Apply Patch', Patch.text(r0)); return r0; }
      const msg = 'Patch akan mengubah Sheet ini (link tetap sama). Salinan cadangan dibuat otomatis di Google Drive Anda.\n\nLangkah:\n- ' +
        pend.map(function (s) { return s.title; }).join('\n- ') + '\n\nLanjutkan?';
      if (ui.alert('Apply Patch', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return null;
    }
    const r = Patch.run(); Util.alert('Apply Patch', Patch.text(r)); return r;
  });
}

function generateAccessCodes() {
  return Util.run('Generate Access Codes', function () {
    const r = Auth.generate();
    Util.alert('Generate Access Codes', 'Kode baru dibuat: ' + r.made + '\n\nKode ada di STAFF MASTER kolom F (staf) dan SETTINGS kolom L (admin). Bagikan ke masing-masing orang secara pribadi. Kode yang sudah ada tidak diubah; untuk mengganti kode, kosongkan selnya lalu jalankan menu ini lagi.');
    return r;
  });
}

function showWebAppUrl() {
  return Util.run('Link Web App', function () {
    let url = ''; try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
    Util.alert('Link Web App', url ? 'Link web app (bagikan bersama kode akses masing-masing):\n\n' + url + '\n\nTips: buka di HP > menu browser > "Tambahkan ke layar utama".'
      : 'Web app belum di-deploy. Apps Script: Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone. Lihat README.');
    return url;
  });
}

function checkEvidenceFolder() {
  return Util.run('Folder Foto Bukti', function () {
    const S = Settings.load();
    const f = Util.withLock(function () { return Photos.root_(S); });
    Util.alert('Folder Foto Bukti', 'Folder: ' + f.getName() + '\n' + f.getUrl());
    return f.getUrl();
  });
}

/** Endpoint web app: halaman HTML (Index.html). */
function doGet(e) {
  let name = 'Staff Performance Tracking';
  try { name = Settings.load().cfg.SYSTEM_NAME || name; } catch (e) { /* tetap tampilkan halaman login */ }
  const t = HtmlService.createTemplateFromFile('Index');
  t.systemName = name;
  t.deep = JSON.stringify(Web.deep(e)).replace(/</g, '\\u003c');
  return t.evaluate().setTitle(name).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Dipanggil browser: google.script.run.webApi(kode, aksi, jsonPayload) -> string JSON {ok, data|error}. */
function webApi(code, action, payloadJson) { return Api.handle(code, action, payloadJson); }
