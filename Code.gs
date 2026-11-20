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
    DASHBOARD: 'DASHBOARD', STAFF_PERF: 'STAFF PERFORMANCE', PROJECT_REPORT: 'PROJECT REPORT'
  },

  // TASK INPUT: posisi kolom tetap (1-based)
  INPUT: {
    HEADER_ROW: 4, FIRST_ROW: 5, LAST_ROW: 304, NCOLS: 16,
    COL: { NO: 1, TASK_ID: 2, INPUT_DATE: 3, CREATED_BY: 4, PROJECT: 5, PIC: 6, PARENT: 7, TITLE: 8, DESC: 9,
           PRIORITY: 10, START: 11, DEADLINE_DATE: 12, DEADLINE_TIME: 13, NOTES: 14, STATUS: 15, MESSAGE: 16 },
    SUBMITTED: 'SUBMITTED'
  },

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
           Q_LABEL: 7, Q_SCORE: 8, PERIOD: 9, ACTIVE: 10, YESNO: 11 },
    CFG_KEY_COL: 13, CFG_VAL_COL: 14, CFG_FIRST: 5, CFG_LAST: 21
  },
  // Nilai default bila key tidak ada di SETTINGS
  DEFAULTS: {
    SYSTEM_NAME: 'Staff Performance Tracking', TASK_ID_PREFIX: 'TSK', TASK_ID_DIGITS: '5', DUE_SOON_DAYS: '3',
    UPDATE_STALE_DAYS: '3', DEFAULT_DEADLINE_TIME: '17:00', ALLOW_PAST_DEADLINE: 'NO',
    ENFORCE_CREATED_BY_MATCH: 'NO', REQUIRE_DELAY_REASON: 'YES', ALLOW_REOPEN_CLOSED: 'NO', EMAIL_ENABLED: 'NO',
    EMAIL_NOTIFY_ON_SUBMIT: 'YES', EMAIL_NOTIFY_ON_PROGRESS: 'YES', EMAIL_CC_ADMINS: 'YES',
    REMINDER_ENABLED: 'YES', REMINDER_HOUR: '8', REMINDER_ADMIN_DIGEST: 'YES'
  },

  STATUS: { NEW: 'NEW', IN_PROGRESS: 'IN PROGRESS', PENDING: 'PENDING', COMPLETED: 'COMPLETED', CANCELLED: 'CANCELLED' },
  LOG_PREFIX: 'LOG', EVD_PREFIX: 'EVD',
  PROP: { TASK_SEQ: 'TASK_SEQ', LOG_SEQ: 'LOG_SEQ', EVD_SEQ: 'EVD_SEQ', LAST_DAILY_RUN: 'LAST_DAILY_RUN' },

  FORMATS: { DATETIME: 'dd/mm/yyyy hh:mm', DATE: 'dd/mm/yyyy', PERCENT: '0%' },

  // Judul pertanyaan Google Form (harus sama persis)
  FORMS: {
    EVIDENCE: { TASK_ID: 'Task ID', UPLOADED_BY: 'Uploaded By', DESCRIPTION: 'Evidence Description', FILES: 'Foto / File' },
    PROGRESS: { TASK_ID: 'Task ID', UPDATED_BY: 'Updated By', NEW_STATUS: 'New Status', PROGRESS: 'Progress Percentage',
                NOTES: 'Update Notes', DELAY: 'Delay Reason', REVISED: 'Revised Estimated Completion' }
  },

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
    try { return fn(); }
    catch (e) {
      Log.error(label, e);
      Util.alert(label + ' - GAGAL', String(e && e.message ? e.message : e) +
        '\n\nTidak ada data yang dihapus. Perbaiki penyebabnya lalu ulangi.');
      return null;
    }
  },

  /** Jalankan fn di dalam script lock. */
  withLock(fn) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      throw new Error('Sistem sedang dipakai pengguna lain. Tunggu beberapa detik lalu coba lagi.');
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

    const cfg = {};
    const cn = L.CFG_LAST - L.CFG_FIRST + 1;
    const keys = sh.getRange(L.CFG_FIRST, L.CFG_KEY_COL, cn, 1).getDisplayValues();
    const vals = sh.getRange(L.CFG_FIRST, L.CFG_VAL_COL, cn, 1).getDisplayValues();
    for (let i = 0; i < cn; i++) { const k = Util.str(keys[i][0]); if (k) cfg[k] = Util.str(vals[i][0]); }
    Object.keys(CONFIG.DEFAULTS).forEach(function (k) { if (!(k in cfg) || cfg[k] === '') cfg[k] = CONFIG.DEFAULTS[k]; });

    return { cfg: cfg, priorities: priorities, statuses: statuses, admins: admins, delayCategories: delayCategories,
             activeLabel: active[0] || 'Active' };
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
    const v = sh.getRange(2, 1, last - 1, 5).getDisplayValues();
    v.forEach(function (r) {
      const name = Util.str(r[1]);
      if (!name) return;
      const o = { id: Util.str(r[0]), name: name, c3: Util.str(r[2]), c4: Util.str(r[3]),
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
   * @param raw  {createdBy, project, pic, parent, title, desc, priority, inputDate, start, deadlineDate, timeText, notes}
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
    if (!title) E.push('Task Title wajib diisi'); else clean.title = title;
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
      let t = null;
      if (Util.str(raw.timeText) === '') t = Util.parseTime(S.cfg.DEFAULT_DEADLINE_TIME) || { h: 17, m: 0 };
      else { t = Util.parseTime(raw.timeText); if (!t) E.push('Format Deadline Time tidak valid (gunakan hh:mm, mis. 17:00)'); }
      if (t) {
        deadline = Util.combine(raw.deadlineDate, t.h, t.m);
        clean.deadline = deadline; clean.timeUsed = t;
        if (Util.startOfDay(deadline).getTime() < Util.startOfDay(start).getTime()) E.push('Deadline lebih awal dari Start Date');
        if (!Settings.yes(S, 'ALLOW_PAST_DEADLINE') && Util.startOfDay(deadline).getTime() < today.getTime()) {
          E.push('Deadline sudah lewat (ubah tanggal, atau set ALLOW_PAST_DEADLINE = YES di SETTINGS)');
        }
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
  run() {
    Util.assertTimezone();
    let result = null;
    Util.withLock(function () { result = Submit.process_(new Date()); });
    // email di luar lock; kegagalan email tidak membatalkan data
    let mail = { sent: 0, failed: [], skipped: 'tidak ada tugas baru' };
    if (result.submitted.length) {
      try { mail = Notify.onSubmit(result); } catch (e) { Log.error('Notify.onSubmit', e); mail = { sent: 0, failed: [String(e.message || e)], skipped: '' }; }
    }
    result.mail = mail;
    return result;
  },

  process_(now) {
    const C = CONFIG.INPUT.COL, IN = CONFIG.INPUT, H = CONFIG.DB_H;
    const S = Settings.load();
    const staff = Masters.staff(S), projects = Masters.projects(S);
    const db = Db.open();
    const dbData = Db.readAll(db);
    const sh = Util.sheet(CONFIG.SHEETS.INPUT);
    const n = IN.LAST_ROW - IN.FIRST_ROW + 1;
    const vals = sh.getRange(IN.FIRST_ROW, 1, n, C.STATUS).getValues();
    const timeTxt = sh.getRange(IN.FIRST_ROW, C.DEADLINE_TIME, n, 1).getDisplayValues();
    const email = Util.activeEmail();

    const result = { submitted: [], errors: [], healed: [], alreadySubmitted: 0, empty: 0, emailUsed: email };
    const batchFp = {};
    const ctx = { S: S, staff: staff, projects: projects, db: dbData, batchFp: batchFp, activeEmail: email, now: now };
    const valid = [];      // {rowIdx, rowNum, clean}
    const healRows = [];   // baris yang ID-nya sudah ada di database tetapi status belum SUBMITTED

    for (let i = 0; i < n; i++) {
      const v = vals[i]; const rowNum = IN.FIRST_ROW + i;
      if (Util.up(v[C.STATUS - 1]) === IN.SUBMITTED) { result.alreadySubmitted++; continue; }
      let hasData = false;
      for (let c = C.INPUT_DATE; c <= C.NOTES; c++) if (Util.str(v[c - 1]) !== '') { hasData = true; break; }
      if (!hasData) { result.empty++; continue; }

      const existingId = Util.up(v[C.TASK_ID - 1]);
      if (existingId && dbData.byId[existingId]) { healRows.push({ i: i, rowNum: rowNum, id: dbData.byId[existingId][H.TASK_ID] }); continue; }

      const raw = {
        createdBy: v[C.CREATED_BY - 1], project: v[C.PROJECT - 1], pic: v[C.PIC - 1], parent: v[C.PARENT - 1],
        title: v[C.TITLE - 1], desc: v[C.DESC - 1], priority: v[C.PRIORITY - 1], inputDate: v[C.INPUT_DATE - 1],
        start: v[C.START - 1], deadlineDate: v[C.DEADLINE_DATE - 1], timeText: timeTxt[i][0], notes: v[C.NOTES - 1]
      };
      const r = Validate.inputRow(raw, ctx);
      if (r.errors.length) { result.errors.push({ row: rowNum, messages: r.errors }); continue; }
      batchFp[r.clean.fp] = rowNum;
      valid.push({ i: i, rowNum: rowNum, clean: r.clean, raw: raw });
    }

    // Pulihkan baris yang sudah ada di database (mis. script terhenti setelah menulis database)
    healRows.forEach(function (h) {
      sh.getRange(h.rowNum, C.TASK_ID).setValue(h.id);
      sh.getRange(h.rowNum, C.STATUS).setValue(IN.SUBMITTED);
      result.healed.push({ row: h.rowNum, id: h.id });
    });

    if (!valid.length) return result;

    // --- reservasi ID, tulis database, lalu tandai baris input
    const existingIds = dbData.rows.map(function (o) { return o[H.TASK_ID]; });
    const prefix = S.cfg.TASK_ID_PREFIX || 'TSK';
    const digits = Math.max(1, Number(S.cfg.TASK_ID_DIGITS) || 5);
    const res = Ids.reserve(CONFIG.PROP.TASK_SEQ, existingIds, prefix, valid.length);

    const objs = valid.map(function (it, k) {
      const c = it.clean;
      const id = Ids.format(prefix, res.first + k, digits);
      it.id = id;
      const desc = c.notes ? (c.desc ? c.desc + '\n' : '') + 'Catatan: ' + c.notes : c.desc;
      const o = {};
      o[H.TASK_ID] = id; o[H.CREATED_AT] = now; o[H.CREATED_BY] = c.createdBy; o[H.PROJECT_ID] = c.project.id;
      o[H.PROJECT_NAME] = c.project.name; o[H.ASSIGNED_TO] = c.staff.name; o[H.PARENT] = c.parent || '';
      o[H.TITLE] = c.title; o[H.DESC] = desc; o[H.PRIORITY] = c.priority; o[H.START] = c.start; o[H.DEADLINE] = c.deadline;
      o[H.STATUS] = CONFIG.STATUS.NEW; o[H.PROGRESS] = 0; o[H.LAST_UPDATED] = now; o[H.SUBMITTED_BY] = email;
      return o;
    });

    Db.append(db, objs);        // jika gagal -> exception, input tidak ditandai, ID belum di-commit
    res.commit();

    valid.forEach(function (it) {
      const c = it.clean;
      sh.getRange(it.rowNum, C.TASK_ID).setValue(it.id);
      if (Util.str(it.raw.inputDate) === '') sh.getRange(it.rowNum, C.INPUT_DATE).setValue(c.inputDate);
      if (Util.str(it.raw.start) === '') sh.getRange(it.rowNum, C.START).setValue(c.start);
      if (Util.str(it.raw.timeText) === '') sh.getRange(it.rowNum, C.DEADLINE_TIME).setValue((c.timeUsed.h * 60 + c.timeUsed.m) / 1440);
      sh.getRange(it.rowNum, C.STATUS).setValue(IN.SUBMITTED);
      result.submitted.push({
        row: it.rowNum, id: it.id, title: c.title, project: c.project.name, pic: c.staff.name, picEmail: c.staff.email,
        priority: c.priority, deadline: c.deadline, createdBy: c.createdBy, adminEmail: c.adminEmail
      });
    });
    return result;
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
    if (r.mail) {
      if (r.mail.skipped) lines.push('Email: dilewati (' + r.mail.skipped + ')');
      else lines.push('Email terkirim: ' + r.mail.sent + (r.mail.failed.length ? ' | gagal: ' + r.mail.failed.join('; ') : ''));
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

  /** Dipanggil trigger Google Form (progress). */
  fromForm(nv) {
    const F = CONFIG.FORMS.PROGRESS, P = CONFIG.PL_H;
    const first = function (k) { return nv[k] && nv[k].length ? Util.str(nv[k][0]) : ''; };
    const sh = Util.sheet(CONFIG.SHEETS.PROGRESS); const h = Util.headerMap(sh, 1);
    const row = Math.max(Util.lastNonEmptyRow(sh, h[P.LOG_ID], 2), Util.lastNonEmptyRow(sh, h[P.TASK_ID], 2)) + 1;
    let p = first(F.PROGRESS); let pn = '';
    if (p !== '') { pn = Number(String(p).replace('%', '')); pn = isNaN(pn) ? '' : (pn > 1 ? pn / 100 : pn); }
    const rev = first(F.REVISED);
    sh.getRange(row, h[P.TASK_ID]).setValue(first(F.TASK_ID));
    sh.getRange(row, h[P.UPDATED_BY]).setValue(first(F.UPDATED_BY));
    sh.getRange(row, h[P.NEW]).setValue(first(F.NEW_STATUS));
    if (pn !== '') sh.getRange(row, h[P.PROGRESS]).setValue(pn);
    sh.getRange(row, h[P.NOTES]).setValue(first(F.NOTES));
    sh.getRange(row, h[P.DELAY]).setValue(first(F.DELAY));
    if (rev) { const d = new Date(rev); if (!isNaN(d.getTime())) sh.getRange(row, h[P.REVISED]).setValue(d); }
    return Progress.run();
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

  send(S, to, cc, subject, html) {
    if (!Notify.validEmail(to)) return { ok: false, error: 'email tujuan tidak valid: "' + to + '"' };
    try {
      if (MailApp.getRemainingDailyQuota() < 1) return { ok: false, error: 'kuota email harian habis' };
      const opt = { to: to, subject: '[' + S.cfg.SYSTEM_NAME + '] ' + subject, htmlBody: html, name: S.cfg.SYSTEM_NAME };
      if (cc) opt.cc = cc;
      MailApp.sendEmail(opt);
      return { ok: true };
    } catch (e) { Log.error('Notify.send', e); return { ok: false, error: String(e.message || e) }; }
  },

  table(headers, rows) {
    const th = headers.map(function (x) { return '<th style="background:#1F2937;color:#fff;padding:4px 8px;text-align:left;font-size:12px">' + Util.esc(x) + '</th>'; }).join('');
    const tr = rows.map(function (r) {
      return '<tr>' + r.map(function (c) { return '<td style="border-bottom:1px solid #ddd;padding:4px 8px;font-size:12px">' + Util.esc(c) + '</td>'; }).join('') + '</tr>';
    }).join('');
    return '<table style="border-collapse:collapse;font-family:Arial,sans-serif"><tr>' + th + '</tr>' + tr + '</table>';
  },

  onSubmit(result) {
    const S = Settings.load();
    const out = { sent: 0, failed: [], skipped: '' };
    if (!Notify.enabled(S)) { out.skipped = 'EMAIL_ENABLED = NO'; return out; }
    if (!Settings.yes(S, 'EMAIL_NOTIFY_ON_SUBMIT')) { out.skipped = 'EMAIL_NOTIFY_ON_SUBMIT = NO'; return out; }
    const groups = {};
    result.submitted.forEach(function (t) { (groups[t.pic] = groups[t.pic] || []).push(t); });
    Object.keys(groups).forEach(function (pic) {
      const list = groups[pic];
      const cc = Settings.yes(S, 'EMAIL_CC_ADMINS')
        ? list.map(function (t) { return t.adminEmail; }).filter(function (e, i, a) { return Notify.validEmail(e) && a.indexOf(e) === i; }).join(',') : '';
      const html = '<p>Halo ' + Util.esc(pic) + ', Anda mendapat ' + list.length + ' tugas baru:</p>' +
        Notify.table(['Task ID', 'Tugas', 'Proyek', 'Prioritas', 'Deadline', 'Pemberi tugas'],
          list.map(function (t) { return [t.id, t.title, t.project, t.priority, Util.fmt(t.deadline), t.createdBy]; })) +
        '<p style="font-size:12px;color:#666">Update progres melalui sheet PROGRESS LOG.</p>';
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
        (p.delay ? '<p>Alasan keterlambatan: ' + Util.esc(p.delay) + '</p>' : '') + (p.notes ? '<p>Catatan: ' + Util.esc(p.notes) + '</p>' : '');
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
    const rowsOf = function (a, lbl) { return a.map(function (t) { return [t.id, t.title, t.project, Util.fmt(t.deadline), t.days + ' hari ' + lbl]; }); };
    Object.keys(byPic).forEach(function (pic) {
      const s = staff.byName[Util.low(pic)];
      if (!s || !s.active) return;
      const g = byPic[pic]; let html = '<p>Halo ' + Util.esc(pic) + ',</p>';
      if (g.overdue.length) html += '<p><b style="color:#C00000">OVERDUE (' + g.overdue.length + ')</b></p>' + Notify.table(['Task ID', 'Tugas', 'Proyek', 'Deadline', 'Terlambat'], rowsOf(g.overdue, 'terlambat'));
      if (g.dueSoon.length) html += '<p><b>Mendekati deadline (' + g.dueSoon.length + ')</b></p>' + Notify.table(['Task ID', 'Tugas', 'Proyek', 'Deadline', 'Sisa'], rowsOf(g.dueSoon, 'lagi'));
      const r = Notify.send(S, s.email, '', 'Reminder: ' + g.overdue.length + ' overdue, ' + g.dueSoon.length + ' mendekati deadline', html);
      if (r.ok) out.sent++; else out.failed.push(pic + ': ' + r.error);
    });
    if (Settings.yes(S, 'REMINDER_ADMIN_DIGEST') && (det.overdue.length || det.dueSoon.length)) {
      const all = det.overdue.map(function (t) { return [t.id, t.title, t.pic, t.project, Util.fmt(t.deadline), 'OVERDUE ' + t.days + ' hari']; })
        .concat(det.dueSoon.map(function (t) { return [t.id, t.title, t.pic, t.project, Util.fmt(t.deadline), 'sisa ' + t.days + ' hari']; }));
      const html = '<p>Ringkasan harian ' + Util.fmt(now, 'dd/MM/yyyy') + ': ' + det.overdue.length + ' overdue, ' + det.dueSoon.length + ' mendekati deadline.</p>' +
        Notify.table(['Task ID', 'Tugas', 'PIC', 'Proyek', 'Deadline', 'Status'], all);
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
      if (['dailyJob', 'onFormSubmitRouter'].indexOf(t.getHandlerFunction()) >= 0) { ScriptApp.deleteTrigger(t); n++; }
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
    [1, 2, 15, 16].forEach(function (c) { warn(inp.getRange(IN.FIRST_ROW, c, IN.LAST_ROW - IN.FIRST_ROW + 1, 1), 'TASK INPUT kolom otomatis'); });
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
    add('Email akun aktif terbaca', !!Util.activeEmail(), Util.activeEmail() || 'kosong (akun mungkin tidak mengizinkan)');
    try { add('Trigger harian terpasang', Triggers.has('dailyJob'), Triggers.has('dailyJob') ? '' : 'jalankan menu Install Triggers'); } catch (e) { add('Trigger harian', false, 'izin belum diberikan: ' + e.message); }
    try { add('Kuota email harian', MailApp.getRemainingDailyQuota() > 0, String(MailApp.getRemainingDailyQuota())); } catch (e) { add('Kuota email', false, e.message); }
    return out;
  },
  text(r) {
    const bad = r.filter(function (x) { return !x.ok; }).length;
    return (bad ? 'ADA ' + bad + ' MASALAH' : 'SEMUA PEMERIKSAAN LULUS') + '\n\n' +
      r.map(function (x) { return (x.ok ? '[OK]   ' : '[GAGAL] ') + x.name + (x.detail ? ' - ' + x.detail : ''); }).join('\n');
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
      .addItem('Clear Submitted Rows (TASK INPUT)', 'clearSubmittedRows')
      .addItem('Detect Overdue (laporan)', 'detectOverdueReport')
      .addItem('Run Daily Reminder Now', 'runDailyReminderNow')
      .addSeparator()
      .addItem('Install Triggers (reminder harian)', 'installTriggers')
      .addItem('Install Form Trigger', 'installFormTrigger')
      .addItem('Remove Triggers', 'removeTriggers')
      .addItem('Apply Protections (peringatan)', 'applyProtections')
      .addSeparator()
      .addItem('Self Test (cek instalasi)', 'selfTest')
      .addItem('Send Test Email', 'sendTestEmail')
      .addToUi();
  } catch (e) { Log.error('onOpen', e); }
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
  try { return Reminders.run(false); } catch (e) { Log.error('dailyJob', e); return null; }
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

/** Trigger installable Google Form (Install Form Trigger): evidence atau progress. */
function onFormSubmitRouter(e) {
  try {
    const nv = e && e.namedValues ? e.namedValues : {};
    if (nv[CONFIG.FORMS.EVIDENCE.FILES]) return Evidence.fromForm(nv);
    if (nv[CONFIG.FORMS.PROGRESS.NEW_STATUS]) return Progress.fromForm(nv);
    Log.info('onFormSubmitRouter', 'form tidak dikenali: ' + Object.keys(nv).join(', '));
  } catch (err) { Log.error('onFormSubmitRouter', err); }
  return null;
}
