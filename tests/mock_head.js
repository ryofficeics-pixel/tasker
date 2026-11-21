const fs = require('fs');
const DUMP = JSON.parse(fs.readFileSync('/home/claude/build/dump.json', 'utf8'));
const TZ = 'Asia/Jakarta';
const p2 = n => String(n).padStart(2, '0');
function conv(v) {
  if (v && typeof v === 'object') {
    if (v.f !== undefined) return { f: v.f };
    if (v.d !== undefined) { const m = /^(\d+)-(\d+)-(\d+)(?:T(\d+):(\d+))?/.exec(v.d); return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)); }
    if (v.t !== undefined) return { t: v.t };
  }
  return v;
}
class Sheet {
  constructor(name, ss) { this.name = name; this.ss = ss; this.cells = new Map(); this.f = new Map(); }
  getName() { return this.name; }
  _has(r, c) { return this.cells.has(r + ',' + c) || this.f.has(r + ',' + c); }
  getLastRow() { let m = 0; for (const k of [...this.cells.keys(), ...this.f.keys()]) m = Math.max(m, +k.split(',')[0]); return m; }
  getLastColumn() { let m = 0; for (const k of [...this.cells.keys(), ...this.f.keys()]) m = Math.max(m, +k.split(',')[1]); return m; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
}
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getSheet() { return this.sh; } getRow() { return this.r; } getColumn() { return this.c; }
  getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  _each(fn) { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(fn(this.r + i, this.c + j)); out.push(row); } return out; }
  _raw(r, c) { const k = r + ',' + c; if (this.sh.f.has(k)) return ''; const v = this.sh.cells.get(k); return v === undefined ? '' : v; }
  getValues() { return this._each((r, c) => { const v = this._raw(r, c); if (v && v.t) { const [h, m] = v.t.split(':').map(Number); return new Date(1899, 11, 30, h, m); } return v; }); }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValues() { return this._each((r, c) => { const v = this._raw(r, c); if (v && v.t) return v.t;
    if (v instanceof Date) { const d = `${p2(v.getDate())}/${p2(v.getMonth() + 1)}/${v.getFullYear()}`; return (v.getHours() || v.getMinutes()) ? d + ` ${p2(v.getHours())}:${p2(v.getMinutes())}` : d; }
    if (typeof v === 'number' && this.sh.name === 'TASK INPUT' && c === 13) { const mins = Math.round(v * 1440); return `${p2(Math.floor(mins / 60))}:${p2(mins % 60)}`; }
    return String(v); }); }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  getFormula() { return this.sh.f.get(this.r + ',' + this.c) || ''; }
  _set(r, c, v) { const k = r + ',' + c; if (v === '' || v === null || v === undefined) { this.sh.cells.delete(k); return; }
    if (typeof v === 'number' && this.sh.name === 'TASK INPUT' && c === 13) v = { t: `${p2(Math.floor(Math.round(v * 1440) / 60))}:${p2(Math.round(v * 1440) % 60)}` };
    if (typeof v === 'string' && this.sh.name === 'TASK INPUT' && c === 13 && /^\d{1,2}:\d{2}$/.test(v)) v = { t: v };
    this.sh.cells.set(k, v); }
  setValues(a) { if (a.length !== this.nr || a[0].length !== this.nc) throw new Error('setValues dimension mismatch ' + a.length + 'x' + a[0].length + ' vs ' + this.nr + 'x' + this.nc);
    a.forEach((row, i) => row.forEach((v, j) => this._set(this.r + i, this.c + j, v))); return this; }
  setValue(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this._set(this.r + i, this.c + j, v); return this; }
  setNumberFormat() { return this; }
  clearContent() { this._each((r, c) => { this.sh.cells.delete(r + ',' + c); }); return this; }
  copyTo(dest) { const src = []; for (let j = 0; j < this.nc; j++) src.push(this.sh.f.get(this.r + ',' + (this.c + j))); for (let i = 0; i < dest.nr; i++) for (let j = 0; j < dest.nc; j++) { const f = src[j % this.nc]; if (f) this.sh.f.set((dest.r + i) + ',' + (dest.c + j), f); } }
}
class Spreadsheet {
  constructor() { this.sheets = []; this.toasts = [];
    DUMP.names.forEach(n => { const sh = new Sheet(n, this); const cells = DUMP.sheets[n].cells;
      for (const k in cells) { const v = conv(cells[k]); if (v && v.f !== undefined) sh.f.set(k, v.f); else sh.cells.set(k, v); } this.sheets.push(sh); }); }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  getSheets() { return this.sheets; }
  getSpreadsheetTimeZone() { return TZ; }
  toast(m) { this.toasts.push(m); }
}
let SS = new Spreadsheet();
const ALERTS = [];
const SpreadsheetApp = {
  getActiveSpreadsheet: () => SS,
  getUi: () => ({ alert: (t, m) => { ALERTS.push([t, m]); return 'OK'; }, ButtonSet: { OK: 1, YES_NO: 2 }, Button: { YES: 'YES' }, createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }) })
};
const Utilities = { formatDate(d, tz, pat) { const f = new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' });
  const o = {}; f.formatToParts(d).forEach(p => o[p.type] = p.value);
  return pat.replace('yyyy', o.year).replace('MM', o.month).replace('dd', o.day).replace('HH', o.hour).replace('mm', o.minute); } };
let ACTIVE_EMAIL = 'manager@example.com';
const Session = { getScriptTimeZone: () => TZ, getActiveUser: () => ({ getEmail: () => ACTIVE_EMAIL }), getEffectiveUser: () => ({ getEmail: () => ACTIVE_EMAIL }) };
let LOCKED = false;
const LockService = { getScriptLock: () => ({ tryLock: () => { if (LOCKED) return false; LOCKED = true; return true; }, releaseLock: () => { LOCKED = false; } }) };
const PROPS = {};
const PropertiesService = { getScriptProperties: () => ({ getProperty: k => (k in PROPS ? PROPS[k] : null), setProperty: (k, v) => { PROPS[k] = v; } }) };
const MAILS = [];
const MailApp = { sendEmail: o => { MAILS.push(o); }, getRemainingDailyQuota: () => 100 - MAILS.length };
const TRIGGERS = [];
const ScriptApp = { getProjectTriggers: () => TRIGGERS.slice(), deleteTrigger: t => { const i = TRIGGERS.indexOf(t); if (i >= 0) TRIGGERS.splice(i, 1); },
  newTrigger: fn => { const b = { fn, kind: '', getHandlerFunction() { return fn; }, forForm(f) { this.form = f; return this; }, forSpreadsheet() { return this; }, onFormSubmit() { this.kind = 'form'; return this; },
    timeBased() { this.kind = 'time'; return this; }, everyDays() { return this; }, atHour() { return this; }, inTimezone() { return this; }, create() { TRIGGERS.push(this); return this; } }; return b; } };
class FItem { constructor(t) { this.type = t; this.title = ''; this.choices = []; this.req = false; this.help = ''; }
  setTitle(t) { this.title = t; return this; } getTitle() { return this.title; } setRequired(b) { this.req = b; return this; } setHelpText(t) { this.help = t; return this; }
  setChoiceValues(v) { if (!v.length) throw new Error('empty choices'); this.choices = v.slice(); return this; } getType() { return this.type; } asListItem() { return this; } asMultipleChoiceItem() { return this; } }
const FORMS = {}; let FSEQ = 0;
class FForm { constructor(n) { this.name = n; this.id = 'form' + (++FSEQ); this.items = []; this.conf = {}; FORMS[this.id] = this; }
  _add(t) { const i = new FItem(t); this.items.push(i); return i; }
  addListItem() { return this._add('LIST'); } addMultipleChoiceItem() { return this._add('MC'); } addDateItem() { return this._add('DATE'); }
  addTextItem() { return this._add('TEXT'); } addParagraphTextItem() { return this._add('PARA'); } getItems() { return this.items; }
  setDescription() { return this; } setConfirmationMessage() { return this; } setLimitOneResponsePerUser() { return this; } setShowLinkToRespondAgain() { return this; }
  setAllowResponseEdits() { return this; } setProgressBar() { return this; } setEmailCollectionType(t) { this.conf.email = t; return this; } setCollectEmail() { return this; } setRequireLogin() { return this; }
  getId() { return this.id; } getPublishedUrl() { return 'https://forms.example/' + this.id + '/viewform'; } getEditUrl() { return 'https://forms.example/' + this.id + '/edit'; } }
const FormApp = { create: n => new FForm(n), openById: id => { if (!FORMS[id]) throw new Error('no form'); return FORMS[id]; },
  ItemType: { LIST: 'LIST', MULTIPLE_CHOICE: 'MC' }, EmailCollectionType: { VERIFIED: 'VERIFIED' } };
const mkResp = (email, map) => ({ getRespondentEmail: () => email, getItemResponses: () => Object.keys(map).map(k => ({ getItem: () => ({ getTitle: () => k }), getResponse: () => map[k] })) });

const Logger = { log: () => {} };
