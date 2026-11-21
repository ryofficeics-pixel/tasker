// ===== mock extensions (turn 4): patch, web app, Drive, Cache =====
const colNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
const a1 = s => { const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(s); if (!m) throw new Error('bad A1 ' + s);
  const c1 = colNum(m[1]), r1 = +m[2]; const c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1; return [r1, c1, r2 - r1 + 1, c2 - c1 + 1]; };
const _gr = Sheet.prototype.getRange;
Sheet.prototype.getRange = function (r, c, nr, nc) { if (typeof r === 'string') { const x = a1(r); return _gr.call(this, x[0], x[1], x[2], x[3]); } return _gr.call(this, r, c, nr, nc); };
['setFontWeight', 'setFontColor', 'setBackground', 'setHorizontalAlignment', 'setFontStyle', 'setFontSize', 'setDataValidation'].forEach(m => { Range.prototype[m] = function () { (this.sh.styleCalls = this.sh.styleCalls || []).push(m); return this; }; });
Range.prototype.setFormulas = function (a) { a.forEach((row, i) => row.forEach((f, j) => { const k = (this.r + i) + ',' + (this.c + j); this.sh.f.set(k, f); this.sh.cells.delete(k); })); return this; };
Range.prototype.getFormulas = function () { return this._each((r, c) => this.sh.f.get(r + ',' + c) || ''); };
const _setVal = Range.prototype._set;
Range.prototype._set = function (r, c, v) { this.sh.f.delete(r + ',' + c); return _setVal.call(this, r, c, v); };
Sheet.prototype._shift = function (col, dir) {   // dir -1: delete col ; +1: insert before col
  [this.cells, this.f].forEach(map => { const ent = [...map.entries()]; map.clear();
    ent.forEach(([k, v]) => { const [r, c] = k.split(',').map(Number);
      if (dir < 0) { if (c === col) return; map.set(r + ',' + (c > col ? c - 1 : c), v); } else map.set(r + ',' + (c >= col ? c + 1 : c), v); }); });
};
Sheet.prototype.deleteColumn = function (c) { this._shift(c, -1); };
Sheet.prototype.insertColumnBefore = function (c) { this._shift(c, +1); };
Sheet.prototype.deleteRow = function (r) { [this.cells, this.f].forEach(m => { const e = [...m.entries()]; m.clear(); e.forEach(([k, v]) => { const [rr, cc] = k.split(',').map(Number); if (rr === r) return; m.set((rr > r ? rr - 1 : rr) + ',' + cc, v); }); }); };
Sheet.prototype.setColumnWidth = function () {}; Sheet.prototype.setFrozenRows = function () {}; Sheet.prototype.setTabColor = function () {};
Spreadsheet.prototype.insertSheet = function (n) { const sh = new Sheet(n, this); this.sheets.push(sh); return sh; };
Spreadsheet.prototype.getId = function () { return 'SSID'; }; Spreadsheet.prototype.getName = function () { return 'Staff Tracking'; };
SpreadsheetApp.CopyPasteType = { PASTE_FORMAT: 'F' };
SpreadsheetApp.newDataValidation = () => { const b = { requireValueInRange() { return b; }, setAllowInvalid() { return b; }, build() { return {}; } }; return b; };
const _ui = SpreadsheetApp.getUi;
SpreadsheetApp.getUi = () => { const u = _ui(); const mm = () => ({ addItem() { return this; }, addSeparator() { return this; }, addSubMenu() { return this; }, addToUi() {} });
  u.createMenu = mm; return u; };
Utilities.formatDate = function (d, tz, pat) { const f = new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' });
  const o = {}; f.formatToParts(d).forEach(p => o[p.type] = p.value);
  return pat.replace('yyyy', o.year).replace('MM', o.month).replace('dd', o.day).replace('HH', o.hour).replace('mm', o.minute).replace('ss', o.second); };
let UUIDN = 0;
Utilities.getUuid = () => { const h = (Math.random().toString(16).slice(2) + '0000000000000000').slice(0, 12) + String(++UUIDN).padStart(20, '0'); return h.slice(0, 8) + '-' + h.slice(8, 12) + '-4' + h.slice(13, 16) + '-a' + h.slice(17, 20) + '-' + h.slice(20, 32); };
Utilities.base64Decode = s => [...Buffer.from(s, 'base64')].map(b => b > 127 ? b - 256 : b);
Utilities.base64Encode = b => Buffer.from(b.map(x => x & 255)).toString('base64');
Utilities.newBlob = (bytes, mime, name) => ({ bytes, mime, name, getBytes: () => bytes, getContentType: () => mime, getName: () => name });
const CACHE = {}; const CacheService = { getScriptCache: () => ({ get: k => (k in CACHE ? CACHE[k] : null), put: (k, v) => { CACHE[k] = v; }, remove: k => { delete CACHE[k]; } }) };
const DRIVE = { files: {}, folders: {}, copies: [], n: 0 };
const mkFile = (blob, parent) => { const id = 'F' + (++DRIVE.n) + 'x'; const f = { id, blob, parent, sharing: null, viewers: [], trashed: false, getId: () => id, getUrl: () => 'https://drive.google.com/file/d/' + id + '/view',
  setSharing(a, p) { if (DRIVE.denySharing) throw new Error('policy'); f.sharing = a + ':' + p; return f; }, addViewer(e) { f.viewers.push(e); return f; }, getBlob: () => f.blob, setTrashed(b) { f.trashed = b; } }; DRIVE.files[id] = f; return f; };
const mkFolder = (name, parent) => { const id = 'D' + (++DRIVE.n) + 'x'; const fo = { id, name, parent, sharing: null, viewers: [], trashed: false, children: [], getId: () => id, getName: () => name, isTrashed: () => fo.trashed,
  getUrl: () => 'https://drive.google.com/drive/folders/' + id, setSharing(a, p) { fo.sharing = a + ':' + p; return fo; }, addViewer(e) { fo.viewers.push(e); return fo; },
  getFoldersByName(n) { const m = fo.children.filter(c => c.name === n && !c.trashed); let i = 0; return { hasNext: () => i < m.length, next: () => m[i++] }; },
  createFolder(n) { const c = mkFolder(n, fo); fo.children.push(c); return c; }, createFile(b) { const f = mkFile(b, fo); fo.files = (fo.files || []).concat([f]); return f; } }; DRIVE.folders[id] = fo; return fo; };
const DriveApp = { Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' }, Permission: { VIEW: 'VIEW' },
  createFolder: n => mkFolder(n, null), getFolderById: id => { if (!DRIVE.folders[id]) throw new Error('no folder'); return DRIVE.folders[id]; },
  getFileById: id => { if (id === 'SSID') return { makeCopy: n => { if (DRIVE.denyCopy) throw new Error('Drive denied'); DRIVE.copies.push(n); return { getUrl: () => 'https://drive/copy/' + DRIVE.copies.length }; } };
    if (!DRIVE.files[id]) throw new Error('no file'); return DRIVE.files[id]; } };
const HtmlService = { XFrameOptionsMode: { ALLOWALL: 1 }, createTemplateFromFile: n => { const t = { n, evaluate() { return { setTitle() { return this; }, addMetaTag() { return this; }, setXFrameOptionsMode() { return this; }, tpl: t }; } }; return t; } };
ScriptApp.getService = () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' });

// plain-text format ('@') & perilaku "=" seperti Google Sheets asli
Range.prototype.setNumberFormat = function (f) { this._each((r, c) => { const k = r + ',' + c; this.sh.textCells = this.sh.textCells || new Set(); if (f === '@') this.sh.textCells.add(k); else this.sh.textCells.delete(k); }); return this; };
const _set2 = Range.prototype._set;
Range.prototype._set = function (r, c, v) {
  const k = r + ',' + c;
  if (typeof v === 'string' && v.startsWith('=') && !(this.sh.textCells && this.sh.textCells.has(k))) { this.sh.f.set(k, v); this.sh.cells.delete(k); return; }
  return _set2.call(this, r, c, v);
};
Utilities.sleep = function (ms) { Utilities._slept = (Utilities._slept || 0) + ms; };
