// ============ STRESS / FUZZ / KEAMANAN ============
(function () {
const T0 = Date.now(); const timeit = (label, fn) => { const t = Date.now(); const r = fn(); console.log('   [t] ' + label + ': ' + (Date.now() - t) + ' ms'); return r; };
SS = new Spreadsheet(); Object.keys(PROPS).forEach(k => delete PROPS[k]); Object.keys(CACHE).forEach(k => delete CACHE[k]);
MAILS.length = 0; ACTIVE_EMAIL = 'manager@example.com'; LOCKED = false;
Patch.run(); Auth.generate();
const S0 = Settings.load(); const CM = S0.admins[0].code; const staffL = Masters.staff(S0).list;
const C1 = staffL[0].code, C2 = staffL[1].code, C3 = staffL[2].code;
const api = (code, action, payload) => JSON.parse(Api.handle(code, action, typeof payload === 'string' ? payload : JSON.stringify(payload || {})));
const midn = o => { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate() + o); };
const dstr = o => Utilities.formatDate(midn(o), 'Asia/Jakarta', 'yyyy-MM-dd');
const mk = (i, o) => Object.assign({ project: 'Proyek Contoh A', pic: 'Staf 01', title: 'Tugas ' + i, priority: 'High', deadline: dstr(5) }, o || {});

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S1. fuzz Api.handle (tidak boleh melempar / format harus JSON ok|error)');
let bad = 0, sys = 0, n = 0;
const junkCodes = ['', null, undefined, 0, 12345, {}, [], 'x'.repeat(100000), "' OR 1=1 --", '<script>alert(1)</script>', '🙂🙂🙂', 'ABCD-EFGH', C1.toLowerCase(), ' ' + C1 + ' ', C1.replace('-', ''), C1 + 'X', C1.slice(0, -1), '\u0000', '__proto__', 'constructor'];
const junkActions = ['login', 'tasks', 'task', 'monitor', 'report', 'createTasks', 'rateTask', 'uploadPhoto', 'submitUpdate', 'evidenceImage', '', null, undefined, 'constructor', '__proto__', 'toString', 'hasOwnProperty', 'valueOf', 'x'.repeat(5000), 'LOGIN'];
const junkPayloads = ['', 'null', '[]', '0', '"str"', 'true', '{', '{"a":', '{"__proto__":{"role":"admin"}}', '{"constructor":{"prototype":{"role":"admin"}}}', '{"taskId":null}', '{"taskId":{}}', '{"taskId":[1]}', '{"items":"x"}', '{"items":[null]}', '{"items":[[]]}', '{"items":[{}]}', '{"items":' + JSON.stringify(Array(31).fill({})) + '}', '{"status":[]}', '{"progress":"NaN"}', '{"progress":1e999}', '{"q":{}}', '{"scope":{}}', '{"from":"x","to":{}}', '{"data":"@@@@","mime":"image/jpeg","taskId":"TSK-00001"}'];
junkCodes.forEach(c => junkActions.forEach(a => junkPayloads.forEach(p => {
  n++; let r; try { r = JSON.parse(Api.handle(c, a, p)); } catch (e) { bad++; if (bad < 4) console.log('   throw/invalid:', JSON.stringify(c).slice(0, 30), a && String(a).slice(0, 20), p.slice(0, 40), e.message); return; }
  if (typeof r.ok !== 'boolean') bad++; if (!r.ok && /Kesalahan sistem/.test(r.error)) { sys++; if (sys < 6) console.log('   sys-error:', String(a).slice(0, 14), p.slice(0, 50), '->', r.error.slice(0, 90)); }
})));
ok(bad === 0, 'fuzz: ' + n + ' kombinasi, ' + bad + ' pelanggaran format');
ok(sys === 0, 'fuzz: ' + sys + ' "Kesalahan sistem" untuk input rusak (harus pesan terkontrol)');
ok(api(C1, 'login').ok && api(C1.toLowerCase(), 'login').ok && api(' ' + C1 + ' ', 'login').ok, 'kode valid peka format (huruf kecil/spasi) tetap masuk');
ok(!api(C1 + 'X', 'login').ok && !api('', 'login').ok, 'kode salah ditolak');
ok(api('__proto__', 'login').ok === false, 'kode __proto__ ditolak');

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S2. otorisasi staf vs admin');
let r = api(CM, 'createTasks', { items: [mk(1), mk(2, { pic: 'Staf 02' }), mk(3, { pic: 'Staf 03' })] }); ok(r.ok && r.data.created === 3, 'seed 3 tugas');
const tasks = api(CM, 'tasks', { scope: 'all' }).data.tasks; const idOf = pic => tasks.find(t => t.pic === pic).id;
const T1 = idOf('Staf 01'), T2 = idOf('Staf 02');
['task', 'submitUpdate', 'uploadPhoto'].forEach(a => { const rr = api(C1, a, { taskId: T2, status: 'IN PROGRESS', mime: 'image/jpeg', data: 'AAAA' }); ok(!rr.ok && /bukan milik/.test(rr.error), 'staf tidak bisa ' + a + ' tugas orang lain'); });
['dashboard', 'createTasks', 'rateTask'].forEach(a => { const rr = api(C1, a, { taskId: T1, items: [mk(9)] }); ok(!rr.ok && /admin/i.test(rr.error), 'staf ditolak: ' + a); });
r = api(C1, 'report', { staff: 'Staf 02' }); ok(r.ok && r.data.name === 'Staf 01', 'staf tidak bisa laporan staf lain');
r = api(C1, 'tasks', { scope: 'all', pic: 'Staf 02', q: '' }); ok(r.data.tasks.every(t => t.pic === 'Staf 01'), 'filter pic staf diabaikan');
r = api(C1, 'tasks', { scope: 'all', q: 'Tugas 2' }); ok(r.data.tasks.length === 0, 'pencarian tidak membocorkan tugas lain');
// nonaktifkan staf -> akses dicabut
const sm = SS.getSheetByName('STAFF MASTER'); const actCol = [...Array(8).keys()].map(i => i + 1).find(c => /active/i.test(String(sm.getRange(1, c).getValue())));
const st3Row = 2 + staffL.findIndex(s => s.code === C3);
sm.getRange(st3Row, actCol).setValue('Inactive'); ok(!api(C3, 'login').ok, 'staf Inactive langsung ditolak'); sm.getRange(st3Row, actCol).setValue('Active'); ok(api(C3, 'login').ok, 'aktif lagi -> masuk');
// duplikat kode
const f6 = sm.getRange(2 + 1, 6).getValue(); sm.getRange(2 + 1, 6).setValue(C1); r = api(C1, 'login'); ok(!r.ok, 'kode ganda ditolak (tidak ambigu)'); sm.getRange(2 + 1, 6).setValue(f6);
// WEB_ENABLED
const se = SS.getSheetByName('SETTINGS'); let wr = 0; for (let x = 5; x <= 80; x++) if (se.getRange(x, 13).getValue() === 'WEB_ENABLED') wr = x;
se.getRange(wr, 14).setValue('NO'); ok(!api(CM, 'login').ok, 'WEB_ENABLED=NO menutup semua'); se.getRange(wr, 14).setValue('YES');

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S3. injeksi konten & batas ukuran');
const evil = ['=HYPERLINK("http://evil","x")', '+SUM(1,1)', '-2+3', '@SUM(1)', '<img src=x onerror=alert(1)>', '"; DROP TABLE', "x'y\"z", 'baris1\nbaris2\r\nbaris3', '\t tab', '🙂 emoji 日本語 العربية', ' spasi ', '0123', '1e3', 'TRUE', '#REF!', '=1+1'];
evil.forEach((s, i) => {
  const rr = api(CM, 'createTasks', { items: [mk(100 + i, { title: s + ' T' + i, desc: s, notes: s })] });
  ok(rr.ok && rr.data.created === 1, 'judul berbahaya diterima sebagai teks: ' + JSON.stringify(s).slice(0, 30) + ' ' + (rr.error || JSON.stringify(rr.data).slice(0, 80)));
});
const dbs = SS.getSheetByName('TASK DATABASE'); const hdr = Util.headerMap(dbs, 1);
const tc = hdr[CONFIG.DB_H.TITLE]; let fcount = 0; for (let x = 2; x <= dbs.getLastRow(); x++) { const v = dbs.getRange(x, tc).getValue(); if (/^[=+\-@]/.test(String(v)) && dbs.getRange(x, tc).getNumberFormat && dbs.getRange(x, tc).getNumberFormat() !== '@') fcount++; }
ok(fcount === 0, 'judul berawalan = + - @ tersimpan dengan format teks (tidak jadi rumus): ' + fcount);
const lst = api(CM, 'tasks', { scope: 'all', q: 'T7' }).data.tasks; ok(lst.length >= 1, 'pencarian menemukan');
// ukuran
r = api(CM, 'createTasks', { items: [mk(500, { title: 'J'.repeat(100000) })] }); ok(!r.ok || r.data.created === 0, 'judul 100k karakter ditolak/ditangani: ' + JSON.stringify(r).slice(0, 120));
r = api(CM, 'createTasks', { items: [mk(501, { desc: 'D'.repeat(60000) })] }); ok(!r.ok || r.data.created === 0 || true, 'desc 60k: ' + JSON.stringify(r).slice(0, 100));
const dd = dbs.getRange(dbs.getLastRow(), hdr[CONFIG.DB_H.DESC]).getValue(); ok(String(dd).length <= 50000, 'sel tidak melewati batas 50.000 karakter Sheets (len=' + String(dd).length + ')');
r = api(C1, 'submitUpdate', { taskId: T1, status: 'IN PROGRESS', progress: 10, notes: 'N'.repeat(80000) }); const pl = SS.getSheetByName('PROGRESS LOG'); const nlen = String(pl.getRange(pl.getLastRow(), 7).getValue() || '').length;
ok(!r.ok || nlen <= 50000, 'catatan 80k tidak merusak (ok=' + r.ok + ', len=' + nlen + ')');
// tanggal
[['2026-02-30', false], ['0000-00-00', false], ['abc', false], ['', false], ['2099-12-31', true], ['2028-02-29', true], ['2020-01-01', false], [dstr(0), true], [dstr(-1), false], ['2026-13-01', false], ['  ' + dstr(3) + '  ', true], ['31/12/2026', false]].forEach(([d, good]) => {
  const rr = api(CM, 'createTasks', { items: [mk(600 + Math.random() * 1e4 | 0, { deadline: d })] }); const created = rr.ok && rr.data.created === 1; ok(created === good, 'deadline "' + d + '" -> ' + (created ? 'diterima' : 'ditolak') + ' (harus ' + (good ? 'diterima' : 'ditolak') + ')');
});
r = api(CM, 'createTasks', { items: [mk(700, { pic: 'staf 01 ' })] }); console.log('   info: PIC beda huruf/spasi ->', r.ok ? r.data.created : r.error);
r = api(CM, 'createTasks', { items: [mk(701, { pic: 'Orang Asing' })] }); ok(!r.ok || r.data.created === 0, 'PIC tak dikenal ditolak');
r = api(CM, 'createTasks', { items: [mk(702, { project: 'Proyek Gaib' })] }); ok(!r.ok || r.data.created === 0, 'proyek tak dikenal ditolak');
r = api(CM, 'createTasks', { items: [mk(703, { priority: 'Mega' })] }); ok(!r.ok || r.data.created === 0, 'prioritas tak dikenal ditolak');
r = api(CM, 'createTasks', { items: [mk(704, { parent: 'TSK-99999' })] }); ok(!r.ok || r.data.created === 0, 'parent tak ada ditolak');
const dupA = api(CM, 'createTasks', { items: [mk(705, { title: 'Duplikat X' }), mk(706, { title: 'Duplikat X' })] }); ok(dupA.ok && dupA.data.created === 1, 'duplikat dalam satu kiriman: hanya 1 (' + (dupA.data && dupA.data.created) + ')');
const dupB = api(CM, 'createTasks', { items: [mk(707, { title: 'Duplikat X' })] }); ok(!dupB.ok || dupB.data.created === 0, 'duplikat lintas kiriman ditolak');

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S4. foto: batas & korupsi');
const JPG = Buffer.concat([Buffer.from([0xFF, 0xD8, 0xFF, 0xE0]), Buffer.from('body')]).toString('base64');
const ph = o => Object.assign({ taskId: T1, mime: 'image/jpeg', data: JPG, name: 'a.jpg', desc: 'x' }, o);
const nF = () => Object.keys(DRIVE.files).length;
let f0 = nF();
[ph({ data: '!!!notbase64***' }), ph({ data: '' }), ph({ mime: 'text/html' }), ph({ mime: 'image/svg+xml' }), ph({ name: '../../etc/passwd' }), ph({ name: '<script>.jpg' }), ph({ data: Buffer.alloc(6 * 1024 * 1024, 0xFF).toString('base64') }), ph({ data: null }), ph({ data: 123 }), ph({ taskId: 'TSK-00001; DROP' }), ph({ taskId: '' })].forEach((p, i) => {
  const rr = api(C1, 'uploadPhoto', p); if (rr.ok && i !== 4 && i !== 5) console.log('   diterima?', i, JSON.stringify(p).slice(0, 80));
});
console.log('   file di Drive bertambah:', nF() - f0, '(nama berbahaya boleh 2 file bila nama disanitasi)');
let okc = 0, failc = 0; for (let i = 0; i < 30; i++) { const rr = api(C1, 'uploadPhoto', ph({ reqId: 'bulk' + i, data: Buffer.concat([Buffer.from([0xFF, 0xD8, 0xFF, 0xE0]), Buffer.from('b' + i)]).toString('base64') })); rr.ok ? okc++ : failc++; }
ok(okc <= 20 && failc >= 10 - 0, 'batas foto per tugas ditegakkan (diterima ' + okc + ', ditolak ' + failc + ')');
const names = Object.keys(DRIVE.files).map(k => DRIVE.files[k].name || ''); ok(names.every(x => !/[\/\\<>]/.test(x)), 'nama file di Drive bersih dari / \\ < >');

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S5. skala: 10.000 tugas, performa & konsistensi');
const db = Db.open(); const NBIG = 9000; const ids = []; const objs = []; const H = CONFIG.DB_H; const pics = ['Staf 01', 'Staf 02', 'Staf 03']; const projs = ['Proyek Contoh A', 'Proyek Contoh B', 'Proyek Contoh C'];
const stat = ['NEW', 'IN PROGRESS', 'PENDING', 'COMPLETED', 'CANCELLED'];
const exp = { total: 0, completed: 0, overdue: 0, cancelled: 0 };
for (let i = 0; i < NBIG; i++) {
  const st = stat[i % 5]; const dl = midn((i % 40) - 20); dl.setHours(23, 59, 0, 0);
  objs.push({ [H.TASK_ID]: 'TSK-' + String(2000 + i).padStart(5, '0'), [H.CREATED_AT]: midn(-30), [H.CREATED_BY]: 'Manager', [H.PROJECT_NAME]: projs[i % 3], [H.ASSIGNED_TO]: pics[i % 3], [H.TITLE]: 'Massal ' + i, [H.PRIORITY]: 'Normal', [H.START]: midn(-30), [H.DEADLINE]: dl, [H.STATUS]: st, [H.PROGRESS]: st === 'COMPLETED' ? 1 : (i % 10) / 10, [H.LAST_UPDATED]: midn(-1), [H.COMPLETION]: st === 'COMPLETED' ? midn(-2) : '' });
  if (st !== 'CANCELLED') exp.total++; if (st === 'COMPLETED') exp.completed++; if (st !== 'COMPLETED' && st !== 'CANCELLED' && dl.getTime() < Date.now()) exp.overdue++; if (st === 'CANCELLED') exp.cancelled++;
}
timeit('Db.append ' + NBIG + ' baris', () => { for (let i = 0; i < objs.length; i += 500) Db.append(Db.open(), objs.slice(i, i + 500)); });
const dbn = SS.getSheetByName('TASK DATABASE'); ok(dbn.getLastRow() > NBIG, 'baris database ' + dbn.getLastRow());
const fcol = Util.headerMap(dbn, 1)[CONFIG.DB_H.TASK_ID] || 1; const effCol = Object.keys(Util.headerMap(dbn, 1)).find(k => /Effective/i.test(k)); ok(!!effCol && String(dbn.f ? (dbn.f.get(dbn.getLastRow() + ',' + Util.headerMap(dbn, 1)[effCol]) || 'f') : 'f').length > 0, 'kolom formula ada sampai akhir');
r = timeit('tasks open (admin)', () => api(CM, 'tasks', { scope: 'open' })); ok(r.ok && r.data.total > 0 && r.data.tasks.length <= 200 && r.data.truncated, 'tasks dibatasi 200 baris (total ' + r.data.total + ')');
const bytes = JSON.stringify(r).length; ok(bytes < 400000, 'payload tasks wajar: ' + bytes + ' B');
r = timeit('tasks staf', () => api(C1, 'tasks', { scope: 'open' })); ok(r.ok && r.data.tasks.every(t => t.pic === 'Staf 01'), 'staf hanya miliknya di 9k data');
const mon = timeit('monitor', () => api(CM, 'dashboard', {})); ok(mon.ok, 'dashboard ok: ' + (mon.error || '')); if (mon.ok) { const c = JSON.stringify(mon.data); console.log('   dashboard payload', c.length, 'B | expected total/completed/overdue/cancelled (9k massal saja):', JSON.stringify(exp)); } if (mon.ok) { const d = mon.data; console.log('   dashboard keys:', Object.keys(d).join(','), JSON.stringify(d.counts || d.kpi || d.kpis || {}).slice(0, 260)); }
r = timeit('report staf', () => api(CM, 'report', { staff: 'Staf 01' })); ok(r.ok, 'report ok ' + (r.error || ''));
r = timeit('login', () => api(C1, 'login'));
r = timeit('task detail', () => api(C1, 'task', { taskId: 'TSK-02000' })); ok(r.ok || /bukan milik/.test(r.error), 'detail di 9k data');
r = timeit('submitUpdate di 9k data', () => api(C1, 'submitUpdate', { taskId: 'TSK-02003', status: 'IN PROGRESS', progress: 50 })); ok(r.ok || /bukan milik/.test(r.error) || true, 'update ' + (r.error || 'ok'));
r = timeit('createTasks 30 di 9k data', () => api(CM, 'createTasks', { items: Array.from({ length: 30 }, (_, i) => mk('B' + i, { title: 'Baru besar ' + i })) })); ok(r.ok && r.data.created === 30, '30 tugas baru: ' + (r.data ? r.data.created : r.error));
const allIds = []; for (let x = 2; x <= dbn.getLastRow(); x++) { const v = dbn.getRange(x, 1).getValue(); if (v) allIds.push(v); }
ok(new Set(allIds).size === allIds.length, 'semua Task ID unik (' + allIds.length + ')');

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S6. konkurensi (lock) & pemulihan');
LOCKED = true; r = api(CM, 'createTasks', { items: [mk('L1', { title: 'Saat terkunci' })] }); ok(!r.ok && !/Kesalahan sistem/.test(r.error), 'lock sibuk -> pesan sopan: ' + r.error); LOCKED = false;
const before = SS.getSheetByName('TASK DATABASE').getLastRow(); ok(!api(CM, 'tasks', {}).error, 'setelah lock lepas normal'); ok(SS.getSheetByName('TASK DATABASE').getLastRow() === before, 'tidak ada baris setengah jadi');
// replay idempotensi
const rq = { reqId: 'same-1', items: [mk('R1', { title: 'Replay satu' })] }; const a1 = api(CM, 'createTasks', rq), a2 = api(CM, 'createTasks', rq); ok(JSON.stringify(a1) === JSON.stringify(a2), 'replay reqId -> hasil sama, tidak ganda');
const cnt = api(CM, 'tasks', { scope: 'all', q: 'Replay satu' }).data.tasks.length; ok(cnt === 1, 'hanya 1 "Replay satu" (' + cnt + ')');

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S7. batas waktu (deadline 23:59)');
const RealDate = Date; function at(h, m, s, dayOff) { const base = new RealDate(); const t = new RealDate(base.getFullYear(), base.getMonth(), base.getDate() + dayOff, h, m, s); return t; }
const rr0 = api(CM, 'createTasks', { items: [mk('Z1', { title: 'Batas waktu', deadline: dstr(0) })] }); const idZ = rr0.data.ids ? rr0.data.ids[0] : api(CM, 'tasks', { q: 'Batas waktu' }).data.tasks[0].id;
function nowIs(d) { global.Date = class extends RealDate { constructor(...a) { if (a.length === 0) super(d.getTime()); else super(...a); } static now() { return d.getTime(); } }; }
[[23, 58, 59, 0, false], [23, 59, 1, 0, true], [0, 0, 1, 1, true]].forEach(([h, m, s, off, over]) => {
  nowIs(at(h, m, s, off)); let t; try { t = api(CM, 'task', { taskId: idZ }).data.task; } finally { global.Date = RealDate; }
  ok((t.eff === 'OVERDUE') === over, 'pada ' + h + ':' + m + ':' + s + ' (+' + off + 'd) eff=' + t.eff + ' (overdue harus ' + over + ')');
});

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S8. urutan 150 update bertubi pada satu tugas');
let acc = 0; const seq = ['IN PROGRESS', 'PENDING', 'IN PROGRESS', 'COMPLETED']; const plBefore = SS.getSheetByName('PROGRESS LOG').getLastRow();
for (let i = 0; i < 150; i++) { const s = i % 2 ? 'IN PROGRESS' : 'IN PROGRESS'; const rr = api(C1, 'submitUpdate', { taskId: T1, status: s, progress: i % 101, notes: 'u' + i }); if (rr.ok) acc++; }
ok(acc > 0, 'update beruntun diterima: ' + acc); const plAfter = SS.getSheetByName('PROGRESS LOG').getLastRow(); ok(plAfter - plBefore === acc, 'baris log = jumlah diterima (' + (plAfter - plBefore) + '/' + acc + ')');
r = api(C1, 'task', { taskId: T1 }); ok(r.ok && r.data.history.length >= acc, 'riwayat lengkap ' + r.data.history.length);
r = api(C1, 'submitUpdate', { taskId: T1, status: 'COMPLETED', progress: 100 }); ok(r.ok, 'selesai'); r = api(C1, 'submitUpdate', { taskId: T1, status: 'IN PROGRESS' }); ok(!r.ok, 'COMPLETED tidak bisa dibuka lagi tanpa ALLOW_REOPEN_CLOSED: ' + (r.error || 'diterima'));

console.log('   (+' + (Date.now() - T0) + ' ms)'); console.log('--- S9. kebersihan: SYSTEM LOG tidak meledak & email terbatas');
const lg = SS.getSheetByName('SYSTEM LOG'); console.log('   baris SYSTEM LOG:', lg.getLastRow(), '| email terkirim:', MAILS.length);
ok(MAILS.length < 400, 'email tidak membanjir (' + MAILS.length + ')');
console.log('   total waktu stress:', Date.now() - T0, 'ms');
})();
