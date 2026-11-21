// ============ TESTS TURN 4: patch, web app, foto ============
(function () {
let pt, g, res, thrown, tin;
function resetWorld(v1) {
  SS = new Spreadsheet();
  Object.keys(PROPS).forEach(k => delete PROPS[k]); Object.keys(CACHE).forEach(k => delete CACHE[k]);
  MAILS.length = 0; ALERTS.length = 0; DRIVE.files = {}; DRIVE.folders = {}; DRIVE.copies = []; DRIVE.denyCopy = false; DRIVE.denySharing = false;
  ACTIVE_EMAIL = 'manager@example.com'; LOCKED = false;
  if (v1) downgradeToV1();
}
function downgradeToV1() {
  const ti = SS.getSheetByName('TASK INPUT');
  ti.insertColumnBefore(13);
  ti.getRange(4, 13).setValue('Deadline Time'); ti.getRange(3, 13).setValue('17:00');
  ti.getRange(5, 15).setValue('(status lama)');
  ti.f.set('16,16', '=BROKEN'); ti.f.set('5,16', '=IF(COUNTA($C5:$N5)=0,"","lama #REF!")');
  const se = SS.getSheetByName('SETTINGS');
  for (let r = 20; r >= 10; r--) for (let c = 13; c <= 15; c++) { const v = se.cells.get(r + ',' + c); if (v === undefined) se.cells.delete((r + 1) + ',' + c); else se.cells.set((r + 1) + ',' + c, v); }
  se.getRange(10, 13).setValue('DEFAULT_DEADLINE_TIME'); se.getRange(10, 14).setValue('17:00'); se.getRange(10, 15).setValue('Jam default');
  const em = [...Array(30).keys()].map(i => i + 5).find(r => se.getRange(r, 13).getValue() === 'EMAIL_ENABLED');
  se.getRange(em, 14).setValue('NO');
  SS.sheets = SS.sheets.filter(s => s.name !== 'SYSTEM LOG');
}
const cfgRowOf = key => { const se = SS.getSheetByName('SETTINGS'); for (let r = 5; r <= 80; r++) if (se.getRange(r, 13).getValue() === key) return r; return 0; };
const cfgVal = key => SS.getSheetByName('SETTINGS').getRange(cfgRowOf(key), 14).getValue();
const setCfg = (key, v) => SS.getSheetByName('SETTINGS').getRange(cfgRowOf(key), 14).setValue(v);
const api = (code, action, payload) => JSON.parse(Api.handle(code, action, JSON.stringify(payload || {})));

console.log('--- Q. patch: layout lama (v1) -> terbaru');
resetWorld(true);
ok(Patch.pending().map(s => s.id).join() === 'input-deadline-time,system-log,settings-web,staff-code', 'semua langkah pending: ' + Patch.pending().map(s => s.id));
ok(Patch.isCurrent() === false, 'belum ditandai terbaru');
pt = Patch.run();
ok(!pt.failed && pt.applied.length === 4, 'patch v1 sukses: ' + JSON.stringify(pt.failed));
ok(DRIVE.copies.length === 1 && /^BACKUP sebelum patch .* - Staff Tracking$/.test(DRIVE.copies[0]), 'backup dibuat: ' + DRIVE.copies);
tin = SS.getSheetByName('TASK INPUT');
eq([tin.getRange(4, 12).getValue(), tin.getRange(4, 13).getValue(), tin.getRange(4, 14).getValue(), tin.getRange(4, 15).getValue()], ['Deadline Date', 'Notes', 'Submission Status', 'Validation Message'], 'header TASK INPUT sesuai layout baru');
const f5 = tin.getRange(5, 15).getFormula();
ok(/\$N5="SUBMITTED"/.test(f5) && !/REF/.test(f5) && /SETTINGS!\$N\$11<>"YES"/.test(f5) && /TEXTJOIN/.test(f5) && !/_xlfn/.test(f5), 'rumus Validation Message dibangun ulang, ALLOW_PAST_DEADLINE baris 11: ' + f5.slice(0, 160));
ok(tin.getRange(304, 15).getFormula().indexOf('$N304') > 0 && !tin.getRange(305, 15).getFormula(), 'rumus sampai baris 304');
ok(/COUNTA\(\$C5:\$M5\)/.test(tin.getRange(5, 1).getFormula()), 'rumus No');
ok(!SS.getSheetByName('TASK INPUT').f.get('16,16'), 'kolom ke-16 lama tidak tersisa');
ok(!!SS.getSheetByName('SYSTEM LOG') && SS.getSheetByName('SYSTEM LOG').getRange(1, 4).getValue() === 'Message', 'SYSTEM LOG dibuat');
ok(cfgVal('EMAIL_ENABLED') === 'YES', 'EMAIL_ENABLED NO -> YES pada layout lama');
ok(/tidak dipakai/.test(String(SS.getSheetByName('SETTINGS').getRange(cfgRowOf('DEFAULT_DEADLINE_TIME'), 15).getValue())), 'DEFAULT_DEADLINE_TIME ditandai tidak dipakai');
['WEB_ENABLED', 'WEB_APP_URL', 'EVIDENCE_PUBLIC_LINK', 'REQUIRE_EVIDENCE_ON_COMPLETE', 'MAX_PHOTOS_PER_TASK'].forEach(k => ok(cfgRowOf(k) >= 32, 'kunci ' + k + ' di baris ' + cfgRowOf(k)));
ok(SS.getSheetByName('SETTINGS').getRange(4, 12).getValue() === 'Admin Access Code' && SS.getSheetByName('STAFF MASTER').getRange(1, 6).getValue() === 'Access Code', 'header kode akses');
ok(PROPS.PATCH_VERSION === Patch.VERSION && Patch.isCurrent(), 'versi patch tercatat');
const lg = SS.getSheetByName('SYSTEM LOG'); const lgTxt = []; for (let r = 2; r <= lg.getLastRow(); r++) lgTxt.push(lg.getRange(r, 4).getValue());
ok(lgTxt.some(x => /Mulai patch/.test(x)) && lgTxt.filter(x => /->/.test(x)).length >= 4, 'patch tercatat di SYSTEM LOG');
// idempotent
pt = Patch.run(); ok(pt.already && !pt.applied.length && DRIVE.copies.length === 1, 'jalankan ulang: tidak ada perubahan & tidak ada backup baru');
// Submit tetap bekerja setelah migrasi
IN().getRange(5, C.CREATED_BY).setValue('Manager'); IN().getRange(5, C.PROJECT).setValue('Proyek Contoh A'); IN().getRange(5, C.PIC).setValue('Staf 01');
IN().getRange(5, C.TITLE).setValue('Setelah patch'); IN().getRange(5, C.PRIORITY).setValue('High'); IN().getRange(5, C.DEADLINE_DATE).setValue(midnight(5));
r = Submit.run(); eq(r.submitted.map(s => s.id), ['TSK-00001'], 'submit setelah migrasi'); eq(IN().getRange(5, C.STATUS).getValue(), 'SUBMITTED', 'status SUBMITTED kolom N');

console.log('--- R. patch: layout v2 (tanpa Deadline Time) & kegagalan');
resetWorld(false);
ok(Patch.pending().map(s => s.id).join() === 'settings-web,staff-code', 'v2 hanya butuh 2 langkah: ' + Patch.pending().map(s => s.id));
setCfg('EMAIL_ENABLED', 'NO');
pt = Patch.run(); ok(!pt.failed && pt.applied.length === 2, 'patch v2 sukses');
ok(cfgVal('EMAIL_ENABLED') === 'NO', 'pilihan EMAIL_ENABLED=NO milik pengguna TIDAK diubah pada v2');
resetWorld(true); DRIVE.denyCopy = true;
thrown = ''; try { Patch.run(); } catch (e) { thrown = e.message; }
ok(/Drive denied/.test(thrown), 'backup gagal -> patch dibatalkan');
ok(Patch.deadlineTimeCol_() === 13 && !SS.getSheetByName('SYSTEM LOG'), 'tidak ada perubahan setelah backup gagal');
DRIVE.denyCopy = false; ok(LOCKED === false, 'lock dilepas setelah error');
resetWorld(true); SS.getSheetByName('TASK INPUT').getRange(4, 12).setValue('Tanggal Aneh');
pt = Patch.run(); ok(pt.failed && pt.failed.id === 'input-deadline-time' && /tidak dikenali/.test(pt.failed.error), 'layout aneh ditolak: ' + JSON.stringify(pt.failed));
ok(pt.applied.length === 0 && Patch.deadlineTimeCol_() === 13 && !Patch.isCurrent(), 'berhenti tanpa mengubah & tidak ditandai terbaru');
resetWorld(false); SS.getSheetByName('STAFF MASTER').getRange(1, 6).setValue('Catatan');
pt = Patch.run(); ok(pt.failed && pt.failed.id === 'staff-code' && /Kosongkan/.test(pt.failed.error), 'F1 terisi -> tidak ditimpa');
ok(SS.getSheetByName('STAFF MASTER').getRange(1, 6).getValue() === 'Catatan', 'isi F1 utuh');

console.log('--- S. kode akses & login');
resetWorld(false); Patch.run();
g = Auth.generate();
const stf = Masters.staff(Settings.load()).list;
ok(g.made === stf.length + 2, 'kode dibuat untuk semua staf + 2 admin: ' + g.made);
const S0 = Settings.load();
const codeOf = n => { const s = Masters.staff(S0).list.find(x => x.name === n); return s ? s.code : S0.admins.find(a => a.name === n).code; };
const CM = codeOf('Manager'), CA = codeOf('Atasan'), C1 = codeOf('Staf 01'), C2 = codeOf('Staf 02'), C3 = codeOf('Staf 03');
const allCodes = [...Masters.staff(S0).list.map(x => x.code), ...S0.admins.map(a => a.code)];
ok(allCodes.every(c => /^[A-HJKMNP-Z2-9]{4}-[A-HJKMNP-Z2-9]{4}$/.test(c)) && new Set(allCodes).size === allCodes.length, 'format & unik: ' + allCodes.slice(0, 3));
const before = JSON.stringify(allCodes); g = Auth.generate(); ok(g.made === 0 && JSON.stringify(Masters.staff(Settings.load()).list.map(x => x.code).concat(Settings.load().admins.map(a => a.code))) === before, 'generate ulang tidak mengubah kode');
res = api('SALAH-1234', 'login'); ok(!res.ok && /salah/.test(res.error), 'kode salah ditolak');
res = api('', 'login'); ok(!res.ok, 'kode kosong ditolak');
res = api(CM.toLowerCase().replace('-', ' '), 'login'); ok(res.ok && res.data.user.role === 'admin' && res.data.user.name === 'Manager', 'login admin (huruf kecil / spasi toleran)');
ok(res.data.projects.length >= 1 && res.data.staff.length >= 1 && res.data.priorities.length === 4, 'meta admin lengkap');
res = api(C1, 'login'); ok(res.ok && res.data.user.role === 'staff' && !res.data.projects && !res.data.staff, 'login staf; meta staf TIDAK memuat daftar staf/proyek');
// staf nonaktif
SS.getSheetByName('STAFF MASTER').getRange(7, 5).setValue('Inactive'); const C6 = codeOf('Staf 06');
res = api(C6, 'login'); ok(!res.ok, 'staf Inactive tidak bisa login');
SS.getSheetByName('STAFF MASTER').getRange(7, 5).setValue('Active');
// kode ganda
SS.getSheetByName('STAFF MASTER').getRange(3, 6).setValue(C1); res = api(C1, 'login'); ok(!res.ok && /ganda/.test(res.error), 'kode ganda ditolak');
SS.getSheetByName('STAFF MASTER').getRange(3, 6).setValue(C2);
// aksi tidak dikenal / prototype
res = api(CM, 'constructor'); ok(!res.ok, 'aksi tidak dikenal'); res = api(CM, 'toString'); ok(!res.ok, 'aksi prototype ditolak');
// throttle
Object.keys(CACHE).forEach(k => delete CACHE[k]);
for (let i = 0; i < 20; i++) api('XXXX-' + (1000 + i), 'login');
Utilities._slept = 0; api('XXXX-9999', 'login'); ok(Utilities._slept >= 2000, 'tebakan salah diperlambat setelah 20 kegagalan');
res = api(CM, 'login'); ok(res.ok, 'kode benar tetap lolos saat serangan (tidak ada DoS lock-out)');
Object.keys(CACHE).forEach(k => delete CACHE[k]); res = api(CM, 'login'); ok(res.ok, 'pulih setelah jendela habis');
setCfg('WEB_ENABLED', 'NO'); res = api(CM, 'login'); ok(!res.ok && /dinonaktifkan/.test(res.error), 'WEB_ENABLED=NO menutup web'); setCfg('WEB_ENABLED', 'YES');

console.log('--- T. admin: input tugas via web');
MAILS.length = 0;
const dstr = off => Utilities.formatDate(midnight(off), 'Asia/Jakarta', 'yyyy-MM-dd');
res = api(C1, 'createTasks', { items: [{ project: 'Proyek Contoh A', pic: 'Staf 01', title: 'x', priority: 'High', deadline: dstr(3) }] });
ok(!res.ok && /admin/.test(res.error), 'staf tidak boleh membuat tugas');
res = api(CM, 'createTasks', { items: [] }); ok(!res.ok, 'daftar kosong ditolak');
const items = [
  { project: 'Proyek Contoh A', pic: 'Staf 01', title: 'Web tugas 1', desc: 'deskripsi', priority: 'High', start: dstr(0), deadline: dstr(5), notes: 'catatan' },
  { project: 'Proyek Contoh B', pic: 'Staf 02', title: 'Web tugas 2', priority: 'Normal', deadline: dstr(2) },
  { project: 'Proyek Contoh A', pic: 'Staf 01', title: 'Web tugas 1', priority: 'High', deadline: dstr(5) },     // duplikat dalam batch
  { project: 'Proyek X', pic: 'Staf 01', title: 'Salah proyek', priority: 'High', deadline: dstr(5) },
  { project: 'Proyek Contoh A', pic: 'Staf 03', title: 'Tanggal salah', priority: 'High', deadline: 'besok' }
];
res = api(CM, 'createTasks', { reqId: 'req-1', items });
ok(res.ok && res.data.created === 2, 'dibuat 2 dari 5: ' + JSON.stringify(res.data && res.data.results));
eq(res.data.results.map(x => x.ok), [true, true, false, false, false], 'hasil per baris');
ok(res.data.results[0].id === 'TSK-00001' && res.data.results[1].id === 'TSK-00002', 'ID berurutan');
ok(/Duplikat/.test(res.data.results[2].errors.join()) && /Proyek X/.test(res.data.results[3].errors.join()) && /Format Deadline/.test(res.data.results[4].errors.join()), 'pesan error per baris');
ok(MAILS.filter(m => m.to === 'staf01@example.com' || m.to === 'staf02@example.com').length >= 2, 'email penugasan terkirim ke PIC');
const DB2 = SS.getSheetByName('TASK DATABASE'); const dh = Util.headerMap(DB2, 1);
eq([DB2.getRange(2, dh['Created By']).getValue(), DB2.getRange(2, dh['Submitted By (Email)']).getValue(), DB2.getRange(2, dh['Current Status']).getValue()], ['Manager', 'manager@example.com', 'NEW'], 'Created By dari kode akses (bukan input)');
ok(DB2.getRange(2, dh['Deadline']).getValue().getHours() === 23, 'deadline = 23:59');
ok(/Catatan: catatan/.test(DB2.getRange(2, dh['Task Description']).getValue()), 'catatan digabung ke deskripsi');
// idempotensi (koneksi putus lalu kirim ulang dengan reqId sama)
const n0 = Util.lastNonEmptyRow(DB2, 1, 2); res = api(CM, 'createTasks', { reqId: 'req-1', items });
ok(res.ok && res.data.created === 2 && Util.lastNonEmptyRow(DB2, 1, 2) === n0, 'kirim ulang dengan reqId sama -> hasil sama, tidak ada baris baru');
res = api(CM, 'createTasks', { items: [items[0]] }); ok(res.ok && res.data.created === 0 && /Duplikat/.test(res.data.results[0].errors.join()), 'tanpa reqId: dicegah oleh deteksi duplikat');
res = api(CA, 'createTasks', { items: [{ project: 'Proyek Contoh B', pic: 'Staf 03', title: 'Dari Atasan', priority: 'Low', deadline: dstr(9) }] });
ok(res.ok && DB2.getRange(4, dh['Created By']).getValue() === 'Atasan', 'Created By = Atasan untuk kode Atasan');

console.log('--- U. staf: lihat & laporkan');
res = api(C1, 'tasks', { scope: 'open' }); ok(res.ok && res.data.tasks.length === 1 && res.data.tasks[0].pic === 'Staf 01', 'staf 01 hanya melihat tugasnya');
res = api(C1, 'tasks', { scope: 'open', pic: 'Staf 02' }); ok(res.data.tasks.every(t => t.pic === 'Staf 01'), 'staf tidak bisa memfilter staf lain');
res = api(CM, 'tasks', { scope: 'open' }); ok(res.data.total === 3, 'admin melihat semua: ' + res.data.total);
res = api(CM, 'tasks', { scope: 'open', pic: 'Staf 02' }); ok(res.data.tasks.length === 1, 'admin filter PIC');
res = api(CM, 'tasks', { scope: 'open', q: 'dari atasan' }); ok(res.data.tasks.length === 1 && res.data.tasks[0].id === 'TSK-00003', 'pencarian teks');
res = api(C1, 'task', { taskId: 'TSK-00002' }); ok(!res.ok && /bukan milik/.test(res.error), 'staf tidak bisa membuka tugas orang lain');
res = api(C1, 'task', { taskId: 'tsk-00001' }); ok(res.ok && res.data.task.title === 'Web tugas 1' && Array.isArray(res.data.history) && res.data.evidence.length === 0, 'detail tugas sendiri (ID tidak peka huruf)');
// update
MAILS.length = 0;
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'IN PROGRESS', progress: 40, notes: 'mulai', reqId: 'u1' });
ok(res.ok && res.data.task.status === 'IN PROGRESS' && res.data.task.progress === 0.4, 'update status & progres');
const PL2 = SS.getSheetByName('PROGRESS LOG'); const lastPL = PL2.getLastRow();
eq([PL2.getRange(lastPL, 3).getValue(), PL2.getRange(lastPL, 4).getValue(), PL2.getRange(lastPL, 5).getValue()], ['TSK-00001', 'Staf 01', 'NEW'], 'baris PROGRESS LOG tercatat (Updated By = nama staf dari kode)');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'IN PROGRESS', progress: 40, notes: 'mulai', reqId: 'u1' }); ok(res.ok && PL2.getLastRow() === lastPL, 'reqId sama: tidak ada log ganda');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00002', status: 'IN PROGRESS' }); ok(!res.ok && /bukan milik/.test(res.error), 'update tugas orang lain ditolak');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'CANCELLED' }); ok(!res.ok && /admin/.test(res.error), 'staf tidak boleh CANCELLED');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'PENDING' }); ok(!res.ok && /Delay Reason/.test(res.error), 'PENDING tanpa alasan ditolak');
ok(PL2.getRange(PL2.getLastRow(), 3).getValue() === 'TSK-00001' && PL2.getLastRow() === lastPL, 'baris ditolak tidak tertinggal di log');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'IN PROGRESS', progress: 150 }); ok(!res.ok && /0-100/.test(res.error), 'progress > 100 ditolak');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'PENDING', delay: 'tunggu material', revised: dstr(8) }); ok(res.ok && res.data.task.status === 'PENDING', 'PENDING dengan alasan');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'NEW' }); ok(!res.ok, 'tidak bisa kembali ke NEW');
ok(MAILS.some(m => m.to === 'manager@example.com' && /PENDING/.test(m.subject)), 'admin diberi email saat PENDING');

console.log('--- V. foto bukti');
const JPG = Buffer.concat([Buffer.from([0xFF, 0xD8, 0xFF, 0xE0]), Buffer.from('fake-jpeg-body')]).toString('base64');
const PNG = Buffer.concat([Buffer.from([0x89, 0x50, 0x4E, 0x47]), Buffer.from('xx')]).toString('base64');
const photo = (o) => Object.assign({ taskId: 'TSK-00001', mime: 'image/jpeg', data: JPG, name: 'a.jpg', desc: 'hasil pekerjaan' }, o);
res = api(C1, 'uploadPhoto', photo({ reqId: 'p1' }));
ok(res.ok && /^EVD-/.test(res.data.evidenceId) && /drive\.google\.com\/file/.test(res.data.url), 'upload foto sukses: ' + JSON.stringify(res.error || res.data));
const EV = SS.getSheetByName('EVIDENCE LOG'); const eh = Util.headerMap(EV, 1);
eq([EV.getRange(2, eh['Task ID']).getValue(), EV.getRange(2, eh['Uploaded By']).getValue(), EV.getRange(2, eh['Processing Note']).getValue()], ['TSK-00001', 'Staf 01', 'OK'], 'EVIDENCE LOG terisi & diproses');
ok(/^drive\.google\.com|^https:\/\/drive\.google\.com\/drive\/folders\//.test(DB2.getRange(2, dh['Evidence Link']).getValue()), 'Evidence Link database = folder tugas');
const rootF = DRIVE.folders[PROPS.EVIDENCE_FOLDER_ID]; ok(rootF && rootF.children.length === 1 && rootF.children[0].name === 'TSK-00001' && rootF.children[0].files.length === 1, 'folder induk > folder tugas > file');
ok(rootF.sharing === 'ANYONE_WITH_LINK:VIEW' && rootF.viewers.indexOf('manager@example.com') >= 0, 'folder dibagikan (link + email admin)');
ok(/^TSK-00001_\d{8}-\d{6}_1\.jpg$/.test(rootF.children[0].files[0].blob.name), 'nama file: ' + rootF.children[0].files[0].blob.name);
res = api(C1, 'uploadPhoto', photo({ reqId: 'p1' })); ok(res.ok && EV.getLastRow() === 2 && rootF.children[0].files.length === 1, 'reqId sama: tidak ada foto ganda');
res = api(C1, 'uploadPhoto', photo({ data: PNG })); ok(!res.ok && /bukan foto/.test(res.error), 'isi tidak cocok dengan mime ditolak');
res = api(C1, 'uploadPhoto', photo({ mime: 'application/pdf' })); ok(!res.ok && /tidak didukung/.test(res.error), 'PDF ditolak');
res = api(C1, 'uploadPhoto', photo({ data: Buffer.alloc(7 * 1024 * 1024, 1).toString('base64') })); ok(!res.ok && /terlalu besar/.test(res.error), 'terlalu besar ditolak');
res = api(C1, 'uploadPhoto', photo({ taskId: 'TSK-00002' })); ok(!res.ok && /bukan milik/.test(res.error), 'foto ke tugas orang lain ditolak');
res = api(C1, 'uploadPhoto', photo({ taskId: 'TSK-99999' })); ok(!res.ok && /tidak ditemukan/.test(res.error), 'tugas tidak ada');
res = api(CM, 'uploadPhoto', photo({ taskId: 'TSK-00002', desc: 'dari admin' })); ok(res.ok, 'admin boleh menambah foto ke tugas mana pun');
setCfg('MAX_PHOTOS_PER_TASK', 2); res = api(C1, 'uploadPhoto', photo({})); ok(res.ok, 'foto ke-2'); res = api(C1, 'uploadPhoto', photo({})); ok(!res.ok && /Batas foto/.test(res.error), 'batas foto per tugas'); setCfg('MAX_PHOTOS_PER_TASK', 20);
res = api(C1, 'task', { taskId: 'TSK-00001' }); ok(res.data.evidence.length === 2 && res.data.task.evidence === 2 && res.data.evidence[0].by === 'Staf 01', 'detail memuat foto');
const evId = res.data.evidence[0].id;
res = api(C1, 'evidenceImage', { evidenceId: evId }); ok(res.ok && res.data.mime === 'image/jpeg' && Buffer.from(res.data.data, 'base64').slice(0, 3).equals(Buffer.from([0xFF, 0xD8, 0xFF])), 'ambil gambar (base64)');
const evOther = api(CM, 'task', { taskId: 'TSK-00002' }).data.evidence[0].id; res = api(C1, 'evidenceImage', { evidenceId: evOther }); ok(!res.ok && /bukan milik/.test(res.error), 'staf tidak bisa mengambil foto tugas lain');
res = api(C1, 'evidenceImage', { evidenceId: 'EVD-99999' }); ok(!res.ok, 'foto tidak ada');
// kebijakan Drive menolak link publik -> tetap jalan, admin diberi akses
DRIVE.denySharing = true; res = api(C1, 'uploadPhoto', photo({ taskId: 'TSK-00001', reqId: 'p9' })); setCfg('MAX_PHOTOS_PER_TASK', 20);
ok(!res.ok || res.ok, 'sharing ditolak tidak membuat crash'); DRIVE.denySharing = false;
// wajib bukti saat selesai
setCfg('REQUIRE_EVIDENCE_ON_COMPLETE', 'YES');
res = api(C1, 'submitUpdate', { taskId: 'TSK-00001', status: 'COMPLETED', notes: 'selesai' }); ok(res.ok && res.data.task.status === 'COMPLETED', 'COMPLETED dengan foto yang sudah ada diterima');
const t2 = 'TSK-00003';
res = api(C3, 'submitUpdate', { taskId: t2, status: 'COMPLETED' }); ok(!res.ok && /foto bukti/.test(res.error), 'COMPLETED tanpa foto ditolak saat REQUIRE_EVIDENCE=YES');
res = api(C3, 'uploadPhoto', photo({ taskId: t2 })); ok(res.ok, 'staf 03 unggah foto');
res = api(C3, 'submitUpdate', { taskId: t2, status: 'COMPLETED', notes: 'beres' }); ok(res.ok && res.data.task.status === 'COMPLETED' && res.data.task.progress === 1, 'COMPLETED dengan foto: progres 100%');
setCfg('REQUIRE_EVIDENCE_ON_COMPLETE', 'NO');


console.log('--- W. admin: penilaian, dashboard, laporan');
res = api(C2, 'rateTask', { taskId: t2, quality: '5 - Sangat Baik' }); ok(!res.ok && /admin/.test(res.error), 'staf tidak boleh menilai');
res = api(CM, 'rateTask', { taskId: t2, quality: 'Bintang 9' }); ok(!res.ok, 'rating tidak valid');
res = api(CM, 'rateTask', { taskId: 'TSK-00002', quality: '4 - Baik' }); ok(!res.ok && /COMPLETED/.test(res.error), 'rating hanya untuk COMPLETED');
res = api(CM, 'rateTask', { taskId: t2, quality: '4 - Baik', delayCategory: 'External Party' }); ok(res.ok, 'admin menilai');
const rowT2 = Db.readAll(Db.open()).byId[t2]; eq([rowT2['Quality Rating'], rowT2['Delay Category']], ['4 - Baik', 'External Party'], 'nilai tersimpan di TASK DATABASE');
const lg2 = SS.getSheetByName('SYSTEM LOG'); const lgs = []; for (let i = 2; i <= lg2.getLastRow(); i++) lgs.push(lg2.getRange(i, 2).getValue()); ok(lgs.indexOf('WEB-RATE') >= 0 && lgs.indexOf('WEB-PHOTO') >= 0 && lgs.indexOf('WEB-UPDATE') >= 0 && lgs.indexOf('WEB-TASK') >= 0, 'aksi web tercatat di SYSTEM LOG: ' + [...new Set(lgs)]);
res = api(C1, 'dashboard', {}); ok(!res.ok && /admin/.test(res.error), 'staf tidak boleh membuka dashboard');
res = api(CM, 'dashboard', {}); ok(res.ok && res.data.kpi.total === 3 && res.data.kpi.completed === 2 && res.data.staff.length === 6 && res.data.projects.length === 3, 'dashboard: KPI, staf, proyek');
res = api(CM, 'dashboard', { from: dstr(100), to: dstr(101) }); ok(res.ok && res.data.kpi.total === 0, 'filter periode kosong');
res = api(C1, 'report', {}); ok(res.ok && res.data.name === 'Staf 01' && res.data.stats.total === 1, 'laporan staf = dirinya sendiri');
res = api(C1, 'report', { staff: 'Staf 02' }); ok(res.ok && res.data.name === 'Staf 01', 'staf tidak bisa melihat laporan orang lain (parameter diabaikan)');
res = api(CM, 'report', { staff: 'Staf 03' }); ok(res.ok && res.data.name === 'Staf 03' && res.data.stats.completed === 1, 'admin melihat laporan staf mana pun');
res = api(CM, 'report', {}); ok(!res.ok, 'admin wajib memilih staf');
ok(SelfTest.run().filter(x => !x.ok).map(x => x.name).indexOf('Patch terbaru (3.1.0)') < 0, 'selfTest: patch terbaru lulus');
// penyuntikan
res = api(C1, 'tasks', { scope: 'open', q: '<script>' }); ok(res.ok && res.data.tasks.length === 0, 'input aneh aman');
res = api(CM, 'createTasks', { items: [{ project: 'Proyek Contoh A', pic: 'Staf 01', title: '=HYPERLINK("http://x")', priority: 'High', deadline: dstr(4) }] });
ok(res.ok, 'judul diawali "=" diterima'); const tt = Db.readAll(Db.open()).rows.find(o => /HYPERLINK/.test(o['Task Title']));
ok(tt && !DB2.f.get(tt._row + ',' + dh['Task Title']), 'judul berawalan "=" tersimpan sebagai TEKS (bukan rumus)');

console.log('--- V2. foto sebelum/sesudah + link email + deep link');
const CM2 = codeOf('Manager'), C2b = codeOf('Staf 02');
res = api(CM2, 'createTasks', { items: [{ project: 'Proyek Contoh A', pic: 'Staf 02', title: 'tugas foto', priority: 'High', deadline: dstr(5) }] }); ok(res.ok, 'buat tugas V2: ' + (res.error || ''));
const TV = res.data.results[0].id;
setCfg('REQUIRE_BEFORE_AFTER', 'YES');
res = api(C2b, 'submitUpdate', { taskId: TV, status: 'COMPLETED' }); ok(!res.ok && /SEBELUM/.test(res.error), 'BEFORE_AFTER=YES: COMPLETED tanpa foto ditolak: ' + res.error);
res = api(C2b, 'uploadPhoto', photo({ taskId: TV, kind: 'BEFORE', desc: 'kondisi awal' })); ok(res.ok && res.data.kind === 'BEFORE', 'upload foto SEBELUM');
res = api(C2b, 'submitUpdate', { taskId: TV, status: 'COMPLETED' }); ok(!res.ok && /SEBELUM/.test(res.error), 'hanya SEBELUM: masih ditolak');
res = api(C2b, 'uploadPhoto', photo({ taskId: TV, kind: 'AFTER', desc: '' })); ok(res.ok && res.data.kind === 'AFTER', 'upload foto SESUDAH');
res = api(C2b, 'uploadPhoto', photo({ taskId: TV, kind: 'HACK<script>' })); ok(res.ok && res.data.kind === 'OTHER', 'kind tidak dikenal -> LAINNYA');
res = api(C2b, 'task', { taskId: TV });
eq(res.data.evidence.map(e => e.kind).sort(), ['AFTER', 'BEFORE', 'OTHER'], 'detail memuat kind foto (SEBELUM, SESUDAH, LAINNYA)');
ok(res.data.evidence.some(e => e.kind === 'BEFORE' && e.desc === 'kondisi awal'), 'desc tanpa prefiks tag');
{ let f = false; const hh = Util.headerMap(EV, 1); for (let i = 2; i <= EV.getLastRow(); i++) if (/^\[BEFORE\] kondisi awal/.test(EV.getRange(i, hh['Evidence Description']).getValue() + '')) f = true; ok(f, 'EVIDENCE LOG menyimpan tag [BEFORE]'); }
ok(Object.keys(DRIVE.folders).some(k => (DRIVE.folders[k].children || []).some(f => (f.files || []).some(x => new RegExp('^' + TV + '_SEBELUM_').test(x.blob.name)))), 'nama file memuat SEBELUM');
res = api(C2b, 'submitUpdate', { taskId: TV, status: 'COMPLETED' }); ok(res.ok && res.data.task.status === 'COMPLETED', 'SEBELUM+SESUDAH ada: COMPLETED diterima');
ok(api(C2b, 'login', {}).data.requireBeforeAfter === true, 'meta memuat requireBeforeAfter');
setCfg('REQUIRE_BEFORE_AFTER', 'NO');

const Sx = Settings.load();
ok(Web.link(Sx, { tab: 'report', staff: 'Staf 01' }) === 'https://script.google.com/macros/s/TEST/exec?tab=report&staff=Staf%2001', 'Web.link: ' + Web.link(Sx, { tab: 'report', staff: 'Staf 01' }));
setCfg('WEB_APP_URL', 'https://evil.example.com/x'); ok(Web.link(Settings.load(), { tab: 'tasks' }).indexOf('https://script.google.com/') === 0, 'WEB_APP_URL non-Google diabaikan'); setCfg('WEB_APP_URL', '');
setCfg('WEB_ENABLED', 'NO'); ok(Web.link(Settings.load(), { tab: 'tasks' }) === '', 'WEB_ENABLED=NO: tanpa link'); setCfg('WEB_ENABLED', 'YES');
eq(Web.deep({ parameter: { tab: 'report', task: 'tsk-00002', staff: 'Staf 01' } }), { tab: 'report', task: 'TSK-00002', staff: 'Staf 01' }, 'deep: param valid');
eq(Web.deep({ parameter: { tab: '"><script>', task: '<img onerror=1>', staff: 'x'.repeat(200) } }), {}, 'deep: param jahat dibuang');
eq(Web.deep(undefined), {}, 'deep: tanpa e');
const tdoc = doGet({ parameter: { tab: 'report', task: '</script><script>alert(1)' } }); ok(true, 'doGet(e) tidak crash'); ok(!/</.test(doGet({ parameter: { tab: 'tasks', task: TV } }).tpl.deep), 'deep di template tanpa karakter "<"');
MAILS.length = 0;
const fakeT = { id: TV, title: 'Tugas <b>x</b>', project: 'P', priority: 'High', deadline: new Date(), createdBy: 'Manager', pic: 'Staf 01', picEmail: 'staf01@example.com', adminEmail: 'manager@example.com' };
Notify.onSubmit({ submitted: [fakeT] });
const ml = MAILS[MAILS.length - 1] || {};
const MH = ml.htmlBody || ml.html || ''; const MB = ml.body || '';
ok(MH.indexOf('exec?tab=tasks&amp;task=' + TV + '"') > 0 && /<a href="https:\/\/script\.google\.com\//.test(MH), 'email penugasan: link detail tugas (href ter-escape)');
ok(/tab=report&amp;staff=Staf%2001/.test(MH), 'email penugasan: link laporan staf');
ok(/tab=report&staff=Staf%2001/.test(MB) || /tab=report/.test(MB), 'versi teks email memuat URL');
ok(MH.indexOf('<b>x</b>') < 0, 'judul tugas tetap di-escape');
MAILS.length = 0; setCfg('WEB_ENABLED', 'NO'); Notify.onSubmit({ submitted: [fakeT] }); ok(MAILS.length === 1 && !/<a /.test(MAILS[0].htmlBody || MAILS[0].html || ''), 'WEB_ENABLED=NO: email tanpa link, tetap terkirim'); setCfg('WEB_ENABLED', 'YES');
MAILS.length = 0; Notify.onProgress({ processed: [{ taskId: TV, title: 'T', by: 'Staf 02', createdBy: 'Manager', prev: 'NEW', status: 'COMPLETED', progress: 1, delay: '', notes: '' }] });
const MP = (MAILS[0] || {}).htmlBody || (MAILS[0] || {}).html || ''; ok(new RegExp('task=' + TV).test(MP) && /staff=Staf%2002/.test(MP), 'email ke admin: link tugas & laporan staf');
MAILS.length = 0; const fd = { overdue: [Object.assign({ days: 2 }, fakeT)], dueSoon: [] }; Notify.digest(Settings.load(), fd, new Date());
ok(MAILS.length >= 1 && MAILS.every(m => /<a href="https:\/\/script\.google\.com/.test(m.htmlBody || m.html || '')), 'digest staf & admin memuat link');


console.log('--- X. analitik web == hitungan LibreOffice (workbook demo)');
{
  const EXP = JSON.parse(fs.readFileSync('/home/claude/build/demo_expected.json', 'utf8'));
  resetWorld(false); Patch.run();
  const dbs = SS.getSheetByName('TASK DATABASE'); const hh = Util.headerMap(dbs, 1);
  EXP.rows.forEach((o, i) => Object.keys(o).forEach(h => { let v = o[h]; if (v === null || v === undefined) return;
    if (v && v.d) { const m = /^(\d+)-(\d+)-(\d+)T(\d+):(\d+)/.exec(v.d); v = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]); }
    dbs.getRange(2 + i, hh[h]).setValue(v); }));
  const S = Settings.load(); const a = Analytics.build(S, new Date(), {});
  const near = (x, y, tol) => (x === null && (y === '' || y === null)) || (x !== null && y !== '' && y !== null && Math.abs(x - y) <= tol);
  let bad = 0;
  EXP.staffPerf.forEach(row => {
    const m = a.staff.find(x => x.name === row[1]); if (!m) { bad++; console.log('  tidak ada staf', row[1]); return; }
    if (row[3] === 0) return;
    const pairs = [['total', 3, 0], ['completed', 4, 0], ['pending', 5, 0], ['overdue', 6, 0], ['overdueStaff', 7, 0], ['onTime', 8, 0], ['late', 9, 0], ['lateExcused', 10, 0],
      ['completionRate', 11, 1e-9], ['onTimeRate', 12, 1e-9], ['onTimeAdj', 13, 1e-9], ['avgDays', 14, 0.051], ['avgQuality', 15, 1e-9], ['consistency', 16, 1e-9], ['score', 17, 0.11]];
    pairs.forEach(([k, i, tol]) => { if (!near(m[k], row[i], tol)) { bad++; console.log('  BEDA', row[1], k, 'web=', m[k], 'sheet=', row[i]); } });
  });
  ok(bad === 0, 'metrik staf sama dengan STAFF PERFORMANCE (6 staf x 15 kolom)');
  const k = a.kpi, ek = EXP.kpi;
  eq([k.total, k.newCount, k.inProgress, k.pending, k.completed, k.overdue], ek.slice(0, 6), 'KPI dashboard (total/new/inprogress/pending/completed/overdue)');
  ok(Math.abs(k.completionRate - ek[6]) < 1e-9 && Math.abs(k.onTimeRate - ek[7]) < 1e-9, 'completion rate & on-time rate dashboard');
  let pbad = 0;
  EXP.proj.forEach(row => { const p = a.projects.find(x => x.name === row[2]); if (!p) { pbad++; return; }
    const got = [p.total, p.completed, p.inProgress, p.pending, p.newCount, p.overdue, p.open]; const want = [row[5], row[6], row[7], row[8], row[9], row[10], row[11]];
    if (JSON.stringify(got) !== JSON.stringify(want)) { pbad++; console.log('  BEDA proyek', row[2], got, want); }
    if (!near(p.rate, row[12], 1e-9)) { pbad++; console.log('  BEDA rate', row[2]); } });
  ok(pbad === 0, 'PROJECT REPORT sama (3 proyek)');
  // filter periode
  const a2 = Analytics.build(S, new Date(), { from: Utilities.formatDate(new Date(Date.now() + 3 * 864e5), 'Asia/Jakarta', 'yyyy-MM-dd') });
  ok(a2.kpi.total < a.kpi.total && a2.kpi.total > 0, 'filter periode mengurangi cakupan: ' + a2.kpi.total + ' < ' + a.kpi.total);
}

})();
