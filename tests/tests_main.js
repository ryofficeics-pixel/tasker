// ============ TESTS ============
let PASS = 0, FAIL = 0;
function ok(c, m) { if (c) PASS++; else { FAIL++; console.log('  FAIL:', m); } }
function eq(a, b, m) { ok(JSON.stringify(a) === JSON.stringify(b), m + ' -> got ' + JSON.stringify(a) + ' expected ' + JSON.stringify(b)); }
const now0 = new Date();
const midnight = (off) => new Date(now0.getFullYear(), now0.getMonth(), now0.getDate() + off);
const IN = () => SS.getSheetByName('TASK INPUT');
function setIn(row, o) { // o: col->value
  Object.keys(o).forEach(c => IN().getRange(row, +c).setValue(o[c]));
}
const C = CONFIG.INPUT.COL;
const mk = (extra) => Object.assign({ [C.INPUT_DATE]: midnight(0), [C.CREATED_BY]: 'Manager', [C.PROJECT]: 'Proyek Contoh A', [C.PIC]: 'Staf 01',
  [C.TITLE]: 'T', [C.PRIORITY]: 'High', [C.START]: midnight(0), [C.DEADLINE_DATE]: midnight(7) }, extra);

console.log('--- A. submit batch');
// make Staf 06 inactive
SS.getSheetByName('STAFF MASTER').getRange(7, 5).setValue('Inactive');
setIn(5, mk({ [C.TITLE]: 'Valid satu' }));
setIn(6, { [C.CREATED_BY]: 'Manager' });                                    // missing
setIn(7, mk({ [C.TITLE]: 'Parent salah', [C.PARENT]: 'TSK-99999' }));
setIn(8, mk({ [C.TITLE]: 'Valid satu hari', [C.PIC]: 'Staf 02', [C.NOTES]: 'bawa laptop', [C.DESC]: 'deskripsi' }));
setIn(9, mk({ [C.TITLE]: 'Valid satu' }));                                  // duplicate of row 5
setIn(10, mk({ [C.TITLE]: 'Lewat', [C.DEADLINE_DATE]: midnight(-3), [C.START]: midnight(-5) }));
setIn(11, mk({ [C.TITLE]: 'Staf nonaktif', [C.PIC]: 'Staf 06' }));
setIn(12, mk({ [C.TITLE]: 'Admin palsu', [C.CREATED_BY]: 'Hacker' }));
setIn(13, mk({ [C.TITLE]: 'Tanggal teks', [C.DEADLINE_DATE]: 'besok' }));
setIn(14, mk({ [C.TITLE]: 'Valid lain', [C.PIC]: 'Staf 03', [C.PROJECT]: 'Proyek Contoh B', [C.CREATED_BY]: 'Atasan' }));
let r = Submit.run();
eq(r.submitted.map(s => s.id), ['TSK-00001', 'TSK-00002', 'TSK-00003'], 'IDs');
eq(r.submitted.map(s => s.row), [5, 8, 14], 'submitted rows');
eq(r.errors.map(e => e.row), [6, 7, 9, 10, 11, 12, 13], 'error rows');
const msgOf = row => r.errors.find(e => e.row === row).messages.join(' | ');
ok(/Project wajib/.test(msgOf(6)) && /PIC\) wajib/.test(msgOf(6)) && /Title wajib/.test(msgOf(6)) && /Priority wajib/.test(msgOf(6)) && /Deadline Date wajib/.test(msgOf(6)), 'row6 msgs: ' + msgOf(6));
ok(/Parent Task ID/.test(msgOf(7)), 'row7'); ok(/Duplikat/.test(msgOf(9)), 'row9 ' + msgOf(9)); ok(/Deadline sudah lewat/.test(msgOf(10)), 'row10');
ok(/Inactive/.test(msgOf(11)), 'row11'); ok(/bukan admin/.test(msgOf(12)), 'row12'); ok(/Format Deadline Date/.test(msgOf(13)), 'row13');
const DB = SS.getSheetByName('TASK DATABASE');
const dbh = Util.headerMap(DB, 1);
const dbv = (row, h) => DB.getRange(row, dbh[h]).getValue();
eq([2, 3, 4].map(x => dbv(x, 'Task ID')), ['TSK-00001', 'TSK-00002', 'TSK-00003'], 'DB ids rows 2..4');
eq(DB.getRange(5, 1).getValue(), '', 'row 5 DB empty');
eq(dbv(2, 'Project ID'), 'PRJ-001', 'project id'); eq(dbv(2, 'Current Status'), 'NEW', 'status NEW'); eq(dbv(2, 'Progress Percentage'), 0, 'progress 0');
eq(dbv(2, 'Submitted By (Email)'), 'manager@example.com', 'submitted by');
ok(dbv(2, 'Deadline').getHours() === 23 && dbv(2, 'Deadline').getMinutes() === 59, 'deadline = end of day 23:59: ' + dbv(2, 'Deadline'));
ok(Util.fmt(dbv(2, 'Deadline'), 'yyyy-MM-dd') === Util.fmt(midnight(7), 'yyyy-MM-dd'), 'deadline date kept');
ok(/Catatan: bawa laptop/.test(dbv(3, 'Task Description')) && /^deskripsi/.test(dbv(3, 'Task Description')), 'notes appended');
eq(dbv(4, 'Created By'), 'Atasan', 'created by Atasan'); eq(dbv(4, 'Assigned To'), 'Staf 03', 'pic');
eq([5, 8, 14].map(x => IN().getRange(x, C.STATUS).getValue()), ['SUBMITTED', 'SUBMITTED', 'SUBMITTED'], 'status SUBMITTED');
eq([5, 8, 14].map(x => IN().getRange(x, C.TASK_ID).getValue()), ['TSK-00001', 'TSK-00002', 'TSK-00003'], 'input task ids');
eq(IN().getRange(8, C.INPUT_DATE).getValue() instanceof Date, true, 'input date present');
eq(IN().getRange(6, C.STATUS).getValue(), '', 'failed row not submitted');

console.log('--- B. resubmit immediately (no duplicates)');
const dbCount = () => Util.lastNonEmptyRow(DB, 1, 2) - 1;
r = Submit.run();
eq(r.submitted.length, 0, 'nothing new'); eq(r.alreadySubmitted, 3, 'already 3'); eq(dbCount(), 3, 'db still 3');

console.log('--- C. user clears status+id of submitted row -> still no duplicate');
IN().getRange(5, C.STATUS).setValue(''); IN().getRange(5, C.TASK_ID).setValue('');
r = Submit.run();
eq(r.submitted.length, 0, 'no new'); eq(dbCount(), 3, 'db still 3 after cleared status');
ok(r.errors.some(e => e.row === 5 && /Duplikat/.test(e.messages.join())), 'row5 flagged duplicate');
IN().getRange(5, C.TASK_ID).setValue('TSK-00001');   // heal path
r = Submit.run();
eq(r.healed.map(h => h.id), ['TSK-00001'], 'healed'); eq(IN().getRange(5, C.STATUS).getValue(), 'SUBMITTED', 'row5 restored'); eq(dbCount(), 3, 'db 3');

console.log('--- D. fix errors, submit again: IDs continue & unique');
setIn(6, mk({ [C.TITLE]: 'Diperbaiki 6' })); setIn(7, mk({ [C.TITLE]: 'Anak', [C.PARENT]: 'tsk-00001' }));
r = Submit.run();
eq(r.submitted.map(s => s.id), ['TSK-00004', 'TSK-00005'], 'new ids ' + JSON.stringify(r.submitted.map(s => s.id)));
eq(dbv(6, 'Parent Task ID'), 'TSK-00001', 'parent canonical');
const ids = []; for (let i = 2; i <= 6; i++) ids.push(dbv(i, 'Task ID')); eq(new Set(ids).size, 5, 'unique ids');

console.log('--- E. crash during DB write: no partial state');
setIn(16, mk({ [C.TITLE]: 'Crash test' }));
const origAppend = Db.append; Db.append = () => { throw new Error('boom'); };
let threw = false; try { Submit.run(); } catch (e) { threw = true; }
Db.append = origAppend;
ok(threw, 'threw'); eq(IN().getRange(16, C.STATUS).getValue(), '', 'row16 not marked'); eq(dbCount(), 5, 'db unchanged'); eq(LOCKED, false, 'lock released');
r = Submit.run(); eq(r.submitted.map(s => s.id), ['TSK-00006'], 'ID reused after failed attempt');

console.log('--- F. progress log');
const PL = SS.getSheetByName('PROGRESS LOG');
const pr = (row, a) => { Object.keys(a).forEach(k => PL.getRange(row, +k).setValue(a[k])); };
// cols: C=3 task, D=4 by, F=6 new, G=7 prog, H=8 notes, I=9 delay
pr(2, { 3: 'TSK-00002', 4: 'Staf 02', 6: 'IN PROGRESS', 7: 0.5, 8: 'mulai' });
pr(3, { 3: 'TSK-00001', 4: 'Staf 01', 6: 'COMPLETED', 8: 'selesai semua' });
pr(4, { 3: 'TSK-00003', 4: 'Staf 03', 6: 'PENDING' });                         // no delay reason
pr(5, { 3: 'TSK-12345', 4: 'Staf 03', 6: 'IN PROGRESS' });                      // unknown task
pr(6, { 3: 'TSK-00004', 4: 'Staf 01', 6: 'IN PROGRESS', 7: 150 });              // invalid progress
pr(7, { 3: 'TSK-00004', 4: 'Staf 01', 6: 'COMPLETED', 7: 0.5 });                // completed with 50
pr(8, { 3: 'TSK-00005', 4: 'Staf 01', 6: 'PENDING', 9: 'menunggu owner' });
pr(9, { 3: 'TSK-00005', 6: 'IN PROGRESS' });                                    // missing updated by
let pg = Progress.run();
eq(pg.processed.map(p => p.logId), ['LOG-00001', 'LOG-00002', 'LOG-00003'], 'log ids');
eq(pg.errors.map(e => e.row), [4, 5, 6, 7, 9], 'progress error rows');
eq(dbv(3, 'Current Status'), 'IN PROGRESS', 'TSK-2 status'); eq(dbv(3, 'Progress Percentage'), 0.5, 'TSK-2 prog');
eq(dbv(2, 'Current Status'), 'COMPLETED', 'TSK-1 completed'); eq(dbv(2, 'Progress Percentage'), 1, 'TSK-1 100%');
ok(dbv(2, 'Completion Date') instanceof Date, 'completion date set'); eq(dbv(2, 'Result Notes'), 'selesai semua', 'result notes');
eq(dbv(6, 'Current Status'), 'PENDING', 'TSK-5 pending'); eq(dbv(6, 'Delay Reason'), 'menunggu owner', 'delay reason');
eq(PL.getRange(2, 5).getValue(), 'NEW', 'previous status logged'); eq(PL.getRange(3, 5).getValue(), 'NEW', 'prev 2'); eq(PL.getRange(8, 5).getValue(), 'NEW', 'prev 3');
ok(/ERROR: .*Delay Reason/.test(PL.getRange(4, 11).getValue()), 'pending needs delay: ' + PL.getRange(4, 11).getValue());
ok(/ERROR/.test(PL.getRange(7, 11).getValue()), 'completed 50 err'); ok(/tidak ada/.test(PL.getRange(5, 11).getValue()), 'unknown task');
const before = JSON.stringify([...DB.cells]); 
pg = Progress.run(); eq(pg.processed.length, 0, 'idempotent: no reprocess');
// closed task cannot be modified
pr(10, { 3: 'TSK-00001', 4: 'Staf 01', 6: 'IN PROGRESS' });
pg = Progress.run(); ok(pg.errors.length >= 1 && /sudah COMPLETED/.test(PL.getRange(10, 11).getValue()), 'closed task protected: ' + PL.getRange(10, 11).getValue());
eq(dbv(2, 'Current Status'), 'COMPLETED', 'still completed');
// Late completion requires delay reason
const row3 = Db.readAll(Db.open()).byId['TSK-00003'];
row3['Deadline'] = new Date(now0.getTime() - 86400000); DB.getRange(4, dbh['Deadline']).setValue(row3['Deadline']);
pr(11, { 3: 'TSK-00003', 4: 'Staf 03', 6: 'IN PROGRESS', 7: 0.2 });
pg = Progress.run(); ok(/melewati deadline/.test(PL.getRange(11, 11).getValue()), 'overdue update requires reason: ' + PL.getRange(11, 11).getValue());
pr(12, { 3: 'TSK-00003', 4: 'Staf 03', 6: 'IN PROGRESS', 7: 0.2, 9: 'material telat' });
pg = Progress.run(); eq(pg.processed.length, 1, 'with reason ok'); eq(PL.getRange(12, 1).getValue(), 'LOG-00004', 'log seq continues');

console.log('--- G. overdue classification');
const n = new Date(2026, 10, 19, 12, 0), H = 3600000;
eq(Overdue.classify('NEW', new Date(n - H), n, 3), 'OVERDUE', 'new past');
eq(Overdue.classify('PENDING', new Date(n - 1), n, 3), 'OVERDUE', 'pending past');
eq(Overdue.classify('IN PROGRESS', new Date(+n + H), n, 3), 'DUE_SOON', 'soon');
eq(Overdue.classify('IN PROGRESS', new Date(+n + 3 * 24 * H), n, 3), 'DUE_SOON', 'exactly 3d');
eq(Overdue.classify('IN PROGRESS', new Date(+n + 3 * 24 * H + 1), n, 3), 'OK', 'beyond 3d');
eq(Overdue.classify('COMPLETED', new Date(n - 100 * H), n, 3), 'CLOSED', 'completed never overdue');
eq(Overdue.classify('CANCELLED', new Date(n - 100 * H), n, 3), 'CLOSED', 'cancelled never overdue');
eq(Overdue.classify('NEW', '', n, 3), 'NO_DEADLINE', 'no deadline');

console.log('--- H. evidence');
const EV = SS.getSheetByName('EVIDENCE LOG');
const ev = (row, a) => Object.keys(a).forEach(k => EV.getRange(row, +k).setValue(a[k]));
ev(2, { 2: 'TSK-00002', 4: 'Staf 02', 5: 'foto 1', 6: 'https://drive.google.com/file/d/abc' });
ev(3, { 2: 'TSK-00002', 4: 'Staf 02', 6: 'ftp://x' });
ev(4, { 2: 'TSK-00002', 4: 'Staf 02', 6: 'https://drive.google.com/file/d/abc' });
ev(5, { 2: 'TSK-00009', 4: 'Staf 02', 6: 'https://drive.google.com/file/d/zzz' });
let evr = Evidence.run();
eq(evr.processed.map(p => p.id), ['EVD-00001'], 'evidence id'); eq(evr.errors.map(e => e.row), [3, 4, 5], 'evidence errors');
eq(dbv(3, 'Evidence Link'), 'https://drive.google.com/file/d/abc', 'evidence link in DB'); ok(EV.getRange(2, 3).getValue() instanceof Date, 'upload date');
// form handler
Evidence.fromForm({ 'Task ID': ['TSK-00003'], 'Uploaded By': ['Staf 03'], 'Evidence Description': ['x'], 'Foto / File': ['https://drive.google.com/open?id=1, https://drive.google.com/open?id=2'] });
eq(dbv(4, 'Evidence Link'), 'https://drive.google.com/open?id=2', 'form evidence last link'); eq(EV.getRange(7, 1).getValue(), 'EVD-00003', 'form evidence ids');

console.log('--- I. guard (locked rows)');
const sh = IN(); const rg = sh.getRange(5, C.TITLE); rg.setValue('DIUBAH');
Guard.onEdit({ range: rg, oldValue: 'Valid satu', value: 'DIUBAH' });
eq(sh.getRange(5, C.TITLE).getValue(), 'Valid satu', 'submitted row edit reverted');
setIn(15, { [C.TITLE]: 'belum kirim' }); const rg2 = sh.getRange(15, C.TITLE); rg2.setValue('boleh');
Guard.onEdit({ range: rg2, oldValue: 'belum kirim', value: 'boleh' });
eq(sh.getRange(15, C.TITLE).getValue(), 'boleh', 'unsubmitted row editable');
const st = sh.getRange(5, C.STATUS); st.setValue(''); Guard.onEdit({ range: st, oldValue: 'SUBMITTED', value: undefined });
eq(sh.getRange(5, C.STATUS).getValue(), 'SUBMITTED', 'status cannot be cleared');
const lg = PL.getRange(2, 8); lg.setValue('edit'); Guard.onEdit({ range: lg, oldValue: 'mulai', value: 'edit' });
eq(PL.getRange(2, 8).getValue(), 'mulai', 'processed log row edit reverted');

console.log('--- J. reminders & email');
eq(SS.getSheetByName('SETTINGS').getRange(14, 13).getValue(), 'EMAIL_ENABLED', 'row of EMAIL_ENABLED');
eq(Settings.load().cfg.EMAIL_ENABLED, 'YES', 'email enabled by default');
MAILS.length = 0;
const rem = Reminders.run(false);
ok(rem.ran && rem.overdue >= 1, 'reminder ran, overdue=' + rem.overdue);
ok(MAILS.length >= 2, 'mails sent: ' + MAILS.length);
ok(MAILS.some(m => m.to === 'staf03@example.com'), 'overdue PIC mailed'); ok(MAILS.some(m => m.to === 'manager@example.com'), 'admin digest');
const cnt = MAILS.length; const rem2 = Reminders.run(false); ok(!rem2.ran && MAILS.length === cnt, 'same-day rerun skipped');
// submit sends email to PIC with cc
MAILS.length = 0;
setIn(17, mk({ [C.TITLE]: 'Email test', [C.PIC]: 'Staf 04' }));
r = Submit.run(); ok(MAILS.length === 1 && MAILS[0].to === 'staf04@example.com' && MAILS[0].cc === 'manager@example.com', 'submit email to PIC + cc: ' + JSON.stringify(MAILS.map(m => [m.to, m.cc])));
eq(r.mail.sent, 1, 'mail count');

console.log('--- K. formula extension & capacity');
const dbo = Db.open();
Db.ensureFormulas(dbo, 1500);
ok(DB.getRange(1500, dbo.autoStart).getFormula() !== '', 'formulas extended to 1500'); 
let capErr = false; try { Db.ensureFormulas(dbo, 10002); } catch (e) { capErr = /penuh/.test(e.message); } ok(capErr, 'capacity guard');

console.log('--- L. selfTest');
const stt = SelfTest.run(); const bad = stt.filter(x => !x.ok).map(x => x.name);
console.log('   selfTest failures (expected: trigger / maybe none):', JSON.stringify(bad));

console.log('--- M. clearSubmittedRows keeps DB');
SpreadsheetApp.getUi = () => ({ alert: (t, m) => { ALERTS.push([t, m]); return 'YES'; }, ButtonSet: { YES_NO: 2, OK: 1 }, Button: { YES: 'YES' } });
const before2 = dbCount(); clearSubmittedRows();
eq(dbCount(), before2, 'DB unchanged'); eq(IN().getRange(5, C.STATUS).getValue(), '', 'submitted rows cleared'); ok(IN().getRange(15, C.TITLE).getValue() === 'boleh', 'unsubmitted preserved');


console.log('--- N. mobile forms');
const FR = CONFIG.FORMS;
SS.getSheetByName('SETTINGS').getRange(14, 14).setValue('NO');
r = (() => { setIn(18, mk({ [C.TITLE]: 'email off' })); return Submit.run(); })();
ok(/EMAIL_ENABLED = NO/.test(r.mail.skipped), 'skipped reason shown: ' + r.mail.skipped); ok(/EMAIL TIDAK DIKIRIM/.test(Submit.summaryText(r)), 'summary says email not sent');
SS.getSheetByName('SETTINGS').getRange(14, 14).setValue('YES');
const setupRes = Forms.setup();
eq(setupRes.created, ['Form Tugas', 'Form Progres'], 'forms created');
const tf = FormApp.openById(PROPS['FORM_TASK_ID']), pf = FormApp.openById(PROPS['FORM_PROGRESS_ID']);
const ch = (f, t) => f.items.find(i => i.title === t).choices;
eq(ch(tf, FR.TASK.PROJECT), ['Proyek Contoh A', 'Proyek Contoh B', 'Proyek Contoh C'], 'project choices');
eq(ch(tf, FR.TASK.PIC), ['Staf 01', 'Staf 02', 'Staf 03', 'Staf 04', 'Staf 05'], 'only active staff');
eq(ch(tf, FR.TASK.PRIORITY), ['Low', 'Normal', 'High', 'Urgent'], 'priority choices');
eq(tf.conf.email, 'VERIFIED', 'verified email collection');
eq(TRIGGERS.filter(t => t.kind === 'form').map(t => t.fn).sort(), ['onProgressFormSubmit', 'onTaskFormSubmit'], 'form triggers');
ok(pf.items.find(i => i.title === FR.PROGRESS.TASK).choices.every(c => /^TSK-\d+ \| /.test(c)), 'progress form lists open tasks');
ok(!pf.items.find(i => i.title === FR.PROGRESS.TASK).choices.some(c => c.startsWith('TSK-00001 ')), 'completed task not listed');
eq(ch(pf, FR.PROGRESS.STATUS), ['IN PROGRESS', 'PENDING', 'COMPLETED', 'CANCELLED'], 'status choices without NEW');
Forms.setup(); eq(TRIGGERS.filter(t => t.kind === 'form').length, 2, 'setup idempotent (no duplicate triggers)'); eq(Object.keys(FORMS).length, 2, 'no duplicate forms');

const ymd = off => Util.fmt(midnight(off), 'yyyy-MM-dd');
const T_ = FR.TASK;
const taskAns = (o) => Object.assign({ [T_.PROJECT]: 'Proyek Contoh B', [T_.PIC]: 'Staf 01', [T_.TITLE]: 'Dari HP', [T_.PRIORITY]: 'Urgent', [T_.DEADLINE]: ymd(5) }, o);
let before3 = dbCount(); MAILS.length = 0;
onTaskFormSubmit({ response: mkResp('manager@example.com', taskAns({ [T_.DESC]: 'desk hp', [T_.NOTES]: 'cepat' })) });
eq(dbCount(), before3 + 1, 'form task added');
const lastRow = Util.lastNonEmptyRow(DB, 1, 2);
eq(dbv(lastRow, 'Created By'), 'Manager', 'created by from email'); eq(dbv(lastRow, 'Submitted By (Email)'), 'manager@example.com', 'submitted by email');
eq(dbv(lastRow, 'Project ID'), 'PRJ-002', 'proj id'); eq(dbv(lastRow, 'Priority'), 'Urgent', 'prio');
ok(dbv(lastRow, 'Deadline').getHours() === 23 && dbv(lastRow, 'Deadline').getMinutes() === 59, 'form deadline eod');
ok(/Catatan: cepat/.test(dbv(lastRow, 'Task Description')), 'form notes appended');
const newId = dbv(lastRow, 'Task ID'); ok(/^TSK-\d{5}$/.test(newId), 'form id ' + newId);
ok(MAILS.some(m => m.to === 'staf01@example.com' && m.cc === 'manager@example.com'), 'PIC mailed (cc admin) from form: ' + JSON.stringify(MAILS.map(m => [m.to, m.cc])));
ok(MAILS.every(m => typeof m.body === 'string' && m.body.length > 0), 'plain-text body present');
// form: second admin via different email
onTaskFormSubmit({ response: mkResp('atasan@example.com', taskAns({ [T_.TITLE]: 'Dari atasan', [T_.START]: ymd(1), [T_.DEADLINE]: ymd(9) })) });
eq(dbv(Util.lastNonEmptyRow(DB, 1, 2), 'Created By'), 'Atasan', 'created by Atasan from email');
// duplicate form submit
const c0 = dbCount(); MAILS.length = 0;
onTaskFormSubmit({ response: mkResp('manager@example.com', taskAns()) });
eq(dbCount(), c0, 'duplicate form submit rejected'); ok(MAILS.some(m => m.to === 'manager@example.com' && /Duplikat/.test(m.body)), 'sender told why');
// non admin
MAILS.length = 0;
onTaskFormSubmit({ response: mkResp('orang.luar@gmail.com', taskAns({ [T_.TITLE]: 'bukan admin' })) });
eq(dbCount(), c0, 'non-admin rejected'); ok(MAILS.some(m => m.to === 'orang.luar@gmail.com' && /bukan admin/.test(m.body)), 'non-admin told');
// no email at all
onTaskFormSubmit({ response: mkResp('', taskAns({ [T_.TITLE]: 'tanpa email' })) }); eq(dbCount(), c0, 'blank email rejected');
// past deadline
onTaskFormSubmit({ response: mkResp('manager@example.com', taskAns({ [T_.TITLE]: 'lewat', [T_.DEADLINE]: ymd(-2) })) }); eq(dbCount(), c0, 'past deadline rejected');
const SL = SS.getSheetByName('SYSTEM LOG'); const slText = () => { let a = []; for (let i = 2; i <= SL.getLastRow(); i++) a.push([2, 3, 4].map(c => SL.getRange(i, c).getValue()).join(' ')); return a.join('\n'); };
ok(/FORM-TASK WARN DITOLAK/.test(slText()) && /MAIL INFO TERKIRIM/.test(slText()), 'SYSTEM LOG records rejections and mail');

console.log('--- O. progress form');
const PRG = FR.PROGRESS;
const choiceFor = id => FormApp.openById(PROPS['FORM_PROGRESS_ID']).items.find(i => i.title === PRG.TASK).choices.find(c => c.startsWith(id + ' '));
const mine = Db.readAll(Db.open()).rows.find(o => o['Assigned To'] === 'Staf 01' && !['COMPLETED', 'CANCELLED'].includes(o['Current Status']) && o['Deadline'] > new Date());
const ch1 = choiceFor(mine['Task ID']); ok(!!ch1, 'my task listed in progress form: ' + mine['Task ID']);
const pAns = (o) => Object.assign({ [PRG.TASK]: ch1, [PRG.STATUS]: 'IN PROGRESS', [PRG.PROGRESS]: '30%', [PRG.NOTES]: 'jalan' }, o);
const rowOf = id => { const d = Db.readAll(Db.open()); return d.byId[id]; };
onProgressFormSubmit({ response: mkResp('staf01@example.com', pAns()) });
eq(rowOf(mine['Task ID'])['Current Status'], 'IN PROGRESS', 'staff updated own task'); eq(rowOf(mine['Task ID'])['Progress Percentage'], 0.3, 'progress 30%');
ok(/staf01|Staf 01/.test(String(PL.getRange(PL.getLastRow(), 4).getValue())), 'Updated By = staff name from email');
// other staff cannot update
MAILS.length = 0; const plRows = PL.getLastRow();
onProgressFormSubmit({ response: mkResp('staf02@example.com', pAns({ [PRG.PROGRESS]: '60%' })) });
eq(rowOf(mine['Task ID'])['Progress Percentage'], 0.3, 'other staff rejected'); ok(MAILS.some(m => m.to === 'staf02@example.com' && /bukan PIC/.test(m.body)), 'told not PIC'); eq(PL.getLastRow(), plRows, 'no log row for rejected (not PIC)');
// staff cannot cancel
onProgressFormSubmit({ response: mkResp('staf01@example.com', pAns({ [PRG.STATUS]: 'CANCELLED' })) }); eq(rowOf(mine['Task ID'])['Current Status'], 'IN PROGRESS', 'staff cannot cancel');
// staff PENDING without reason -> rejected by rules, row cleared
MAILS.length = 0;
onProgressFormSubmit({ response: mkResp('staf01@example.com', pAns({ [PRG.STATUS]: 'PENDING' })) });
eq(rowOf(mine['Task ID'])['Current Status'], 'IN PROGRESS', 'pending without reason rejected'); ok(MAILS.some(m => /Delay Reason/.test(m.body)), 'told delay reason needed');
const lr = PL.getLastRow(); eq(Util.str(PL.getRange(lr, 3).getValue()) === '' || lr < plRows + 2, true, 'rejected form row cleared');
const pr2 = Progress.run(); eq(pr2.errors.filter(e => e.row >= plRows).length, 0, 'no lingering errors reprocessed');
// pending with reason ok, then admin cancels
onProgressFormSubmit({ response: mkResp('staf01@example.com', pAns({ [PRG.STATUS]: 'PENDING', [PRG.DELAY]: 'tunggu material' })) });
eq(rowOf(mine['Task ID'])['Current Status'], 'PENDING', 'pending with reason ok');
onProgressFormSubmit({ response: mkResp('manager@example.com', pAns({ [PRG.STATUS]: 'CANCELLED', [PRG.PROGRESS]: '(tidak berubah)' })) });
eq(rowOf(mine['Task ID'])['Current Status'], 'CANCELLED', 'admin can cancel');
ok(!choiceFor(mine['Task ID']), 'cancelled task removed from progress form choices');
// unknown account
MAILS.length = 0; onProgressFormSubmit({ response: mkResp('asing@gmail.com', pAns()) }); ok(MAILS.some(m => m.to === 'asing@gmail.com'), 'unknown account told');

console.log('--- P. selfTest after forms');
const st2 = SelfTest.run().filter(x => !x.ok).map(x => x.name); console.log('   failures:', JSON.stringify(st2));
ok(st2.indexOf('Tidak ada email contoh (@example.com)') >= 0, 'self test warns about example.com emails');
ok(st2.indexOf('Form HP tugas & progres dibuat') < 0, 'forms detected');

