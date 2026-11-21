console.log('--- Y. guard fungsi menu dari web');
(function(){ const keep = ACTIVE_EMAIL; const n0 = MAILS.length; ACTIVE_EMAIL = '';
  const uiBak = Util.ui; Util.ui = () => null;
  const r = dailyJob.length === 0 && runDailyReminderNow(); ok(r === null, 'runDailyReminderNow ditolak dari konteks web');
  ok(MAILS.length === n0, 'tidak ada email terkirim');
  ACTIVE_EMAIL = 'manager@example.com'; const r2 = selfTest(); ok(Array.isArray(r2), 'selfTest jalan di konteks editor (email aktif = efektif)');
  Util.ui = uiBak; ACTIVE_EMAIL = keep; })();
