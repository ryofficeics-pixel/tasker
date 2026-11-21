let PASS = 0, FAIL = 0;
function ok(c, m) { if (c) PASS++; else { FAIL++; console.log('  FAIL:', m); } }
function eq(a, b, m) { ok(JSON.stringify(a) === JSON.stringify(b), m + ' got ' + JSON.stringify(a)); }
