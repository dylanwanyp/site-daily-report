/**
 * 外勤工程地盤日報系統 — Google Apps Script Backend
 * Deploy: Execute as me → Access: Anyone (even anonymous)
 */

// ── Sheet Tab Names ──
var CFG = 'config';
var SITES = 'sites';
var EMPLOYEES = 'employees';
var SUBS = 'subsidiaries';
var RECS = 'records';

// ── Entry Points ──
function doGet(e) { return route(e, false); }
function doPost(e) { return route(e, true); }

function route(e, isPost) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    return routeAction(action, e, isPost);
  } catch (err) {
    return json({error: err.message || 'Unknown error'});
  }
}

function routeAction(action, e, isPost) {
  switch (action) {
    case 'getSites':         return json(getSites());
    case 'getEmployees':     return json(getEmployees());
    case 'getSubsidiaries':  return json(getSubs());
    case 'getAllData':       return json({sites: getSites(), employees: getEmployees(), subsidiaries: getSubs()});
    case 'checkMissing':     return json(checkMissingDays(e));
    case 'getExistingDates': return json(getExistingDates(e));
    case 'submitRecords':    return json(handleSubmit(isPost ? JSON.parse(e.postData.contents) : JSON.parse(e.parameter.data || '[]')));
    case 'verifyPassword':   return json(verifyPw(e, isPost));
    case 'getRecordsForReport':
      var auth = checkAdmin(e, isPost);
      if (!auth.ok) return json(auth);
      return json(getReport(e, isPost));
    case 'adminUpdateSite': case 'adminDeleteSite':
    case 'adminAddEmployee': case 'adminDeleteEmployee':
    case 'adminAddSubsidiary': case 'adminDeleteSubsidiary':
      return json(adminAction(action, e, isPost));
    case 'getSubsidiaryReport':
      var a2 = checkAdmin(e, isPost);
      if (!a2.ok) return json(a2);
      return json(getSubsidiaryReport(e, isPost));
    default: return json({error: 'Unknown action: ' + action});
  }
}

// ── Admin Guard ──
function checkAdmin(e, isPost) {
  var pw = '';
  if (isPost && e.postData) {
    try {
      var body = JSON.parse(e.postData.contents || '{}');
      pw = body.password || '';
    } catch (exc) {
      pw = (e && e.parameter && e.parameter.password) || '';
    }
  } else {
    pw = (e && e.parameter && e.parameter.password) || '';
  }
  var stored = getConfig('admin_password');
  return stored === pw ? {ok: true} : {ok: false, error: 'Invalid password'};
}

function verifyPw(e, isPost) {
  var pw = '';
  if (isPost && e.postData) {
    try {
      var body = JSON.parse(e.postData.contents || '{}');
      pw = body.password || '';
    } catch (exc) {
      pw = (e && e.parameter && e.parameter.password) || '';
    }
  } else {
    pw = (e && e.parameter && e.parameter.password) || '';
  }
  var stored = getConfig('admin_password');
  return {valid: stored === pw};
}

function adminAction(action, e, isPost) {
  // Support both GET query params and POST body
  var body = parseParams(e, isPost);
  var auth = checkAdmin(e, isPost);
  if (!auth.ok) return auth;
  switch (action) {
    case 'adminUpdateSite':       return updateSite(body);
    case 'adminDeleteSite':       return deleteSite(body);
    case 'adminAddEmployee':      return addEmployee(body);
    case 'adminDeleteEmployee':   return deleteEmployee(body);
    case 'adminAddSubsidiary':    return addSub(body);
    case 'adminDeleteSubsidiary': return deleteSub(body);
    default:                      return {error: 'Unknown admin action'};
  }
}

// ── Config ──
function getConfig(key) {
  var s = ss().getSheetByName(CFG);
  if (!s) return '';
  var d = s.getDataRange().getValues();
  for (var i = 0; i < d.length; i++) {
    if (String(d[i][0]).trim() === key) return String(d[i][1] || '');
  }
  return '';
}

// ── Public Getters ──
function getSites() {
  return readSheet(SITES).map(function(r) { return {name: r[0], subsidiary: r[1] || ''}; });
}
function getEmployees() {
  return readSheet(EMPLOYEES).map(function(r) { return {name: r[0], monthStartDay: Number(r[1]) || 1}; });
}
function getSubs() {
  return readSheet(SUBS).map(function(r) { return {name: r[0]}; });
}

// ── Missing Days ──
function checkMissingDays(e) {
  var emp = (e.parameter.employee || '').trim();
  if (!emp) return {missingDates: []};
  var rows = readSheet(RECS);
  var empRecs = {};
  for (var i = 0; i < rows.length; i++)
    if (String(rows[i][1]).trim() === emp) empRecs[String(rows[i][0]).trim()] = true;
  var today = new Date();
  var missing = [];
  for (var n = 1; n <= 30; n++) {
    var dt = new Date(today);
    dt.setDate(dt.getDate() - n);
    var ds = fmtDate(dt);
    if (!empRecs[ds]) missing.push(ds);
  }
  missing.reverse();
  return {missingDates: missing};
}

// ── Duplicate Check ──
function getExistingDates(e) {
  var emp = (e.parameter.employee || '').trim();
  var datesParam = e.parameter.dates || '';
  if (!emp || !datesParam) return {existingDates: []};
  var target = datesParam.split(',');
  var rows = readSheet(RECS);
  var exist = {};
  for (var i = 0; i < rows.length; i++)
    if (String(rows[i][1]).trim() === emp) exist[String(rows[i][0]).trim()] = true;
  var out = [];
  for (var j = 0; j < target.length; j++)
    if (exist[target[j].trim()]) out.push(target[j].trim());
  return {existingDates: out};
}

// ── Submit ──
function handleSubmit(records) {
  if (!records || !records.length) return {success: false, count: 0, error: 'No records'};
  var sheet = ss().getSheetByName(RECS);
  if (!sheet) return {success: false, count: 0, error: 'Sheet not found'};
  var now = new Date();
  var rows = [];
  for (var i = 0; i < records.length; i++) {
    var r = records[i];
    rows.push([r.date||'', r.employee||'', r.site||'', r.subsidiary||'', Number(r.hours)||0, r.note||'', fmtDateTime(now)]);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  return {success: true, count: rows.length};
}

// ── Report ──
function getReport(e, isPost) {
  var p = parseParams(e, isPost);
  var emp = (p.employee || '').trim();
  var month = p.month || '';
  if (!emp) return {records: [], error: 'Employee required'};
  if (!month) return {records: [], error: 'Month required'};

  var parts = month.split('-');
  var year = parseInt(parts[0], 10);
  var mon = parseInt(parts[1], 10);
  if (emp === 'all') {
    var rStart = new Date(year, mon - 1, 1);
    var rEnd = new Date(year, mon, 0, 23, 59, 59);
    var rows = readSheet(RECS);
    var out = [];
    for (var j = 0; j < rows.length; j++) {
      var d = new Date(rows[j][0]);
      if (isNaN(d.getTime())) continue;
      if (d >= rStart && d <= rEnd)
        out.push({date: fmtDate(d), employee: rows[j][1], site: rows[j][2], subsidiary: rows[j][3], hours: rows[j][4], note: rows[j][5]});
    }
    return {records: out, rangeStart: fmtDate(rStart), rangeEnd: fmtDate(rEnd), mode: 'all'};
  }
  var emps = readSheet(EMPLOYEES);
  var startDay = 1;
  for (var i = 0; i < emps.length; i++)
    if (String(emps[i][0]).trim() === emp) { startDay = Number(emps[i][1]) || 1; break; }

 var rStart, rEnd;
  if (startDay === 1) {
    rStart = new Date(year, mon - 1, 1);
    rEnd = new Date(year, mon, 0, 23, 59, 59);
  } else {
    var pm = mon - 2;
    var py = year;
    if (pm < 0) { pm = 11; py--; }
    rStart = new Date(py, pm, startDay);
    rEnd = new Date(year, mon - 1, startDay - 1, 23, 59, 59);
    if (mon === 1) {
      rStart = new Date(year - 1, 11, startDay);
      rEnd = new Date(year, 0, startDay - 1, 23, 59, 59);
    }
  }

  var rows = readSheet(RECS);
  var out = [];
  for (var j = 0; j < rows.length; j++) {
    if (String(rows[j][1]).trim() !== emp) continue;
    var d = new Date(rows[j][0]);
    if (isNaN(d.getTime())) continue;
    if (d >= rStart && d <= rEnd)
      out.push({date: fmtDate(d), employee: rows[j][1], site: rows[j][2], subsidiary: rows[j][3], hours: rows[j][4], note: rows[j][5]});
  }
  return {records: out, rangeStart: fmtDate(rStart), rangeEnd: fmtDate(rEnd)};
}

// ── Uniform param parser (GET query params + POST body) ──
function parseParams(e, isPost) {
  var body = {};
  if (e && e.parameter) {
    for (var key in e.parameter) {
      if (key !== 'action') body[key] = e.parameter[key];
    }
  }
  if (isPost && e.postData && e.postData.contents) {
    try {
      var p = JSON.parse(e.postData.contents);
      for (var k in p) {
        if (k !== 'action') body[k] = p[k];
      }
    } catch (exc) {
      // POST body parse failed, params already from e.parameter
    }
  }
  return body;
}

// ── Site CRUD ──
function updateSite(b) {
  var n = (b.name||'').trim(), sub = (b.subsidiary||'').trim();
  if (!n) return {error: 'Name required'};
  var s = ss().getSheetByName(SITES);
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++)
    if (String(rows[i][0]).trim() === n) { s.getRange(i+1,2).setValue(sub); return {success:true}; }
  s.appendRow([n, sub]);
  return {success: true};
}
function deleteSite(b) { return deleteRow(SITES, b.name); }

// ── Employee CRUD ──
function addEmployee(b) {
  var n = (b.name||'').trim();
  if (!n) return {error: 'Name required'};
  var sd = b.monthStartDay || 1;
  var s = ss().getSheetByName(EMPLOYEES);
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++)
    if (String(rows[i][0]).trim() === n) return {error: 'Already exists'};
  s.appendRow([n, sd]);
  return {success: true};
}
function deleteEmployee(b) { return deleteRow(EMPLOYEES, b.name); }

// ── Subsidiary CRUD ──
function addSub(b) {
  var n = (b.name||'').trim();
  if (!n) return {error: 'Name required'};
  var s = ss().getSheetByName(SUBS);
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++)
    if (String(rows[i][0]).trim() === n) return {error: 'Already exists'};
  s.appendRow([n]);
  return {success: true};
}
function deleteSub(b) { return deleteRow(SUBS, b.name); }

// ── Shared Helpers ──
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function readSheet(n) {
  var s = ss().getSheetByName(n);
  if (!s) return [];
  var d = s.getDataRange().getValues();
  return d.length > 1 ? d.slice(1) : [];
}
function deleteRow(sn, val) {
  var s = ss().getSheetByName(sn);
  if (!s) return {error: 'Sheet not found'};
  var rows = s.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++)
    if (String(rows[i][0]).trim() === (val||'').trim()) { s.deleteRow(i+1); return {success: true}; }
  return {error: 'Not found'};
}
function json(d) { return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); }
function fmtDate(d) {
  return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2);
}
function fmtDateTime(d) {
  return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2)+' '+('0'+d.getHours()).slice(-2)+':'+('0'+d.getMinutes()).slice(-2);
}
function ds(v) { var s = String(v).substring(0,10); return s; }

// ── Subsidiary Report ──
function getSubsidiaryReport(e, isPost) {
  var p = parseParams(e, isPost);
  var subName = (p.subsidiary || '').trim();
  var month = p.month || '';
  if (!subName) return {error: 'Subsidiary required'};
  if (!month) return {error: 'Month required'};
  var parts = month.split('-');
  var year = parseInt(parts[0], 10);
  var mon = parseInt(parts[1], 10);
  var rStart = new Date(year, mon - 1, 1);
  var rEnd = new Date(year, mon, 0, 23, 59, 59);
 var rows = readSheet(RECS);
  if (subName === 'all') {
    var allData = {};
    for (var j = 0; j < rows.length; j++) {
      var d = new Date(rows[j][0]);
      if (isNaN(d.getTime())) continue;
      if (d < rStart || d > rEnd) continue;
      var emp = String(rows[j][1]).trim();
      var site = String(rows[j][2]).trim();
      var sub = String(rows[j][3] || '').trim();
      if (!sub) continue;
      if (!allData[sub]) allData[sub] = {};
      if (!allData[sub][site]) allData[sub][site] = {};
      if (!allData[sub][site][emp]) allData[sub][site][emp] = 0;
      allData[sub][site][emp]++;
    }
    var out = [];
    Object.keys(allData).sort().forEach(function(sn) {
      var siteList = [];
      var subTotal = 0;
      Object.keys(allData[sn]).sort().forEach(function(siteName) {
        var wl = [];
        Object.keys(allData[sn][siteName]).sort().forEach(function(en) {
          wl.push({name: en, days: allData[sn][siteName][en]});
        });
        var sd = 0;
        Object.keys(allData[sn][siteName]).forEach(function(k) { sd += allData[sn][siteName][k]; });
        siteList.push({site: siteName, totalWorkerDays: sd, workers: wl});
        subTotal += sd;
      });
      out.push({subsidiary: sn, sites: siteList, totalWorkerDays: subTotal});
    });
    return {mode: 'all', subsidiaries: out, rangeStart: fmtDate(rStart), rangeEnd: fmtDate(rEnd)};
  }
  var siteEmpDays = {};
  var siteTotals = {};
 for (var j = 0; j < rows.length; j++) {
    if (String(rows[j][3] || '').trim() !== subName) continue;
    var d = new Date(rows[j][0]);
    if (isNaN(d.getTime())) continue;
    if (d < rStart || d > rEnd) continue;
    var emp = String(rows[j][1]).trim();
    var site = String(rows[j][2]).trim();
    if (!siteEmpDays[site]) siteEmpDays[site] = {};
    if (!siteEmpDays[site][emp]) siteEmpDays[site][emp] = 0;
    siteEmpDays[site][emp]++;
    if (!siteTotals[site]) siteTotals[site] = 0;
    siteTotals[site]++;
  }
  var out = [];
  Object.keys(siteTotals).sort().forEach(function(sn) {
    var wl = [];
    Object.keys(siteEmpDays[sn]).sort().forEach(function(en) {
      wl.push({name: en, days: siteEmpDays[sn][en]});
    });
    out.push({site: sn, totalWorkerDays: siteTotals[sn], workers: wl});
  });
  var gt = 0;
  Object.keys(siteTotals).forEach(function(s) { gt += siteTotals[s]; });
  return {subsidiary: subName, rangeStart: fmtDate(rStart), rangeEnd: fmtDate(rEnd), sites: out, grandTotal: gt};
}
