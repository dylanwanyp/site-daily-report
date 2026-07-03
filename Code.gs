/**
 * 外勤工程地盤日報系統 — Google Apps Script Backend
 * Last amended: 2026-07-03 (v15 - fixed ds() Date check)
 * Deploy: Execute as me → Access: Anyone (even anonymous)
 */

// ── Sheet Tab Names ──
var CFG = 'config';
var SITES = 'sites';
var EMPLOYEES = 'employees';
var SUBS = 'subsidiaries';
var RECS = 'records';
var APPR = 'approvals';

// ── Entry Points ──
function doGet(e) { return route(e, false); }

// ── Debug endpoint (no auth) ──
function debugGetAllRecords() {
  var rows = readSheet(RECS);
  return {total: rows.length, records: rows.slice(0, 50).map(function(r) {
    var dv = r[0];
    return {
      date: String(dv || "").substring(0,20),
      dateIsDate: dv && typeof dv.getMonth === 'function',
      dateType: typeof dv,
      employee: String(r[1] || ""),
      site: String(r[2] || ""),
      subsidiary: String(r[3] || ""),
      hours: r[4],
      note: String(r[5] || ""),
      overtimeType: String(r.length > 7 ? r[7] : "")
    };
  })};
}
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
    case 'getSites':          return json(getSites());
    case 'getEmployees':      return json(getEmployees());
    case 'getSubsidiaries':   return json(getSubs());
    case 'getAllData':        return json({sites: getSites(), employees: getEmployees(), subsidiaries: getSubs()});
    case 'debugGetAllRecords': return json(debugGetAllRecords());

    case 'checkMissing':      return json(checkMissingDays(e));
    case 'getExistingDates':  return json(getExistingDates(e));
    case 'submitRecords':     return json(handleSubmit(isPost ? JSON.parse(e.postData.contents) : JSON.parse(e.parameter.data || '[]')));
    case 'verifyPassword':    return json(verifyPw(e, isPost));
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
    case 'approveEmployee':
    case 'unapproveEmployee':
    case 'getApprovals':
    case 'getApprovalStatus':
      var a3 = checkAdmin(e, isPost);
      if (!a3.ok) return json(a3);
      return json(approvalAction(action, e, isPost));
    default:
      return json({error: 'Unknown action: ' + action});
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

// ── Approval Actions ──
function approvalAction(action, e, isPost) {
  var body = parseParams(e, isPost);
  switch (action) {
    case 'approveEmployee':
      return approveEmployee(body);
    case 'unapproveEmployee':
      return unapproveEmployee(body);
    case 'getApprovals':
      return getApprovals(body);
    case 'getApprovalStatus':
      return getApprovalStatus(body);
    default:
      return {error: 'Unknown approval action'};
  }
}

function approveEmployee(b) {
  var emp = (b.employee || '').trim();
  var month = (b.month || '').trim();
  if (!emp || !month) return {error: 'Employee and month required'};
  var sheet = ss().getSheetByName(APPR);
  if (!sheet) {
    sheet = ss().insertSheet(APPR);
    sheet.appendRow(['員工姓名', '月份', '審批時間']);
  }
  // Remove existing entry for this employee+month
  var rows = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === emp && String(rows[i][1]).trim() === month) {
      foundRow = i + 1;
      break;
    }
  }
  if (foundRow > 0) {
    sheet.deleteRow(foundRow);
  }
  sheet.appendRow([emp, month, fmtDateTime(new Date())]);
  return {success: true};
}

function unapproveEmployee(b) {
  var emp = (b.employee || '').trim();
  var month = (b.month || '').trim();
  if (!emp || !month) return {error: 'Employee and month required'};
  var sheet = ss().getSheetByName(APPR);
  if (!sheet) return {success: false, error: 'Approvals sheet not found'};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === emp && String(rows[i][1]).trim() === month) {
      sheet.deleteRow(i + 1);
      return {success: true};
    }
  }
  return {success: false, error: 'Approval not found'};
}

function getApprovals(b) {
  var month = (b.month || '').trim();
  var sheet = ss().getSheetByName(APPR);
  if (!sheet) return {approvals: []};
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return {approvals: []};
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (month && String(rows[i][1]).trim() !== month) continue;
    out.push({employee: String(rows[i][0]).trim(), month: String(rows[i][1]).trim(), approvedAt: String(rows[i][2] || '')});
  }
  return {approvals: out};
}

function getApprovalStatus(b) {
  var month = (b.month || '').trim();
  var sheet = ss().getSheetByName(APPR);
  if (!sheet) return {status: {}};
  var rows = sheet.getDataRange().getValues();
  var status = {};
  for (var i = 1; i < rows.length; i++) {
    var m = String(rows[i][1]).trim();
    var emp = String(rows[i][0]).trim();
    if (month && m !== month) continue;
    status[emp] = {approved: true, month: m, approvedAt: String(rows[i][2] || '')};
  }
  return {status: status};
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
    if (String(rows[i][1]).trim() === emp) empRecs[ds(rows[i][0])] = true;
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
    if (String(rows[i][1]).trim() === emp) exist[ds(rows[i][0])] = true;
  var out = [];
  for (var j = 0; j < target.length; j++)
    if (exist[target[j].trim()]) out.push(target[j].trim());
  return {existingDates: out};
}

// ── Submit ──
// Records columns: 0=date,1=employee,2=site,3=subsidiary,4=hours,5=note,6=timestamp,7=overtimeType
function handleSubmit(records) {
  if (!records || !records.length) return {success: false, count: 0, error: 'No records'};
  var sheet = ss().getSheetByName(RECS);
  if (!sheet) return {success: false, count: 0, error: 'Sheet not found'};
  var now = new Date();
  var rows = [];
  for (var i = 0; i < records.length; i++) {
    var r = records[i];
    rows.push([r.date||'', r.employee||'', r.site||'', r.subsidiary||'', Number(r.hours)||0, r.note||'', fmtDateTime(now), r.overtimeType||'正常工時']);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  return {success: true, count: rows.length};
}

// ── Weighted Hours Calculator ──
function calcWeightedHours(hours, overtimeType) {
  var h = Number(hours) || 0;
  var t = (overtimeType || '正常工時').trim();
  var multiplier = 1.0;
  if (t === '晚間加班') multiplier = 1.5;
  else if (t === '深夜加班') multiplier = 2.0;
  return Math.round(h * multiplier * 10) / 10;
}

function formatRecordRow(row) {
  var dateStr = ds(row[0]);
  var overtimeType = row.length > 7 ? String(row[7] || '正常工時').trim() : '正常工時';
  var hours = Number(row[4]) || 0;
  return {
    date: dateStr,
    employee: String(row[1] || '').trim(),
    site: String(row[2] || '').trim(),
    subsidiary: String(row[3] || '').trim(),
    hours: hours,
    note: String(row[5] || '').trim(),
    timestamp: String(row[6] || ''),
    overtimeType: overtimeType,
    weightedHours: calcWeightedHours(hours, overtimeType)
  };
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

  // Get approval status
  var appStatus = {};
  var appSheet = ss().getSheetByName(APPR);
  if (appSheet) {
    var appRows = appSheet.getDataRange().getValues();
    for (var ai = 1; ai < appRows.length; ai++) {
      var am = String(appRows[ai][1]).trim();
      var ae = String(appRows[ai][0]).trim();
      if (am === month) {
        appStatus[ae] = {approved: true, approvedAt: String(appRows[ai][2] || '')};
      }
    }
  }

  if (emp === 'all') {
    var rStartStr = fmtDate(new Date(year, mon - 1, 1));
    var rEndStr = fmtDate(new Date(year, mon, 0));
    var rows = readSheet(RECS);
    var out = [];
    for (var j = 0; j < rows.length; j++) {
      var dateStr = ds(rows[j][0]);
      if (!dateStr) continue;
      if (dateStr >= rStartStr && dateStr <= rEndStr) {
        out.push(formatRecordRow(rows[j]));
      }
    }
    return {records: out, rangeStart: rStartStr, rangeEnd: rEndStr, mode: 'all', approvalStatus: appStatus};
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

  var rStartStr = fmtDate(rStart);
  var rEndStr = fmtDate(rEnd);
  var rows = readSheet(RECS);
  var out = [];
  for (var j = 0; j < rows.length; j++) {
    if (String(rows[j][1]).trim() !== emp) continue;
    var dateStr = ds(rows[j][0]);
    if (!dateStr) continue;
    if (dateStr >= rStartStr && dateStr <= rEndStr) {
      out.push(formatRecordRow(rows[j]));
    }
  }
  return {records: out, rangeStart: fmtDate(rStart), rangeEnd: fmtDate(rEnd), approvalStatus: appStatus};
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
    } catch (exc) { }
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
function ds(v) {
  if (v && typeof v.getMonth === 'function') return fmtDate(v);
  return String(v || '').substring(0, 10);
}

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
  var rStartStr = fmtDate(rStart);
  var rEndStr = fmtDate(rEnd);

  var rows = readSheet(RECS);
  if (subName === 'all') {
    var allData = {};
    for (var j = 0; j < rows.length; j++) {
      var dateStr = ds(rows[j][0]);
      if (!dateStr) continue;
      if (dateStr < rStartStr || dateStr > rEndStr) continue;
      var emp = String(rows[j][1]).trim();
      var site = String(rows[j][2]).trim();
      var sub = String(rows[j][3] || '').trim();
      if (!sub) continue;
      if (!allData[sub]) allData[sub] = {};
      if (!allData[sub][site]) allData[sub][site] = {};
      if (!allData[sub][site][emp]) allData[sub][site][emp] = {days: 0, totalHours: 0, totalWeighted: 0};
      allData[sub][site][emp].days++;
      var h = Number(rows[j][4]) || 0;
      allData[sub][site][emp].totalHours += h;
      var ot = rows[j].length > 7 ? String(rows[j][7] || '正常工時').trim() : '正常工時';
      allData[sub][site][emp].totalWeighted += calcWeightedHours(h, ot);
    }
    var out = [];
    Object.keys(allData).sort().forEach(function(sn) {
      var siteList = [];
      var subTotal = 0;
      var subWeightedTotal = 0;
      Object.keys(allData[sn]).sort().forEach(function(siteName) {
        var wl = [];
        Object.keys(allData[sn][siteName]).sort().forEach(function(en) {
          wl.push({name: en, days: allData[sn][siteName][en].days, totalHours: Math.round(allData[sn][siteName][en].totalHours*10)/10, totalWeighted: Math.round(allData[sn][siteName][en].totalWeighted*10)/10});
        });
        var sd = 0;
        var sw = 0;
        Object.keys(allData[sn][siteName]).forEach(function(k) { sd += allData[sn][siteName][k].days; sw += allData[sn][siteName][k].totalWeighted; });
        siteList.push({site: siteName, totalWorkerDays: sd, totalWeighted: Math.round(sw*10)/10, workers: wl});
        subTotal += sd;
        subWeightedTotal += sw;
      });
      out.push({subsidiary: sn, sites: siteList, totalWorkerDays: subTotal, totalWeighted: Math.round(subWeightedTotal*10)/10});
    });
    return {mode: 'all', subsidiaries: out, rangeStart: rStartStr, rangeEnd: rEndStr};
  }

  var siteEmpDays = {};
  var siteTotals = {};
  var siteWeightedTotals = {};
  for (var j = 0; j < rows.length; j++) {
    if (String(rows[j][3] || '').trim() !== subName) continue;
    var dateStr = ds(rows[j][0]);
    if (!dateStr) continue;
    if (dateStr < rStartStr || dateStr > rEndStr) continue;
    var emp = String(rows[j][1]).trim();
    var site = String(rows[j][2]).trim();
    if (!siteEmpDays[site]) siteEmpDays[site] = {};
    if (!siteEmpDays[site][emp]) siteEmpDays[site][emp] = {days: 0, totalHours: 0, totalWeighted: 0};
    siteEmpDays[site][emp].days++;
    var h = Number(rows[j][4]) || 0;
    siteEmpDays[site][emp].totalHours += h;
    var ot = rows[j].length > 7 ? String(rows[j][7] || '正常工時').trim() : '正常工時';
    siteEmpDays[site][emp].totalWeighted += calcWeightedHours(h, ot);
    if (!siteTotals[site]) siteTotals[site] = 0;
    siteTotals[site]++;
    if (!siteWeightedTotals[site]) siteWeightedTotals[site] = 0;
    siteWeightedTotals[site] += calcWeightedHours(h, ot);
  }
  var out = [];
  Object.keys(siteTotals).sort().forEach(function(sn) {
    var wl = [];
    Object.keys(siteEmpDays[sn]).sort().forEach(function(en) {
      wl.push({name: en, days: siteEmpDays[sn][en].days, totalHours: Math.round(siteEmpDays[sn][en].totalHours*10)/10, totalWeighted: Math.round(siteEmpDays[sn][en].totalWeighted*10)/10});
    });
    out.push({site: sn, totalWorkerDays: siteTotals[sn], totalWeighted: Math.round(siteWeightedTotals[sn]*10)/10, workers: wl});
  });
  var gt = 0;
  var gtw = 0;
  Object.keys(siteTotals).forEach(function(s) { gt += siteTotals[s]; gtw += siteWeightedTotals[s]; });
  return {subsidiary: subName, rangeStart: fmtDate(rStart), rangeEnd: fmtDate(rEnd), sites: out, grandTotal: gt, grandTotalWeighted: Math.round(gtw*10)/10};
}
