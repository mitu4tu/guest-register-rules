'use strict';
try {
var T = JSON.parse(ctx.files.tabs);
var tabNames = Object.keys(T.tabs);
var csvRows = parseCSV(ctx.files.rules);
var rules = Object.create(null);
csvRows.slice(1).forEach(function (r) { var k = r[0] + '.' + r[1]; if (!(k in rules)) rules[k] = r[2]; });
var managed = {};
tabNames.slice().sort().forEach(function (name) {
  managed[name] = [];
  T.tabs[name].cells.forEach(function (rect) { each(rect[0], function (a) { managed[name].push(a); }); });
  managed[name].sort();
});

function parseCSV(text) {
  var rows = [], row = [], field = '', quoted = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    if (ch === '"') {
      if (quoted && text[i + 1] === '"') { field += '"'; i++; }
      else quoted = !quoted;
    } else if (!quoted && (ch === ',' || ch === '\r' || ch === '\n')) {
      row.push(field); field = '';
      if (ch !== ',') { rows.push(row); row = []; if (ch === '\r' && text[i + 1] === '\n') i++; }
    } else field += ch;
  }
  if (quoted) throw new Error('rules_parse: unclosed quote');
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows.filter(function (r) { return r.length > 1 || r[0] !== ''; }).map(function (r) { return [r[0] || '', r[1] || '', r[2] || '']; });
}
function column(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
function letters(n) { var s = ''; for (; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; }
function bounds(a, max) {
  var m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d*)?)?$/.exec(a);
  if (!m) throw new Error('rebuild_tab: invalid range');
  return { r: +m[2], c: column(m[1]), er: m[3] ? +m[4] || max || Infinity : +m[2], ec: column(m[3] || m[1]) };
}
function each(a, fn, max) {
  var b = bounds(a, max);
  for (var r = b.r; r <= b.er; r++) for (var c = b.c; c <= b.ec; c++) fn(letters(c) + r, r - b.r, c - b.c);
}
function kept(name, a) {
  var p = bounds(a);
  return T.tabs[name].keep.some(function (k) { var b = bounds(k); return p.r >= b.r && p.r <= b.er && p.c >= b.c && p.c <= b.ec; });
}
function get(c, k, fallback) { var v = c.props.getProperty(k); return v === null ? fallback : JSON.parse(v); }
function put(c, k, v) { var s = JSON.stringify(v); if (c.props.getProperty(k) !== s) c.props.setProperty(k, s); }
function del(c, k) { if (c.props.getProperty(k) !== null) c.props.deleteProperty(k); }
function textValue(s) { return s === '' ? '' : "'" + s; }
function value(s, formula) {
  if (typeof s !== 'string' || s === '') return s;
  if (formula && s[0] === '=') return s;
  return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : textValue(s);
}
function equal(a, b) { return a && b && a.getTime && b.getTime ? +a === +b : a === b; }
function literal(s) { var v = value(s, false); return typeof v === 'string' && v[0] === "'" ? v.slice(1) : v; }
function sheet(c, name) { return c.ss.getSheetByName(name); }
function message(key, vars) { return T.strings[key].replace(/\{(\w+)\}/g, function (_, k) { return vars[k]; }); }
function target(s) { var i = s.lastIndexOf('!'); return [s.slice(0, i).replace(/^'(.*)'$/, '$1').replace(/''/g, "'"), s.slice(i + 1)]; }
function tallyName(c) {
  var tallySheet = sheet(c, T.tally.tab);
  var tally = tallySheet ? String(tallySheet.getRange(T.tally.cell).getValue()).trim() : '';
  return tally || T.tally.default;
}
function preflight(c) {
  if (tabNames.indexOf(tallyName(c)) >= 0) return { ok: false, reason: 'tally-name' };
  var previous = get(c, 'managed', managed), reason = '';
  tabNames.some(function (name) {
    var sh = sheet(c, name); if (!sh) return false;
    return managed[name].some(function (a) {
      if ((previous[name] || []).indexOf(a) >= 0) return false;
      var r = sh.getRange(a);
      if (r.getValue() !== '' || r.getFormula() !== '') { reason = 'conflict:' + name + '!' + a; return true; }
      return false;
    });
  });
  return { ok: !reason, reason: reason };
}
function snapshotInputs(c) {
  var snap = { v: 1, cells: {} };
  tabNames.forEach(function (name) {
    var sh = sheet(c, name), found = Object.create(null); snap.cells[name] = [];
    if (!sh) return;
    T.tabs[name].inputs.forEach(function (a) {
      var rg = sh.getRange(a), vals = rg.getValues(), forms = rg.getFormulas();
      each(a, function (cell, i, j) {
        var v = vals[i][j], f = forms[i][j];
        if (found[cell] || (v === '' && !f)) return;
        found[cell] = true;
        var type = f ? 'f' : v && v.getTime ? 'd' : typeof v === 'number' ? 'n' : typeof v === 'boolean' ? 'b' : 's';
        snap.cells[name].push([cell, type, f || (type === 'd' ? +v : v)]);
      }, sh.getMaxRows());
    });
  });
  return snap;
}
function restoreInputs(c, snap) {
  if (snap == null) return;
  var tally = tallyName(c);
  tabNames.forEach(function (name) {
    if (name === tally) return;
    var sh = sheet(c, name); if (!sh) return;
    var entries = Object.create(null);
    (snap.cells[name] || []).forEach(function (v) { entries[v[0]] = v; });
    T.tabs[name].inputs.forEach(function (a) {
      var rg = sh.getRange(a), formats = rg.getNumberFormats(), b = bounds(a, sh.getMaxRows()), rows = [];
      for (var r = b.r; r <= b.er; r++) {
        var row = [];
        for (var col = b.c; col <= b.ec; col++) {
          var entry = entries[letters(col) + r];
          // A typed snapshot must preserve numeric-looking strings as strings.
          row.push(!entry ? '' : entry[1] === 'd' ? new Date(entry[2]) : entry[1] === 's' && formats[r - b.r][col - b.c] !== '@' ? textValue(entry[2]) : entry[2]);
        }
        rows.push(row);
      }
      rg.setValues(rows);
    });
  });
}
function applyRules(c) {
  if (T.names.rules === tallyName(c)) return;
  var sh = sheet(c, T.names.rules) || c.ss.insertSheet(T.names.rules), last = sh.getLastRow();
  if (csvRows.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), csvRows.length - sh.getMaxRows());
  if (csvRows.length) sh.getRange(1, 1, csvRows.length, 3).setValues(csvRows.map(function (r) { return r.map(function (v) { return value(v, false); }); }));
  if (last > csvRows.length) sh.getRange(csvRows.length + 1, 1, last - csvRows.length, 3).clearContent();
}
function rebuild(c) {
  var key = sha(c.files.tabs + '\n' + c.files.logic), cursor = get(c, 'rb', null);
  var step = cursor && cursor.key === key ? cursor.step : 0;
  var previous = get(c, 'managed', managed);
  var tally = tallyName(c), writable = tabNames.filter(function (name) { return name !== tally; });
  var steps = [
    function () {
      var union = {};
      Object.keys(previous).concat(tabNames).sort().forEach(function (name) {
        if (union[name]) return;
        union[name] = (previous[name] || []).concat(managed[name] || []).filter(function (v, i, all) { return all.indexOf(v) === i; }).sort();
      });
      put(c, 'managed', union);
    },
    function () { writable.forEach(function (name) {
      var sh = sheet(c, name) || c.ss.insertSheet(name), columns = sh.getMaxColumns();
      var required = T.tabs[name].minColumns || T.minColumns || 26;
      if (columns < required) sh.insertColumnsAfter(columns, required - columns);
    }); },
    function () {
      writable.forEach(function (name) {
        var sh = sheet(c, name);
        T.tabs[name].cells.forEach(function (rect) {
          var rg = sh.getRange(rect[0]), vals = rg.getValues(), forms = rg.getFormulas();
          var matches = rect[1].every(function (row, i) { return row.every(function (v, j) { return v[0] === '=' ? forms[i][j] === v : !forms[i][j] && equal(vals[i][j], literal(v)); }); });
          if (!matches) rg.setValues(rect[1].map(function (row) { return row.map(function (v) { return value(v, true); }); }));
        });
      });
    },
    function () {
      // Only current managed tabs are writable, even if an old set names others.
      writable.forEach(function (name) {
        (previous[name] || []).forEach(function (a) {
          if (managed[name].indexOf(a) >= 0 || kept(name, a)) return;
          var r = sheet(c, name).getRange(a); if (r.getValue() !== '' || r.getFormula()) r.setValue('');
        });
      });
    },
    function () {
      writable.forEach(function (name) {
        var sh = sheet(c, name), def = T.tabs[name];
        def.formats.forEach(function (f) {
          var r = sh.getRange(f.range);
          if (f.numberFormat) r.setNumberFormat(f.numberFormat);
          if (f.bold) r.setFontWeight('bold');
          if (f.background) r.setBackground(f.background);
          if (f.fontColor) r.setFontColor(f.fontColor);
        });
        sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
        def.validations.forEach(function (v) {
          var b = SpreadsheetApp.newDataValidation();
          if (v.list) b.requireValueInList(v.list, true).setAllowInvalid(false);
          else if (v.checkbox) b.requireCheckbox();
          else b.requireValueInRange(sh.getRange(v.inRange), true).setAllowInvalid(false);
          sh.getRange(v.range).setDataValidation(b.build());
        });
        sh.setConditionalFormatRules(def.conditional.map(function (rule) {
          var b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(rule.formula);
          if (rule.background) b.setBackground(rule.background);
          if (rule.fontColor) b.setFontColor(rule.fontColor);
          return b.setRanges([sh.getRange(rule.range)]).build();
        }));
        Object.keys(def.widths).forEach(function (col) { sh.setColumnWidth(column(col), def.widths[col]); });
        sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
        sh.protect().setDescription(T.protectDescription).setWarningOnly(true).setUnprotectedRanges(def.unprotected.map(function (a) { return sh.getRange(a); }));
      });
    },
    function () {
      var slots = [], active = null;
      c.ss.getSheets().forEach(function (sh, i) { if (writable.indexOf(sh.getName()) >= 0) slots.push(i + 1); });
      T.order.filter(function (name) { return name !== tally; }).forEach(function (name, i) {
        var wanted = sheet(c, name), position = wanted.getIndex(), displaced = c.ss.getSheets()[slots[i] - 1];
        if (displaced.getName() === name) return;
        if (!active) active = c.ss.getActiveSheet();
        // Swap managed sheets using only managed-sheet move operations.
        c.ss.setActiveSheet(wanted); c.ss.moveActiveSheet(slots[i]);
        c.ss.setActiveSheet(displaced); c.ss.moveActiveSheet(position);
      });
      if (active) c.ss.setActiveSheet(active);
      // Show first so Sheets never has to hide its last visible tab.
      T.visible.forEach(function (name) { if (name !== tally) sheet(c, name).showSheet(); });
      writable.forEach(function (name) { if (T.visible.indexOf(name) < 0) sheet(c, name).hideSheet(); });
    },
    function () { Object.keys(T.named).forEach(function (name) { var t = target(T.named[name]); if (t[0] !== tally) c.ss.setNamedRange(name, sheet(c, t[0]).getRange(t[1])); }); },
    function () { put(c, 'managed', managed); }
  ];
  for (; step < steps.length; step++) {
    if (Date.now() > c.deadlineMs - 20000) { put(c, 'rb', { key: key, step: step }); return { done: false }; }
    steps[step]();
    put(c, 'rb', { key: key, step: step + 1 });
  }
  del(c, 'rb'); return { done: true };
}
function selfCheck(c) {
  var problems = [], tally = tallyName(c);
  SpreadsheetApp.flush();
  tabNames.forEach(function (name) {
    if (name === tally) return;
    var sh = sheet(c, name); if (!sh) { problems.push('missing:' + name); return; }
    T.tabs[name].cells.forEach(function (rect) {
      var rg = sh.getRange(rect[0]), vals = rg.getValues(), forms = rg.getFormulas(), displays = rg.getDisplayValues();
      each(rect[0], function (a, i, j) {
        var v = rect[1][i][j];
        if (v[0] === '=') { if (!forms[i][j] || /^(#NAME\?|#ERROR!)/.test(displays[i][j])) problems.push('formula:' + name + '!' + a); }
        else if (forms[i][j] || !equal(vals[i][j], literal(v))) problems.push('literal:' + name + '!' + a);
      });
    });
  });
  Object.keys(T.named).forEach(function (name) {
    var actual = c.ss.getRangeByName(name), t = target(T.named[name]), sh = sheet(c, t[0]);
    if (t[0] === tally) return;
    if (!actual || !sh || actual.getSheet().getName() !== t[0] || actual.getA1Notation() !== sh.getRange(t[1]).getA1Notation()) problems.push('named:' + name);
  });
  var sh = sheet(c, T.names.rules);
  if (sh && T.names.rules !== tally) {
    var size = Math.max(csvRows.length, sh.getLastRow());
    if (size) {
      var rg = sh.getRange(1, 1, size, 3), rows = rg.getValues(), forms = rg.getFormulas();
      rows.forEach(function (row, i) { row.forEach(function (v, j) { if (forms[i][j] || !equal(v, literal((csvRows[i] || ['', '', ''])[j]))) problems.push('rules:' + (i + 1) + ':' + (j + 1)); }); });
    }
  }
  return { ok: !problems.length, problems: problems };
}
function sha(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function date(c, ms, pattern) { return Utilities.formatDate(new Date(ms), c.ss.getSpreadsheetTimeZone(), pattern); }
function licenseDay(ms) { return Utilities.formatDate(new Date(ms), 'Asia/Tokyo', 'yyyy-MM-dd'); }
function licenseRows(text) {
  return text.split(/\r?\n/).map(function (line) {
    return line.split(',').map(function (field) { return field.trim().replace(/^"+|"+$/g, '').trim(); });
  }).filter(function (row) {
    return row.length >= 3 && /^[a-f0-9]{64}$/i.test(row[0]) && (!row[2] || /^\d{4}-\d{2}-\d{2}$/.test(row[2]));
  });
}
function license(c) {
  var sh = sheet(c, T.names.settings), def = Number(rules['license.default_listings']) || 1;
  var key = String(sh.getRange(T.license.keyCell).getValue()).trim().toUpperCase();
  var n = def, msg = message('licNone', { def: def });
  if (key) {
    var rows = null, h = sha(key), prev = get(c, 'lic', null), today = licenseDay(+c.now);
    var seen = get(c, 'licSeen', []);
    // Remember positive checks independently of the current, revocable grant.
    if (prev && seen.indexOf(prev.h) < 0) { seen.push(prev.h); put(c, 'licSeen', seen); }
    var okDay = prev && (prev.okDay || (prev.okAt !== undefined ? licenseDay(prev.okAt) : ''));
    var graceEnd = okDay ? licenseDay(+new Date(okDay + 'T00:00:00Z') + T.license.graceDays * 86400000) : '';
    var grace = prev && prev.h === h && graceEnd && today <= graceEnd && (!prev.until || prev.until >= today);
    try {
      var res = rules['license.url'] ? UrlFetchApp.fetch(rules['license.url'], { muteHttpExceptions: true }) : null;
      if (res && res.getResponseCode() === 200) rows = licenseRows(res.getContentText());
    } catch (e) { rows = null; /* Grace handles transport and license parsing errors. */ }
    if (rows === null) {
      if (grace) return prev.n;
      msg = message(seen.indexOf(h) < 0 ? 'licUnreachableFirst' : 'licUnreachable', { def: def, days: T.license.graceDays });
    } else {
      var hit = rows.filter(function (r) { return r[0].toLowerCase() === h; })[0];
      var until = hit ? hit[2] : '';
      if (!hit) {
        if (grace) { n = prev.n; msg = message('licNotFoundGrace', { date: graceEnd.replace(/-/g, '/'), n: n }); }
        else msg = message('licNotFound', { def: def });
      } else if (until && until < today) { del(c, 'lic'); msg = message('licExpired', { until: until, def: def }); }
      else {
        n = parseInt(hit[1], 10); if (isNaN(n)) n = def;
        msg = message(until ? 'licOkUntil' : 'licOk', { n: n, until: until });
        // A fresh positive result refreshes the grace clock, even with count zero.
        c.props.setProperty('lic', JSON.stringify({ h: h, n: n, until: until, okDay: today }));
        if (seen.indexOf(h) < 0) { seen.push(h); put(c, 'licSeen', seen); }
      }
    }
  }
  sh.getRange(T.license.displayCell).setValue(textValue(msg)); return n;
}
function serial(d) { return Math.floor(+d / 86400000) + 25569; }
function ymd(ev, prefix) {
  var i = prefix ? ev.indexOf(prefix) : -1; if (i < 0) return '';
  var s = ev.substr(i + prefix.length, 8); if (!/^\d{8}$/.test(s)) return '';
  return (Date.UTC(+s.slice(0, 4), +s.slice(4, 6) - 1, +s.slice(6, 8)) - Date.UTC(1899, 11, 30)) / 86400000;
}
function parseIcal(text) {
  var begin = rules['ical.event_begin'] || 'BEGIN:VEVENT', end = rules['ical.event_end'] || 'END:VEVENT';
  var drop = rules['ical.drop_line_prefix'] || 'DTSTAMP', cp = new RegExp(rules['ical.code_pattern'] || 'HM[A-Z0-9]+');
  text = text.replace(/\r\n?/g, '\n').replace(/\n[ \t]/g, '');
  return text.split(begin).slice(1).map(function (part) {
    var raw = (begin + part.split(end)[0] + end).split('\n').filter(function (line) { return line !== '' && line.indexOf(drop) !== 0; }).join('\n');
    var m = cp.exec(raw);
    return { code: m ? m[1] || m[0] : '', ci: ymd(raw, rules['ical.start_prefix'] || 'DTSTART;VALUE=DATE:'), co: ymd(raw, rules['ical.end_prefix'] || 'DTEND;VALUE=DATE:'), raw: raw };
  });
}
function merge(data, no, events, now) {
  var index = Object.create(null);
  data.forEach(function (d, i) { if (d[0] === no) index[d[6]] = i; });
  events.forEach(function (e) {
    if (index[e.raw] !== undefined) data[index[e.raw]][5] = now;
    else { data.push([no, e.code, e.ci, e.co, now, now, e.raw]); index[e.raw] = data.length - 1; }
  });
}
function readSaved(sh) { return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().filter(function (r) { return r[6] !== ''; }); }
function writeSaved(sh, rows) {
  var last = sh.getLastRow(), need = rows.length + 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows.map(function (r) { return r.map(function (v) { return typeof v === 'string' ? textValue(v) : v; }); }));
  if (last > rows.length + 1) sh.getRange(rows.length + 2, 1, last - rows.length - 1, 7).clearContent();
}
function notify(c, day) {
  // The data schema has no notify-cell field: locate its template formula.
  var address = null;
  T.tabs[T.names.notify].cells.some(function (rect) {
    each(rect[0], function (a, i, j) { if (!address && rect[1][i][j][0] === '=') address = a; }); return !!address;
  });
  if (!address) return;
  var body = String(sheet(c, T.names.notify).getRange(address).getDisplayValue()).trim(); if (!body) return;
  var h = sha(body), prev = get(c, 'mail', null); if (prev && prev.day === day && prev.h === h) return;
  try { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), T.strings.mailSubject, body + '\n\n' + T.strings.mailSheet + c.ss.getUrl()); }
  catch (error) { var failure = new Error('notify_send: operation failed'); failure.logicCode = 'notify_send'; throw failure; }
  put(c, 'mail', { day: day, h: h });
}
function dailyRun(c) {
  var limit = license(c), set = sheet(c, T.names.settings), sh = sheet(c, T.names.ical);
  var list = set.getRange(T.list.top, 1, T.list.rows, 8).getValues(), data = readSaved(sh);
  var day = date(c, +c.now, 'yyyy-MM-dd'), cursor = get(c, 'ical', null);
  var start = cursor && cursor.day === day ? cursor.next : 0;
  if (cursor && cursor.day !== day) del(c, 'ical');
  var status = [], seen = Object.create(null), used = 0;
  function write() {
    writeSaved(sh, data);
    if (status.length) set.getRange(T.list.top + start, 7, status.length, 2).setValues(status.map(function (r) { return r.map(function (v) { return typeof v === 'string' ? textValue(v) : v; }); }));
  }
  for (var i = 0; i < list.length; i++) {
    var row = list[i], url = String(row[2]).trim(), msg = '', last = row[7], fetch = false;
    if (!url) msg = '';
    else if (row[5] === true) msg = T.strings.stStopped;
    else if (seen[url]) msg = T.strings.stDuplicate;
    else { seen[url] = true; if (++used > limit) msg = message('stOverLimit', { limit: limit }); else fetch = true; }
    if (i < start) continue;
    if (fetch) {
      if (Date.now() + 60000 > c.deadlineMs) { write(); put(c, 'ical', { day: day, next: i }); return { done: false }; }
      var res = null;
      try { res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); } catch (e) { msg = T.strings.stFetchError; }
      if (res) {
        if (res.getResponseCode() !== 200) msg = message('stHttp', { code: res.getResponseCode() });
        else if (res.getContentText().indexOf('BEGIN:VCALENDAR') < 0) msg = T.strings.stNotIcal;
        else {
          var events = parseIcal(res.getContentText()); merge(data, row[0], events, c.now);
          var codes = events.filter(function (e) { return e.code; }), noDate = codes.filter(function (e) { return e.ci === '' || e.co === ''; }).length;
          msg = message('stOk', { count: codes.length }) + (noDate ? message('stOkNoDate', { n: noDate }) : ''); last = c.now;
        }
      }
    }
    status.push([msg, last]);
  }
  var today = serial(c.now), days = Number(rules['ical.keep_days']) || 365;
  data = data.filter(function (d) { var co = d[3] && d[3].getTime ? serial(d[3]) : d[3]; var ref = typeof co === 'number' ? co : d[5] && d[5].getTime ? serial(d[5]) : null; return ref === null || ref >= today - days; });
  write(); del(c, 'ical'); SpreadsheetApp.flush(); notify(c, day); return { done: true };
}
function reparseAll(c) {
  var sh = sheet(c, T.names.ical), data = readSaved(sh);
  data.forEach(function (d) { var e = parseIcal(d[6])[0]; if (e) { d[1] = e.code; d[2] = e.ci; d[3] = e.co; } });
  writeSaved(sh, data); c.ss.toast(message('reparsed', { count: data.length }), T.strings.toastTitle, 10);
}
var api = { api: 1, layout: 1, preflight: preflight, snapshotInputs: snapshotInputs, restoreInputs: restoreInputs, applyRules: applyRules, rebuild: rebuild, selfCheck: selfCheck, dailyRun: dailyRun, reparseAll: reparseAll, commands: {} };
function copyState(c) {
  var id = String(c.installId || '');
  if (c.props.getProperty('install') === id) return;
  ['lic', 'licSeen', 'mail', 'ical', 'rb'].forEach(function (key) { del(c, key); });
  c.props.setProperty('install', id);
}
var codes = { preflight: 'preflight', snapshotInputs: 'snapshot', restoreInputs: 'restore', applyRules: 'rules_parse', rebuild: 'rebuild_tab', selfCheck: 'self_check', dailyRun: 'ical_write', reparseAll: 'ical_write' };
Object.keys(codes).forEach(function (name) {
  var fn = api[name];
  api[name] = function (c, arg) {
    copyState(c);
    try { return fn(c, arg); }
    catch (error) {
      // A service read also fences fake execution termination in local tests.
      c.props.getProperty('install');
      throw new Error((name === 'dailyRun' && error && error.logicCode === 'notify_send' ? 'notify_send' : codes[name]) + ': operation failed');
    }
  };
});
return api;
} catch (error) {
  var code = /^([a-z0-9_]{1,32}):/.exec(String(error && error.message || ''));
  throw new Error((code ? code[1] : 'rules_parse') + ': invalid configuration');
}
