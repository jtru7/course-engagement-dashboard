// ============================================================
// CONFIGURATION — paste your Google Sheet ID below, then
// deploy as a Web App (Execute as: Me, Access: Anyone).
// ============================================================
var SPREADSHEET_ID = '1kgFYR_9HHlGU2E-GLfqASzJGzPO6Z-wxzCSM8pwziPQ';

var SHEETS = {
  sections:   { name: 'sections',   headers: ['id', 'name', 'created_at'] },
  sessions:   { name: 'sessions',   headers: ['id', 'section_id', 'filename', 'date', 'title', 'duration_sec'] },
  attendance: { name: 'attendance', headers: ['session_id', 'section_id', 'name', 'login', 'email', 'duration_sec'] },
  canvas:     { name: 'canvas',     headers: ['section_id', 'student_name', 'login', 'completed', 'total', 'updated_at'] }
};

// ============================================================
// HTTP HANDLERS
// ============================================================

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var result;
  try {
    switch (params.action) {
      case 'getSections': result = getSections(); break;
      case 'getSection':  result = getSectionData(params.sectionId); break;
      default: result = { error: 'Unknown action: ' + params.action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return jsonOut(result);
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut({ error: 'Invalid request body' }); }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (err) { return jsonOut({ error: 'Server busy — try again in a moment.' }); }

  var result;
  try {
    switch (body.action) {
      case 'createSection': result = createSection(body.name); break;
      case 'deleteSection': result = deleteSection(body.sectionId); break;
      case 'addSession':    result = addSession(body.sectionId, body.session); break;
      case 'replaceCanvas': result = replaceCanvas(body.sectionId, body.students); break;
      case 'resetSection':  result = resetSection(body.sectionId); break;
      case 'canvasProxy':   result = canvasProxy(body.path, body.pat); break;
      default: result = { error: 'Unknown action: ' + body.action };
    }
  } catch (err) {
    result = { error: err.message };
  } finally {
    lock.releaseLock();
  }
  return jsonOut(result);
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SHEET HELPERS
// ============================================================

function openSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(key) {
  var cfg = SHEETS[key];
  var ss = openSS();
  var sheet = ss.getSheetByName(cfg.name);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.name);
    sheet.appendRow(cfg.headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetRows(key) {
  var sheet = getSheet(key);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow(key, obj) {
  getSheet(key).appendRow(SHEETS[key].headers.map(function(h) {
    return obj[h] !== undefined ? obj[h] : '';
  }));
}

function appendBulk(key, objs) {
  if (!objs || objs.length === 0) return;
  var sheet = getSheet(key);
  var headers = SHEETS[key].headers;
  var values = objs.map(function(obj) {
    return headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
}

function deleteRowsWhere(key, col, val) {
  var sheet = getSheet(key);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var colIdx = data[0].indexOf(col);
  if (colIdx < 0) return;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIdx]) === String(val)) sheet.deleteRow(i + 1);
  }
}

function uid() { return Utilities.getUuid(); }

// ============================================================
// SECTIONS
// ============================================================

function getSections() {
  return { sections: sheetRows('sections') };
}

function getSectionData(sectionId) {
  if (!sectionId) throw new Error('sectionId required');

  var sec = sheetRows('sections').filter(function(s) { return s.id === sectionId; })[0];
  if (!sec) throw new Error('Section not found');

  var sessions   = sheetRows('sessions').filter(function(s)  { return s.section_id === sectionId; });
  var attendance = sheetRows('attendance').filter(function(a) { return a.section_id === sectionId; });
  var canvasRows = sheetRows('canvas').filter(function(c)    { return c.section_id === sectionId; });

  var bySession = {};
  attendance.forEach(function(a) {
    if (!bySession[a.session_id]) bySession[a.session_id] = [];
    bySession[a.session_id].push({
      name: a.name, login: a.login, email: a.email,
      durationSec: Number(a.duration_sec)
    });
  });

  return {
    section: sec,
    sessions: sessions.map(function(s) {
      return {
        id: s.id, filename: s.filename, date: s.date,
        title: s.title, durationSec: Number(s.duration_sec),
        participants: bySession[s.id] || []
      };
    }),
    canvas: canvasRows.length ? {
      students: canvasRows.map(function(c) {
        return {
          name: c.student_name, login: c.login,
          completedAssignments: Number(c.completed),
          totalAssignments: Number(c.total)
        };
      }),
      updatedAt: canvasRows[0].updated_at
    } : null
  };
}

function createSection(name) {
  if (!name || !String(name).trim()) throw new Error('Section name is required');
  var rec = { id: uid(), name: String(name).trim(), created_at: new Date().toISOString() };
  appendRow('sections', rec);
  return rec;
}

function deleteSection(sectionId) {
  if (!sectionId) throw new Error('sectionId required');
  deleteRowsWhere('attendance', 'section_id', sectionId);
  deleteRowsWhere('sessions',   'section_id', sectionId);
  deleteRowsWhere('canvas',     'section_id', sectionId);
  deleteRowsWhere('sections',   'id',         sectionId);
  return { ok: true };
}

function resetSection(sectionId) {
  if (!sectionId) throw new Error('sectionId required');
  deleteRowsWhere('attendance', 'section_id', sectionId);
  deleteRowsWhere('sessions',   'section_id', sectionId);
  deleteRowsWhere('canvas',     'section_id', sectionId);
  return { ok: true };
}

// ============================================================
// SESSIONS
// ============================================================

function addSession(sectionId, session) {
  if (!sectionId || !session) throw new Error('sectionId and session required');

  var dup = sheetRows('sessions').some(function(s) {
    return s.section_id === sectionId && s.filename === session.filename;
  });
  if (dup) return { ok: false, duplicate: true };

  var sessionId = uid();
  appendRow('sessions', {
    id: sessionId,       section_id: sectionId,
    filename: session.filename, date: session.date,
    title: session.title, duration_sec: session.durationSec
  });
  appendBulk('attendance', (session.participants || []).map(function(p) {
    return {
      session_id: sessionId, section_id: sectionId,
      name: p.name, login: p.login, email: p.email, duration_sec: p.durationSec
    };
  }));
  return { ok: true, sessionId: sessionId };
}

// ============================================================
// CANVAS
// ============================================================

// ============================================================
// CANVAS PROXY — fetches Canvas API server-side (no CORS)
// ============================================================

function canvasProxy(path, pat) {
  if (!path || !pat) throw new Error('path and pat required');
  var CANVAS_BASE = 'https://byuird.instructure.com';
  var results = [];
  var url = CANVAS_BASE + path;
  while (url) {
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + pat },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 401) throw new Error('Invalid Canvas API key — please update it.');
    if (code !== 200) throw new Error('Canvas API error: HTTP ' + code);
    var page = JSON.parse(response.getContentText());
    results = results.concat(page);
    var headers = response.getHeaders();
    var linkHeader = '';
    for (var h in headers) {
      if (h.toLowerCase() === 'link') { linkHeader = String(headers[h]); break; }
    }
    var match = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
    url = match ? match[1] : null;
  }
  return { data: results };
}

function replaceCanvas(sectionId, students) {
  if (!sectionId) throw new Error('sectionId required');
  deleteRowsWhere('canvas', 'section_id', sectionId);
  var now = new Date().toISOString();
  appendBulk('canvas', (students || []).map(function(s) {
    return {
      section_id: sectionId, student_name: s.name, login: s.login,
      completed: s.completedAssignments, total: s.totalAssignments, updated_at: now
    };
  }));
  return { ok: true, count: (students || []).length };
}
