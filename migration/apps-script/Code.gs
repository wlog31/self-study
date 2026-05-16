const APP_CONFIG = {
  timezone: "Asia/Seoul",
  monthFilePrefix: "SELF_STUDY_LOG",
  indexFileName: "SELF_STUDY_MONTH_INDEX",
  indexSheetName: "MONTH_INDEX",
  monthInfoSheetName: "_MONTH_INFO",
  dailySheetNameFormat: "MM-dd",
  maxReadRangeDays: 62,
  dailyHeader: [
    "created_at",
    "record_date",
    "actor_email",
    "actor_name",
    "student_id",
    "student_name",
    "subject",
    "duration_minutes",
    "memo",
    "client_timestamp",
    "client_tag",
  ],
};

function doGet(e) {
  return route_(e || {});
}

function doPost(e) {
  return route_(e || {});
}

function route_(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "health";
    switch (action) {
      case "health":
        return json_({
          ok: true,
          timezone: APP_CONFIG.timezone,
          now: formatDateTime_(new Date()),
        });
      case "appendLog":
        return json_(appendLog_(e));
      case "getLogs":
        return json_(getLogs_(e));
      case "getMonthFile":
        return json_(getMonthFile_(e));
      default:
        return json_({ ok: false, error: "Unknown action: " + action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function appendLog_(e) {
  const payload = parsePayload_(e);
  assertToken_(payload.token);

  const recordDate = parseIsoDate_(payload.date) || new Date();
  const subject = clean_(payload.subject);
  const duration = Number(payload.duration);
  if (!subject) {
    throw new Error("subject is required");
  }
  if (!isFinite(duration) || duration <= 0 || duration > 1440) {
    throw new Error("duration must be 1..1440");
  }

  const ss = getOrCreateMonthlySpreadsheet_(recordDate);
  const sheet = getOrCreateDailySheet_(ss, recordDate);
  const actorEmail = Session.getActiveUser().getEmail() || "";

  const row = [
    formatDateTime_(new Date()),
    formatDate_(recordDate),
    actorEmail,
    clean_(payload.actorName),
    clean_(payload.studentId),
    clean_(payload.studentName),
    subject,
    duration,
    clean_(payload.memo),
    clean_(payload.clientTimestamp),
    clean_(payload.clientTag),
  ];
  sheet.appendRow(row);

  return {
    ok: true,
    monthKey: monthKey_(recordDate),
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    sheetName: sheet.getName(),
    totalRows: sheet.getLastRow() - 1,
  };
}

function getLogs_(e) {
  const q = e.parameter || {};
  assertToken_(q.token || "");

  const from = parseIsoDate_(q.from) || firstDayOfCurrentMonth_();
  const to = parseIsoDate_(q.to) || new Date();
  if (from.getTime() > to.getTime()) {
    throw new Error("from must be <= to");
  }

  const dayCount = Math.floor((to.getTime() - from.getTime()) / 86400000) + 1;
  if (dayCount > APP_CONFIG.maxReadRangeDays) {
    throw new Error("Date range too large. max days: " + APP_CONFIG.maxReadRangeDays);
  }

  const studentIdFilter = clean_(q.studentId);
  const rows = [];
  const cursor = new Date(from.getTime());

  while (cursor.getTime() <= to.getTime()) {
    const ss = getMonthlySpreadsheetIfExists_(cursor);
    if (ss) {
      const sheetName = daySheetName_(cursor);
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, APP_CONFIG.dailyHeader.length).getValues();
        for (let i = 0; i < values.length; i += 1) {
          const row = rowToObject_(values[i]);
          if (studentIdFilter && row.studentId !== studentIdFilter) {
            continue;
          }
          rows.push(row);
        }
      }
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  rows.sort(function(a, b) {
    if (a.createdAt < b.createdAt) return 1;
    if (a.createdAt > b.createdAt) return -1;
    return 0;
  });

  return {
    ok: true,
    from: formatDate_(from),
    to: formatDate_(to),
    rows: rows,
  };
}

function getMonthFile_(e) {
  const q = e.parameter || {};
  assertToken_(q.token || "");

  const targetDate = parseIsoDate_(q.date) || new Date();
  const ss = getOrCreateMonthlySpreadsheet_(targetDate);
  return {
    ok: true,
    monthKey: monthKey_(targetDate),
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
  };
}

function getOrCreateMonthlySpreadsheet_(targetDate) {
  const key = monthKey_(targetDate);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const existing = openSpreadsheetByMonthKey_(key);
    if (existing) return existing;

    const ss = SpreadsheetApp.create(APP_CONFIG.monthFilePrefix + "_" + key);
    initializeMonthlySpreadsheet_(ss, key);
    moveFileToRootFolderIfConfigured_(ss.getId());
    upsertMonthIndex_(key, ss);
    return ss;
  } finally {
    lock.releaseLock();
  }
}

function getMonthlySpreadsheetIfExists_(targetDate) {
  const key = monthKey_(targetDate);
  return openSpreadsheetByMonthKey_(key);
}

function openSpreadsheetByMonthKey_(key) {
  const indexSheet = getOrCreateIndexSheet_();
  const values = indexSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    const monthKey = String(values[i][0] || "");
    const spreadsheetId = String(values[i][1] || "");
    if (monthKey !== key || !spreadsheetId) continue;
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (err) {
      return null;
    }
  }
  return null;
}

function upsertMonthIndex_(key, ss) {
  const indexSheet = getOrCreateIndexSheet_();
  const values = indexSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || "") === key) {
      indexSheet.getRange(i + 1, 2, 1, 3).setValues([[ss.getId(), ss.getUrl(), formatDateTime_(new Date())]]);
      return;
    }
  }
  indexSheet.appendRow([key, ss.getId(), ss.getUrl(), formatDateTime_(new Date())]);
}

function initializeMonthlySpreadsheet_(ss, key) {
  const sheets = ss.getSheets();
  if (sheets.length === 1 && sheets[0].getName() === "Sheet1") {
    sheets[0].setName(APP_CONFIG.monthInfoSheetName);
  }

  let info = ss.getSheetByName(APP_CONFIG.monthInfoSheetName);
  if (!info) {
    info = ss.insertSheet(APP_CONFIG.monthInfoSheetName, 0);
  }
  info.clear();
  info.getRange(1, 1, 4, 2).setValues([
    ["month_key", key],
    ["created_at", formatDateTime_(new Date())],
    ["timezone", APP_CONFIG.timezone],
    ["note", "Daily sheets are created on first write of each date."],
  ]);
  info.setFrozenRows(1);
}

function getOrCreateDailySheet_(ss, recordDate) {
  const sheetName = daySheetName_(recordDate);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, APP_CONFIG.dailyHeader.length).setValues([APP_CONFIG.dailyHeader]);
    sheet.setFrozenRows(1);
  } else {
    const firstRow = sheet.getRange(1, 1, 1, APP_CONFIG.dailyHeader.length).getValues()[0];
    if (String(firstRow[0] || "") !== APP_CONFIG.dailyHeader[0]) {
      sheet.insertRows(1);
      sheet.getRange(1, 1, 1, APP_CONFIG.dailyHeader.length).setValues([APP_CONFIG.dailyHeader]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function getOrCreateIndexSheet_() {
  const props = PropertiesService.getScriptProperties();
  let indexSpreadsheetId = props.getProperty("MONTH_INDEX_SPREADSHEET_ID");
  let ss = null;

  if (indexSpreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(indexSpreadsheetId);
    } catch (err) {
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create(APP_CONFIG.indexFileName);
    props.setProperty("MONTH_INDEX_SPREADSHEET_ID", ss.getId());
    moveFileToRootFolderIfConfigured_(ss.getId());
  }

  let sheet = ss.getSheetByName(APP_CONFIG.indexSheetName);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(APP_CONFIG.indexSheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["month_key", "spreadsheet_id", "spreadsheet_url", "updated_at"]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function moveFileToRootFolderIfConfigured_(fileId) {
  const folderId = PropertiesService.getScriptProperties().getProperty("ROOT_FOLDER_ID");
  if (!folderId) return;

  try {
    const file = DriveApp.getFileById(fileId);
    const folder = DriveApp.getFolderById(folderId);
    file.moveTo(folder);
  } catch (err) {
    throw new Error("ROOT_FOLDER_ID is invalid or inaccessible: " + err.message);
  }
}

function parsePayload_(e) {
  if (!e || !e.parameter) return {};
  if (!e.parameter.payload) return {};
  try {
    return JSON.parse(e.parameter.payload);
  } catch (err) {
    throw new Error("Invalid payload JSON");
  }
}

function rowToObject_(row) {
  return {
    createdAt: String(row[0] || ""),
    recordDate: String(row[1] || ""),
    actorEmail: String(row[2] || ""),
    actorName: String(row[3] || ""),
    studentId: String(row[4] || ""),
    studentName: String(row[5] || ""),
    subject: String(row[6] || ""),
    duration: Number(row[7] || 0),
    memo: String(row[8] || ""),
    clientTimestamp: String(row[9] || ""),
    clientTag: String(row[10] || ""),
  };
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(value) {
  return String(value || "").trim();
}

function parseIsoDate_(value) {
  const s = clean_(value);
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
  if (isNaN(d.getTime())) return null;
  return d;
}

function firstDayOfCurrentMonth_() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
}

function daySheetName_(d) {
  return Utilities.formatDate(d, APP_CONFIG.timezone, APP_CONFIG.dailySheetNameFormat);
}

function monthKey_(d) {
  return Utilities.formatDate(d, APP_CONFIG.timezone, "yyyy-MM");
}

function formatDate_(d) {
  return Utilities.formatDate(d, APP_CONFIG.timezone, "yyyy-MM-dd");
}

function formatDateTime_(d) {
  return Utilities.formatDate(d, APP_CONFIG.timezone, "yyyy-MM-dd HH:mm:ss");
}

function assertToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty("WRITE_TOKEN");
  if (!expected) return;
  if (clean_(token) !== clean_(expected)) {
    throw new Error("Unauthorized token");
  }
}
