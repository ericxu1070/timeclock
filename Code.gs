/**
 * Restaurant Time Clock — Google Apps Script backend
 *
 * Deploy: Deploy → New deployment → Web app
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * IMPORTANT: set the script timezone to the restaurant's timezone
 * (Project Settings → Time zone) — all stored timestamps use it.
 *
 * Sheets expected in the bound spreadsheet:
 *   Punches:   Timestamp | Employee Name | Employee ID | Action | Date
 *   Employees: ID | Name | PIN Hash        (row with ID=0 is the admin PIN)
 *   PayPeriod: B1 = PAY_START, B2 = PAY_END, headers in row 3, data from row 4
 */

var SHEET_PUNCHES = "Punches";
var SHEET_EMPLOYEES = "Employees";
var SHEET_PAYPERIOD = "PayPeriod";

// ---------------------------------------------------------------- dispatch

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === "GET_STATE") return getState();
    return jsonOut({ status: "error", error: "unknown_action" });
  } catch (err) {
    return jsonOut({ status: "error", error: String(err) });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    switch (data.action) {
      case "PUNCH":
        if (!verifyPin(data.employeeId, data.pinHash))
          return jsonOut({ status: "error", error: "bad_pin" });
        return recordPunch(data);

      case "VERIFY_ADMIN":
        return verifyAdmin(data.adminPinHash)
          ? jsonOut({ status: "ok" })
          : jsonOut({ status: "error", error: "unauthorized" });

      case "ADD_EMPLOYEE":
        if (!verifyAdmin(data.adminPinHash))
          return jsonOut({ status: "error", error: "unauthorized" });
        return addEmployee(data);

      case "DELETE_EMPLOYEE":
        if (!verifyAdmin(data.adminPinHash))
          return jsonOut({ status: "error", error: "unauthorized" });
        return deleteEmployee(data);

      case "SET_EMPLOYEE_PIN":
        if (!verifyAdmin(data.adminPinHash))
          return jsonOut({ status: "error", error: "unauthorized" });
        return setEmployeePin(data);

      case "PURGE_PUNCHES":
        if (!verifyAdmin(data.adminPinHash))
          return jsonOut({ status: "error", error: "unauthorized" });
        return purgePunches(data);

      case "CHANGE_ADMIN_PIN":
        if (!verifyAdmin(data.adminPinHash))
          return jsonOut({ status: "error", error: "unauthorized" });
        return changeAdminPin(data);

      default:
        return jsonOut({ status: "error", error: "unknown_action" });
    }
  } catch (err) {
    return jsonOut({ status: "error", error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ---------------------------------------------------------------- helpers

function getSheet(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error("Missing sheet: " + name);
  return sheet;
}

function employeeRows() {
  return getSheet(SHEET_EMPLOYEES).getDataRange().getValues().slice(1);
}

function verifyPin(employeeId, pinHash) {
  if (!pinHash) return false;
  var rows = employeeRows();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(employeeId)) return rows[i][2] === pinHash;
  }
  return false;
}

function verifyAdmin(adminPinHash) {
  if (!adminPinHash) return false;
  var rows = employeeRows();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === "0") return rows[i][2] === adminPinHash;
  }
  return false;
}

function toDate(v) {
  return v instanceof Date ? v : new Date(v);
}

function fmtTimestamp(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function fmtDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// Decimal hours -> readable "Xh Ym" (e.g. 38.5 -> "38h 30m"),
// rounded to the nearest whole minute.
function hoursToHM(decimalHours) {
  var totalMin = Math.round(decimalHours * 60);
  var h = Math.floor(totalMin / 60);
  var m = totalMin % 60;
  return h + "h " + m + "m";
}

// ---------------------------------------------------------------- GET_STATE

function getState() {
  var empRows = employeeRows();
  var punchRows = getSheet(SHEET_PUNCHES).getDataRange().getValues().slice(1);

  // Punches are appended chronologically, so the last row seen per
  // employee is their most recent punch.
  var last = {};
  for (var i = 0; i < punchRows.length; i++) {
    var id = String(punchRows[i][2]);
    if (id === "") continue;
    last[id] = { action: String(punchRows[i][3]), ts: toDate(punchRows[i][0]) };
  }

  var employees = [];
  for (var j = 0; j < empRows.length; j++) {
    var empId = String(empRows[j][0]);
    if (empId === "" || empId === "0") continue; // skip blanks and the admin row
    var lp = last[empId];
    var emp = {
      id: Number(empId),
      name: String(empRows[j][1]),
      status: lp && lp.action === "IN" ? "IN" : "OUT",
    };
    if (emp.status === "IN") emp.clockedInAt = fmtTimestamp(lp.ts);
    employees.push(emp);
  }
  return jsonOut({ status: "ok", employees: employees });
}

// ---------------------------------------------------------------- PUNCH

function recordPunch(data) {
  if (data.punchType !== "IN" && data.punchType !== "OUT")
    return jsonOut({ status: "error", error: "invalid_punch_type" });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Server clock is the source of truth; the tablet's clock is ignored.
    var now = new Date();
    var ts = fmtTimestamp(now);
    getSheet(SHEET_PUNCHES).appendRow([
      ts,
      String(data.employeeName || ""),
      Number(data.employeeId),
      data.punchType,
      fmtDate(now),
    ]);
    return jsonOut({ status: "ok", timestamp: ts });
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- admin: employees

function addEmployee(data) {
  var name = String(data.name || "").trim();
  var pinHash = String(data.pinHash || "");
  if (!name || pinHash.length !== 64)
    return jsonOut({ status: "error", error: "invalid_input" });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rows = employeeRows();
    var maxId = 0;
    for (var i = 0; i < rows.length; i++) {
      var n = Number(rows[i][0]);
      if (!isNaN(n) && n > maxId) maxId = n;
    }
    var newId = maxId + 1;
    getSheet(SHEET_EMPLOYEES).appendRow([newId, name, pinHash]);
    return jsonOut({ status: "ok", employee: { id: newId, name: name, status: "OUT" } });
  } finally {
    lock.releaseLock();
  }
}

function deleteEmployee(data) {
  var idStr = String(data.employeeId);
  if (idStr === "0") return jsonOut({ status: "error", error: "cannot_delete_admin" });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var empSheet = getSheet(SHEET_EMPLOYEES);
    var rows = empSheet.getDataRange().getValues();
    var rowIndex = -1;
    var name = "";
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === idStr) {
        rowIndex = i + 1;
        name = String(rows[i][1]);
        break;
      }
    }
    if (rowIndex === -1) return jsonOut({ status: "error", error: "not_found" });

    // Close an open shift so their hours stay countable after deletion.
    var punchSheet = getSheet(SHEET_PUNCHES);
    var pRows = punchSheet.getDataRange().getValues().slice(1);
    var lastAction = null;
    for (var k = 0; k < pRows.length; k++) {
      if (String(pRows[k][2]) === idStr) lastAction = String(pRows[k][3]);
    }
    if (lastAction === "IN") {
      var now = new Date();
      punchSheet.appendRow([fmtTimestamp(now), name, Number(idStr), "OUT", fmtDate(now)]);
    }

    empSheet.deleteRow(rowIndex);
    return jsonOut({ status: "ok" });
  } finally {
    lock.releaseLock();
  }
}

function setEmployeePin(data) {
  var newHash = String(data.newPinHash || "");
  if (newHash.length !== 64) return jsonOut({ status: "error", error: "invalid_input" });
  if (String(data.employeeId) === "0")
    return jsonOut({ status: "error", error: "use_change_admin_pin" });

  var sheet = getSheet(SHEET_EMPLOYEES);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.employeeId)) {
      sheet.getRange(i + 1, 3).setValue(newHash);
      return jsonOut({ status: "ok" });
    }
  }
  return jsonOut({ status: "error", error: "not_found" });
}

function changeAdminPin(data) {
  var newHash = String(data.newAdminPinHash || "");
  if (newHash.length !== 64) return jsonOut({ status: "error", error: "invalid_input" });

  var sheet = getSheet(SHEET_EMPLOYEES);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === "0") {
      sheet.getRange(i + 1, 3).setValue(newHash);
      return jsonOut({ status: "ok" });
    }
  }
  return jsonOut({ status: "error", error: "admin_row_missing" });
}

// ---------------------------------------------------------------- admin: purge

function purgePunches(data) {
  var cutoff = new Date(data.before);
  if (isNaN(cutoff.getTime())) return jsonOut({ status: "error", error: "invalid_date" });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet(SHEET_PUNCHES);
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return jsonOut({ status: "ok", deleted: 0 });

    // Rewrite kept rows in one batch instead of deleting row-by-row
    // (deleteRow is one API call per row and crawls on large sheets).
    var kept = [];
    var deleted = 0;
    for (var i = 1; i < values.length; i++) {
      var d = toDate(values[i][4]); // column E = Date
      if (d < cutoff) deleted++;
      else kept.push(values[i]);
    }
    if (deleted > 0) {
      sheet.getRange(2, 1, values.length - 1, values[0].length).clearContent();
      if (kept.length > 0)
        sheet.getRange(2, 1, kept.length, values[0].length).setValues(kept);
    }
    return jsonOut({ status: "ok", deleted: deleted });
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- pay period

/**
 * Recomputes the PayPeriod summary. Run on a time-driven trigger
 * (Triggers → Add → updatePayPeriod → Time-driven → every 30 minutes).
 *
 * Overtime is computed PER 7-DAY WORKWEEK anchored to PAY_START (FLSA
 * rule), not on the pay-period total — a bi-weekly period with 35h/35h
 * has zero overtime. Each shift is attributed to the week its IN falls in.
 */
function updatePayPeriod() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var punches = ss.getSheetByName(SHEET_PUNCHES);
  var summary = ss.getSheetByName(SHEET_PAYPERIOD);

  var payStart = new Date(summary.getRange("B1").getValue());
  var payEnd = new Date(summary.getRange("B2").getValue());
  if (isNaN(payStart.getTime()) || isNaN(payEnd.getTime())) return;
  payEnd.setHours(23, 59, 59); // include the full end day

  var rows = punches.getDataRange().getValues().slice(1);

  var byEmployee = {};
  for (var i = 0; i < rows.length; i++) {
    var ts = toDate(rows[i][0]);
    if (ts < payStart || ts > payEnd) continue;
    var name = String(rows[i][1]);
    if (!byEmployee[name]) byEmployee[name] = [];
    byEmployee[name].push({ ts: ts, action: String(rows[i][3]) });
  }

  var MS_PER_WEEK = 7 * 24 * 3600000;
  var results = [];
  for (var name2 in byEmployee) {
    var list = byEmployee[name2];
    list.sort(function (a, b) {
      return a.ts - b.ts;
    });

    var weekHours = {};
    var openIn = null;
    for (var j = 0; j < list.length; j++) {
      var p = list[j];
      if (p.action === "IN") openIn = p.ts;
      if (p.action === "OUT" && openIn) {
        var hours = (p.ts - openIn) / 3600000;
        var week = Math.floor((openIn - payStart) / MS_PER_WEEK);
        weekHours[week] = (weekHours[week] || 0) + hours;
        openIn = null;
      }
    }

    var total = 0,
      regular = 0,
      overtime = 0;
    for (var w in weekHours) {
      var h = weekHours[w];
      total += h;
      regular += Math.min(h, 40);
      overtime += Math.max(0, h - 40);
    }

    var note = openIn ? "On the clock since " + fmtTimestamp(openIn) : "";
    // Columns: Name | Total Hours | Total (h:m) | Regular Hours | Regular (h:m)
    //        | Overtime Hours | Overtime (h:m) | Notes
    results.push([
      name2,
      total.toFixed(2),    hoursToHM(total),
      regular.toFixed(2),  hoursToHM(regular),
      overtime.toFixed(2), hoursToHM(overtime),
      note,
    ]);
  }

  results.sort(function (a, b) {
    return a[0].localeCompare(b[0]);
  });

  // Clear old data rows generously, then write the fresh set (8 columns).
  var lastRow = summary.getLastRow();
  var clearRows = Math.max(lastRow - 3, results.length, 1);
  summary.getRange(4, 1, clearRows, 8).clearContent();
  if (results.length > 0) summary.getRange(4, 1, results.length, 8).setValues(results);
}
