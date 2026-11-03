/**
 * EDYC Duties — Google Apps Script
 *
 * Adds an "EDYC Website" menu to the spreadsheet.
 * "Refresh Duties" prompts for credentials, calls the EDYC API,
 * and writes all allocated duties as rows into a "Duties" sheet.
 */

// ─── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('EDYC Website')
    .addItem('Refresh Duties', 'showDutiesLoginDialog')
    .addToUi();
}

// ─── Login dialog (HTML popup) ────────────────────────────────────────────────

function showDutiesLoginDialog() {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: Arial, sans-serif;
      padding: 20px;
      margin: 0;
      background: #f8f9fa;
    }
    h3 {
      margin: 0 0 16px;
      color: #1a73e8;
      font-size: 16px;
    }
    label {
      display: block;
      font-size: 12px;
      color: #555;
      margin-bottom: 4px;
      margin-top: 12px;
    }
    input {
      width: 100%;
      padding: 8px 10px;
      border: 1px solid #ccc;
      border-radius: 4px;
      font-size: 14px;
      box-sizing: border-box;
    }
    input:focus { outline: none; border-color: #1a73e8; }
    #error {
      color: #d93025;
      font-size: 12px;
      margin-top: 10px;
      display: none;
    }
    #buttons {
      margin-top: 20px;
      display: flex;
      justify-content: flex-end;
      gap: 8px;
    }
    button {
      padding: 8px 20px;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
    }
    #cancelBtn {
      background: #fff;
      border: 1px solid #ccc;
      color: #333;
    }
    #submitBtn {
      background: #1a73e8;
      color: #fff;
    }
    #submitBtn:disabled { background: #aaa; cursor: not-allowed; }
    #spinner { display: none; font-size: 12px; color: #555; margin-top: 8px; }
  </style>
</head>
<body>
  <h3>🔐 EDYC — Fetch Allocated Duties</h3>

  <label for="username">Username</label>
  <input type="text" id="username" placeholder="e.g. WeAreEDYC" autocomplete="username" />

  <label for="password">Password</label>
  <input type="password" id="password" placeholder="Password" autocomplete="current-password" />

  <div id="error"></div>
  <div id="spinner">⏳ Fetching duties, please wait…</div>

  <div id="buttons">
    <button id="cancelBtn" onclick="google.script.host.close()">Cancel</button>
    <button id="submitBtn" onclick="submit()">Refresh Duties</button>
  </div>

  <script>
    document.getElementById('username').focus();

    // Allow Enter key to submit
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') submit();
    });

    function submit() {
      const username = document.getElementById('username').value.trim();
      const password = document.getElementById('password').value;

      if (!username || !password) {
        showError('Please enter both username and password.');
        return;
      }

      document.getElementById('submitBtn').disabled = true;
      document.getElementById('error').style.display = 'none';
      document.getElementById('spinner').style.display = 'block';

      google.script.run
        .withSuccessHandler(function(result) {
          google.script.host.close();
        })
        .withFailureHandler(function(err) {
          showError(err.message || String(err));
          document.getElementById('submitBtn').disabled = false;
          document.getElementById('spinner').style.display = 'none';
        })
        .fetchAndWriteDuties(username, password);
    }

    function showError(msg) {
      const el = document.getElementById('error');
      el.textContent = '❌ ' + msg;
      el.style.display = 'block';
    }
  </script>
</body>
</html>
  `)
  .setWidth(360)
  .setHeight(280);

  SpreadsheetApp.getUi().showModalDialog(html, 'EDYC Duties');
}

// ─── Server-side: authenticate + fetch + write ────────────────────────────────

function fetchAndWriteDuties(username, password) {
  const API_BASE = 'https://production-api-production-a02e.up.railway.app';

  // 1. Authenticate
  const authRes = UrlFetchApp.fetch(API_BASE + '/api/auth/login', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ username: username, password: password }),
    headers: {
      'Origin':  'https://edyc.co.uk',
      'Referer': 'https://edyc.co.uk/',
    },
    muteHttpExceptions: true,
  });

  if (authRes.getResponseCode() !== 200) {
    throw new Error('Authentication failed (HTTP ' + authRes.getResponseCode() + '). Check your credentials.');
  }

  const authJson = JSON.parse(authRes.getContentText());
  const token =
    authJson.token                ||
    authJson.accessToken          ||
    authJson.access_token         ||
    (authJson.data && (authJson.data.accessToken || authJson.data.token));

  if (!token) {
    throw new Error('Authenticated but no token returned. Response: ' + authRes.getContentText().slice(0, 200));
  }

  // 2. Fetch duties
  const url = API_BASE + '/api/duties?limit=100000';
  const dutiesRes = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type':  'application/json',
      'Accept':        '*/*',
      'Origin':        'https://edyc.co.uk',
      'Referer':       'https://edyc.co.uk/',
    },
    muteHttpExceptions: true,
  });

  if (dutiesRes.getResponseCode() !== 200) {
    throw new Error('Failed to fetch duties (HTTP ' + dutiesRes.getResponseCode() + ').');
  }

  const dutiesJson = JSON.parse(dutiesRes.getContentText());

  // API shape: { success: true, data: [...] }
  const duties =
    (dutiesJson.data && Array.isArray(dutiesJson.data))   ? dutiesJson.data   :
    Array.isArray(dutiesJson)                              ? dutiesJson        :
    Array.isArray(dutiesJson.duties)                       ? dutiesJson.duties :
    Array.isArray(dutiesJson.items)                        ? dutiesJson.items  :
    null;

  if (!duties) {
    throw new Error('Unexpected API response shape: ' + dutiesRes.getContentText().slice(0, 200));
  }

  // 3. Write to spreadsheet
  writeDutiesToSheet(duties);
}

// ─── Write rows to sheet (roster layout matching duties.xlsx) ────────────────
//
// Layout mirrors duties.xlsx:
//   • One MAIN row per date  → Date | Day | Time | RO | RIB | Daisy | DHO | first bar shift
//   • CONTINUATION rows      → blank | blank | bar time | blank sailing cols | next bar shift
//
// Columns (fixed order per user request):
//   A: Date   B: Day   C: Bar Shift Time
//   D: Race Officer   E: Rib Driver   F: Daisy Driver   G: Duty House Officer
//   H: Bar Duty   I: Bar Duty Food

function writeDutiesToSheet(duties) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('Duties');

  if (!sheet) sheet = ss.insertSheet('Duties');

  sheet.clearContents();
  sheet.clearFormats();

  // ── Column definitions ────────────────────────────────────────────────────────

  // Duty-type columns in user-requested order
  // Keys MUST match API dutyType.name values (case-insensitive)
  const SAILING_COLS = [
    'Race Officer',
    'Rib Driver',
    'Daisy Driver',
    'Duty House Officer',
  ];
  const BAR_COLS = [
    'Bar Duty',
    'Bar Duty Food',
  ];
  const ALL_DUTY_COLS = SAILING_COLS.concat(BAR_COLS);

  // Headers: A–C are fixed, D onwards are duty types
  const headers  = ['Date', 'Day', 'Bar Shift Time'].concat(ALL_DUTY_COLS);
  const totalCols = headers.length; // = 9

  // Helper: case-insensitive column index (0-based within ALL_DUTY_COLS)
  function dutyColIndex(typeName) {
    const n = (typeName || '').toLowerCase().trim();
    for (var i = 0; i < ALL_DUTY_COLS.length; i++) {
      if (ALL_DUTY_COLS[i].toLowerCase() === n) return i;
    }
    return -1;
  }

  function isBarType(typeName) {
    const n = (typeName || '').toLowerCase().trim();
    return BAR_COLS.some(function(c) { return c.toLowerCase() === n; });
  }

  // ── Group duties by date ──────────────────────────────────────────────────────

  const byDate = {};
  duties.forEach(function(duty) {
    const raw = duty.dutyDate || '';
    const key = raw ? raw.slice(0, 10) : 'unknown';
    if (!byDate[key]) byDate[key] = { dateObj: raw ? new Date(raw) : null, duties: [] };
    byDate[key].duties.push(duty);
  });

  const sortedKeys = Object.keys(byDate).sort();

  // ── Build rows ────────────────────────────────────────────────────────────────

  const rows        = [headers];
  const mainRowNums = [];   // 1-based sheet row numbers of "main" (date) rows — used for styling

  sortedKeys.forEach(function(key) {
    const entry   = byDate[key];
    const dateObj = entry.dateObj;
    const dateLabel = dateObj ? formatDutyDate(dateObj) : key;
    const dayLabel  = dateObj
      ? dateObj.toLocaleDateString('en-GB', { weekday: 'long' })
      : '';

    // ── Separate sailing duties from bar duties ─────────────────────────────

    // sailingByCol[colIndex] = [name, name, ...]
    const sailingByCol = {};
    SAILING_COLS.forEach(function(_, i) { sailingByCol[i] = []; });

    // barSlots: array of { timeKey, startTime, endTime, byCol: {colIndex: [names]} }
    // sorted by startTime
    const barSlotMap = {};  // "HH:MM-HH:MM" → { startTime, endTime, byCol }

    // General sailing time (from non-bar duties)
    var generalSailingTime = '';

    entry.duties.forEach(function(d) {
      const typeName = (d.dutyType && d.dutyType.name) || '';
      const assignee = (d.assigneeName ||
        (d.member ? (d.member.firstName + ' ' + d.member.lastName).trim() : '')) || '';
      const colIdx = dutyColIndex(typeName);
      if (colIdx === -1) return;  // unknown duty type — skip

      if (isBarType(typeName)) {
        // Group bar duties by time slot
        const slotKey = (d.startTime || '') + '-' + (d.endTime || '');
        if (!barSlotMap[slotKey]) {
          barSlotMap[slotKey] = {
            startTime: d.startTime || '',
            endTime:   d.endTime   || '',
            byCol:     {},
          };
          BAR_COLS.forEach(function(_, bi) { barSlotMap[slotKey].byCol[SAILING_COLS.length + bi] = []; });
        }
        if (assignee) barSlotMap[slotKey].byCol[colIdx].push(assignee);
      } else {
        // Sailing duty
        const sailIdx = colIdx;  // 0–3
        if (sailingByCol[sailIdx] !== undefined && assignee) {
          sailingByCol[sailIdx].push(assignee);
        }
        if (!generalSailingTime && d.startTime && d.endTime) {
          generalSailingTime = d.startTime + '–' + d.endTime;
        }
      }
    });

    // Sort bar slots by start time string (HH:MM lexicographic works fine)
    const barSlots = Object.keys(barSlotMap)
      .sort(function(a, b) {
        return barSlotMap[a].startTime.localeCompare(barSlotMap[b].startTime);
      })
      .map(function(k) { return barSlotMap[k]; });

    const hasSailing = SAILING_COLS.some(function(_, i) { return sailingByCol[i].length > 0; });
    const hasBar     = barSlots.length > 0;

    // ── MAIN row ─────────────────────────────────────────────────────────────
    // Shows: Date, Day, first bar shift time (or general sailing time),
    //        all sailing duties, + first bar shift people (if any)

    const firstSlot    = hasBar ? barSlots[0] : null;
    const mainTimeLabel = firstSlot
      ? (firstSlot.startTime + '–' + firstSlot.endTime)
      : generalSailingTime;

    var mainRow = [dateLabel, dayLabel, mainTimeLabel];
    ALL_DUTY_COLS.forEach(function(_, colIdx) {
      if (colIdx < SAILING_COLS.length) {
        // Sailing column — put all names (stacked) on main row
        mainRow.push(sailingByCol[colIdx].join('\n'));
      } else {
        // Bar column — put first slot's people on main row
        if (firstSlot) {
          mainRow.push(firstSlot.byCol[colIdx].join(', '));
        } else {
          mainRow.push('');
        }
      }
    });
    rows.push(mainRow);
    mainRowNums.push(rows.length);  // 1-based (header=1, this is rows.length after push)

    // ── CONTINUATION rows (bar slots 2, 3, …) ────────────────────────────────
    for (var si = 1; si < barSlots.length; si++) {
      var slot = barSlots[si];
      var contTimeLabel = slot.startTime + '–' + slot.endTime;
      var contRow = ['', '', contTimeLabel];
      ALL_DUTY_COLS.forEach(function(_, colIdx) {
        if (colIdx < SAILING_COLS.length) {
          contRow.push('');  // sailing duties only on main row
        } else {
          contRow.push(slot.byCol[colIdx].join(', '));
        }
      });
      rows.push(contRow);
    }
  });

  // ── Write all rows at once ────────────────────────────────────────────────────

  sheet.getRange(1, 1, rows.length, totalCols).setValues(rows);

  // ── Styling ───────────────────────────────────────────────────────────────────

  const totalRows = rows.length;

  // Header — dark blue
  sheet.getRange(1, 1, 1, totalCols)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Data rows: default white background
  if (totalRows > 1) {
    sheet.getRange(2, 1, totalRows - 1, totalCols).setBackground('#ffffff');
  }

  // Main (date) rows — light blue background + bold Date & Day
  mainRowNums.forEach(function(r) {
    sheet.getRange(r, 1, 1, totalCols).setBackground('#e8f0fe');
    sheet.getRange(r, 1, 1, 2).setFontWeight('bold');  // cols A & B bold
  });

  // Centre-align the Time column (C) and duty columns (D–I) for all data rows
  if (totalRows > 1) {
    sheet.getRange(2, 3, totalRows - 1, totalCols - 2).setHorizontalAlignment('center');
  }

  // Wrap text so stacked names display properly
  sheet.getRange(1, 1, totalRows, totalCols).setWrap(true);

  // Set sensible column widths
  sheet.setColumnWidth(1, 100);   // A: Date
  sheet.setColumnWidth(2, 90);    // B: Day
  sheet.setColumnWidth(3, 105);   // C: Bar Shift Time
  // D–G: sailing duty columns
  for (var c = 4; c <= 7; c++) sheet.setColumnWidth(c, 140);
  // H–I: bar duty columns
  for (var c = 8; c <= 9; c++) sheet.setColumnWidth(c, 140);

  // Freeze header + first two columns
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // Activate the sheet
  ss.setActiveSheet(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    duties.length + ' duty record(s) across ' + sortedKeys.length + ' date(s).',
    '✅ EDYC Duties Refreshed',
    5
  );
}

// ─── Utility ──────────────────────────────────────────────────────────────────

function formatDutyDate(value) {
  if (!value) return '';
  try {
    return new Date(value).toLocaleDateString('en-GB', {
      day:   '2-digit',
      month: 'short',
      year:  'numeric',
    });
  } catch (e) {
    return String(value);
  }
}

