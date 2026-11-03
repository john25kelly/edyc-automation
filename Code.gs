/**
 * EDYC Invitations — Google Apps Script
 *
 * Adds an "EDYC Invitations" menu to the spreadsheet.
 * "Refresh List" prompts for credentials, calls the EDYC API,
 * and writes all pending invitations as rows.
 */

// ─── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('EDYC Invitations')
    .addItem('Refresh List', 'showLoginDialog')
    .addSeparator()
    .addItem('Send Email Reminders', 'sendEmailReminders')
    .addToUi();
}

// ─── Login dialog (HTML popup) ────────────────────────────────────────────────

function showLoginDialog() {
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
  <h3>🔐 EDYC — Fetch Pending Invitations</h3>

  <label for="username">Username</label>
  <input type="text" id="username" placeholder="e.g. WeAreEDYC" autocomplete="username" />

  <label for="password">Password</label>
  <input type="password" id="password" placeholder="Password" autocomplete="current-password" />

  <div id="error"></div>
  <div id="spinner">⏳ Fetching invitations, please wait…</div>

  <div id="buttons">
    <button id="cancelBtn" onclick="google.script.host.close()">Cancel</button>
    <button id="submitBtn" onclick="submit()">Refresh List</button>
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
        .fetchAndWriteInvitations(username, password);
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

  SpreadsheetApp.getUi().showModalDialog(html, 'EDYC Invitations');
}

// ─── Server-side: authenticate + fetch + write ────────────────────────────────

function fetchAndWriteInvitations(username, password) {
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

  // 2. Fetch all pages of pending invitations
  const allInvitations = [];
  let page = 1;
  const limit = 25;

  while (true) {
    const url = API_BASE + '/api/membership/invitations?page=' + page + '&limit=' + limit + '&status=pending';
    const invRes = UrlFetchApp.fetch(url, {
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

    if (invRes.getResponseCode() !== 200) {
      throw new Error('Failed to fetch invitations (HTTP ' + invRes.getResponseCode() + ').');
    }

    const invJson = JSON.parse(invRes.getContentText());

    // API shape: { success: true, data: { invitations: [...], totalCount: N } }
    const items =
      (invJson.data && Array.isArray(invJson.data.invitations)) ? invJson.data.invitations :
      Array.isArray(invJson)                                    ? invJson                  :
      Array.isArray(invJson.data)                               ? invJson.data             :
      Array.isArray(invJson.items)                              ? invJson.items            :
      Array.isArray(invJson.invitations)                        ? invJson.invitations      :
      null;

    if (!items) {
      throw new Error('Unexpected API response shape: ' + invRes.getContentText().slice(0, 200));
    }

    allInvitations.push.apply(allInvitations, items);

    const total =
      (invJson.data && (invJson.data.totalCount || invJson.data.total)) ||
      invJson.total || invJson.totalCount;

    const hasMore = items.length === limit && (total == null || allInvitations.length < total);
    if (!hasMore) break;
    page++;
  }

  // 3. Write to spreadsheet
  writeToSheet(allInvitations);
}

// ─── Write rows to sheet ──────────────────────────────────────────────────────

function writeToSheet(invitations) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('Pending Invitations');

  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Pending Invitations');
  }

  // Clear existing content
  sheet.clearContents();

  // Header row
  const headers = [
    '#',
    'First Name',
    'Last Name',
    'Email',
    'Membership Type',
    'Invited By',
    'Invited At',
    'Expires At',
    'EDYC Join Date',
    'Status',
    'Notes',
  ];

  const rows = [headers];

  invitations.forEach(function(inv, i) {
    rows.push([
      i + 1,
      inv.firstName  || inv.first_name  || '',
      inv.lastName   || inv.last_name   || '',
      inv.email      || '',
      inv.membershipType || inv.membership_type || inv.type || '',
      inv.invitedBy  || inv.invited_by  || '',
      formatDate(inv.invitedAt  || inv.invited_at  || inv.createdAt || inv.created_at),
      formatDate(inv.expiresAt  || inv.expires_at),
      formatDate(inv.edycJoinDate || inv.edyc_join_date),
      inv.status     || 'pending',
      inv.notes      || '',
    ]);
  });

  // Write all rows at once (much faster than row-by-row)
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  // Style the header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Auto-resize columns
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // Activate the sheet so the user sees the result
  ss.setActiveSheet(sheet);

  // Show completion toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Found ' + invitations.length + ' pending invitation(s).',
    '✅ EDYC Invitations Refreshed',
    5
  );
}

// ─── Send Email Reminders ─────────────────────────────────────────────────────

function sendEmailReminders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pending Invitations');

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No "Pending Invitations" sheet found.\nPlease run "Refresh List" first.'
    );
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No invitations found in the sheet.');
    return;
  }

  // Confirm before sending
  const ui       = SpreadsheetApp.getUi();
  const count    = lastRow - 1; // exclude header
  const response = ui.alert(
    'Send Email Reminders',
    'This will send a reminder email to ' + count + ' pending member(s).\n\nDo you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Read columns B (First Name) and D (Email) — rows 2 onwards
  // Col A=#, B=First Name, C=Last Name, D=Email
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 4); // cols A–D
  const data      = dataRange.getValues();

  let sent   = 0;
  let failed = 0;
  const errors = [];

  data.forEach(function(row) {
    const firstName = String(row[1] || '').trim(); // col B
    const email     = String(row[3] || '').trim(); // col D

    if (!email || !email.includes('@')) {
      failed++;
      errors.push('Skipped (no valid email): ' + (firstName || '(unknown)'));
      return;
    }

    const subject = 'Complete Your EDYC Membership Registration';
    const body    = buildEmailBody(firstName);

    try {
      GmailApp.sendEmail(email, subject, body, {
        from:     'honsec@edyc.co.uk',
        name:     'EDYC',
        htmlBody: buildEmailHtml(firstName),
      });
      sent++;
    } catch (e) {
      failed++;
      errors.push(email + ': ' + e.message);
    }
  });

  // Summary toast + alert
  let summary = '✅ Sent: ' + sent + ' email(s).';
  if (failed > 0) summary += '\n⚠️ Failed/Skipped: ' + failed + '.';
  if (errors.length > 0) summary += '\n\nDetails:\n' + errors.join('\n');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Sent ' + sent + ' reminder(s)' + (failed > 0 ? ', ' + failed + ' failed.' : '.'),
    '✅ Email Reminders',
    8
  );

  if (failed > 0) {
    ui.alert('Email Reminder Summary', summary, ui.ButtonSet.OK);
  }
}

// ─── Email templates ──────────────────────────────────────────────────────────

function buildEmailBody(firstName) {
  const name = firstName || 'there';
  return [
    'Hi ' + name + ',',
    '',
    'This is a friendly reminder that you have a pending invitation to join EDYC as a member.',
    '',
    'To complete your membership registration, please visit:',
    'https://edyc.co.uk',
    '',
    'If you have any questions or need assistance, please don\'t hesitate to get in touch.',
    '',
    'Kind regards,',
    'The EDYC Team',
  ].join('\n');
}

function buildEmailHtml(firstName) {
  const name = firstName || 'there';
  return `
<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 0;">
    <tr>
      <td align="center">
        <table width="560" cellpadding="0" cellspacing="0"
               style="background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">

          <!-- Header -->
          <tr>
            <td style="background:#1a73e8;padding:28px 36px;">
              <h1 style="margin:0;color:#ffffff;font-size:22px;font-weight:700;letter-spacing:-0.5px;">
                EDYC
              </h1>
              <p style="margin:4px 0 0;color:#c8deff;font-size:13px;">East Dunbartonshire Youth Council</p>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:32px 36px 24px;">
              <p style="margin:0 0 16px;font-size:16px;color:#333;">Hi <strong>${name}</strong>,</p>

              <p style="margin:0 0 16px;font-size:15px;color:#444;line-height:1.6;">
                This is a friendly reminder that you have a <strong>pending invitation</strong>
                to join EDYC as a member — and we'd love to have you on board!
              </p>

              <p style="margin:0 0 24px;font-size:15px;color:#444;line-height:1.6;">
                To complete your membership registration, simply click the button below:
              </p>

              <!-- CTA Button -->
              <table cellpadding="0" cellspacing="0" style="margin:0 0 28px;">
                <tr>
                  <td style="background:#1a73e8;border-radius:6px;">
                    <a href="https://edyc.co.uk"
                       style="display:inline-block;padding:12px 28px;color:#ffffff;font-size:15px;
                              font-weight:600;text-decoration:none;">
                      Complete My Registration →
                    </a>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 16px;font-size:14px;color:#666;line-height:1.6;">
                If you have any questions or need help, please don't hesitate to reach out to us
                — we're happy to assist.
              </p>

              <p style="margin:0;font-size:14px;color:#444;">
                Kind regards,<br/>
                <strong>The EDYC Team</strong>
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#f8f9fa;padding:16px 36px;border-top:1px solid #e8eaed;">
              <p style="margin:0;font-size:12px;color:#999;text-align:center;">
                © ${new Date().getFullYear()} EDYC — East Dunbartonshire Youth Council
                &nbsp;|&nbsp;
                <a href="https://edyc.co.uk" style="color:#1a73e8;text-decoration:none;">edyc.co.uk</a>
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
`;
}

// ─── Utility ──────────────────────────────────────────────────────────────────

function formatDate(value) {
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
