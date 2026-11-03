/**
 * EDYC Duty Roster — Race Officer Duties (DRY RUN by default)
 *
 * Reads duties.xlsx (Sheet1, from row 124 onwards).
 * For every row where column I or column J has a name:
 *   - Opens the "+ New Duty" modal
 *   - Duty Type  : Race Officer (I) | Rib Driver (J) | Daisy Driver (K) | Duty House Officer (L) | Bar Duty (N, times from M)
 *   - Date       : from column B (falls back to nearest preceding populated date)
 *   - Start Time : 17:00
 *   - End Time   : 21:30
 *   - Assigns to the member named in the relevant column
 *   - Pauses 5 s, then clicks Cancel (dry run) or Create (--live)
 *
 * Usage:
 *   node duties.js          # dry run  — clicks Cancel
 *   node duties.js --live   # live run — clicks Create
 */

const { chromium } = require('playwright');
const XLSX         = require('xlsx');
const path         = require('path');
const readline     = require('readline');
const fs           = require('fs');

// ─── Configuration ────────────────────────────────────────────────────────────

const LOGIN_URL    = 'https://edyc.co.uk/login';
const SPREADSHEET  = path.join(__dirname, 'duties.xlsx');
const START_ROW    = 124;   // 1-based row number where data begins
const PAUSE_MS     = 1000;  // pause after filling form so you can verify

const LIVE_MODE    = process.argv.includes('--live');

// ─── Helpers ──────────────────────────────────────────────────────────────────

function waitForEnter(prompt) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(prompt, () => { rl.close(); resolve(); });
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Convert Excel date serial → JS Date (UTC) */
function excelSerialToDate(serial) {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
}

/** Format a JS Date as YYYY-MM-DD (for HTML date inputs) */
function toInputDate(jsDate) {
  const y  = jsDate.getUTCFullYear();
  const m  = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
  const d  = String(jsDate.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/** Format a JS Date as DD/MM/YYYY (for display/report) */
function toDisplayDate(jsDate) {
  const y  = jsDate.getUTCFullYear();
  const m  = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
  const d  = String(jsDate.getUTCDate()).padStart(2, '0');
  return `${d}/${m}/${y}`;
}

/** Format a JS Date as "D MMM YYYY" matching how the UI renders it, e.g. "5 Aug 2026" */
function toUIDate(jsDate) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${jsDate.getUTCDate()} ${months[jsDate.getUTCMonth()]} ${jsDate.getUTCFullYear()}`;
}

/** Strip parenthetical suffixes, e.g. "Alan Bell (Vice Commodore)" → "Alan Bell" */
function stripRole(name) {
  return name.replace(/\s*\(.*?\)\s*/g, '').trim();
}

/**
 * Parse a time-range string from column M into { startHH, startMM, endHH, endMM }.
 * Handles mixed separators and trailing text, e.g.:
 *   "17.30-19.45"  "12:00-13:30"  "16.00 - 18.30"  "10.00 - 11.45 food"
 * Returns null if the string cannot be parsed.
 */
function parseTimeRange(raw) {
  if (!raw) return null;
  const m = String(raw).match(/(\d{1,2})[.:](\d{2})\s*[-–]\s*(\d{1,2})[.:](\d{2})/);
  if (!m) return null;
  return {
    startHH: String(m[1]).padStart(2, '0'),
    startMM: m[2],
    endHH:   String(m[3]).padStart(2, '0'),
    endMM:   m[4],
  };
}

// ─── Read spreadsheet ─────────────────────────────────────────────────────────

function readDuties() {
  console.log(`\n📂  Reading: ${SPREADSHEET}`);

  const workbook = XLSX.readFile(SPREADSHEET);
  const sheet    = workbook.Sheets[workbook.SheetNames[0]];
  const rows     = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const duties         = [];
  let   lastDateSerial = null;
  const today          = new Date();
  today.setUTCHours(0, 0, 0, 0);  // compare date only, not time

  // Column → duty type mapping  (0-based column index)
  // timesColIndex: if set, read start/end times from that column (Bar Duty).
  // Otherwise, times are determined by day-of-week (Saturday vs weekday).
  const DUTY_COLUMNS = [
    { colIndex: 8,  dutyTypeLabel: 'Race Officer'       },                      // Column I
    { colIndex: 9,  dutyTypeLabel: 'Rib Driver'         },                      // Column J
    { colIndex: 10, dutyTypeLabel: 'Daisy Driver'       },                      // Column K
    { colIndex: 11, dutyTypeLabel: 'Duty House Officer' },                      // Column L
    { colIndex: 13, dutyTypeLabel: 'Bar Duty', timesColIndex: 12 },             // Column N (times from M)
  ];

  for (let i = START_ROW - 1; i < rows.length; i++) {
    const row      = rows[i];
    const dateCell = row[1];   // Column B

    // Update the rolling date
    if (dateCell && typeof dateCell === 'number' && dateCell > 40000) {
      lastDateSerial = dateCell;
    }

    for (const { colIndex, dutyTypeLabel, timesColIndex } of DUTY_COLUMNS) {
      const memberName = String(row[colIndex] || '').trim();
      if (!memberName) continue;

      if (!lastDateSerial) {
        console.warn(`  ⚠️  Row ${i + 1} [${dutyTypeLabel}]: "${memberName}" — no date found yet, skipping.`);
        continue;
      }

      const jsDate      = excelSerialToDate(lastDateSerial);
      const inputDate   = toInputDate(jsDate);
      const displayDate = toDisplayDate(jsDate);
      const uiDate      = toUIDate(jsDate);

      // Skip duties in the past
      if (jsDate < today) continue;

      let startHH, startMM, endHH, endMM, isSaturday;

      if (timesColIndex !== undefined) {
        // Bar Duty: read times from the specified column
        const timeRaw = String(row[timesColIndex] || '').trim();
        const parsed  = parseTimeRange(timeRaw);
        if (!parsed) {
          console.warn(`  ⚠️  Row ${i + 1} [${dutyTypeLabel}]: "${memberName}" — could not parse time "${timeRaw}", skipping.`);
          continue;
        }
        ({ startHH, startMM, endHH, endMM } = parsed);
        isSaturday = jsDate.getUTCDay() === 6;
      } else {
        // All other duties: times based on day of week
        isSaturday = jsDate.getUTCDay() === 6;
        startHH    = isSaturday ? '15' : '17';
        startMM    = isSaturday ? '00' : '30';
        endHH      = isSaturday ? '19' : '22';
        endMM      = isSaturday ? '00' : '45';
      }

      duties.push({
        rowNum:        i + 1,
        memberName:    memberName,
        searchName:    stripRole(memberName),
        inputDate:     inputDate,
        displayDate:   displayDate,
        uiDate:        uiDate,
        isSaturday:    isSaturday,
        startHH, startMM, endHH, endMM,
        dutyTypeLabel: dutyTypeLabel,
      });
    }
  }

  console.log(`✅  ${duties.length} duties to process (Race Officer + Rib Driver + Daisy Driver + Duty House Officer + Bar Duty).\n`);
  return duties;
}

// ─── Fetch pending invitations ────────────────────────────────────────────────

/**
 * Navigates to the Membership → Invitations tab and scrapes every page of
 * the invitations table, returning an array of lowercased row text strings.
 * These are cached once and used for fast in-memory membership checks.
 */
async function fetchPendingInvitations(page) {
  console.log('\n──────────────────────────────────────────────────────────');
  console.log('📋  Fetching pending invitations (all pages)…');
  console.log('──────────────────────────────────────────────────────────');

  // Navigate: Membership sidebar item
  console.log('🖱️   Clicking "Membership" in the sidebar…');
  await page.click('text=Membership');
  await page.waitForLoadState('domcontentloaded');
  await sleep(600);

  // Click the Invitations tab
  console.log('🖱️   Clicking "Invitations" tab…');
  await page.click('text=Invitations');
  await sleep(1000);

  const allRowTexts = [];
  let pageNum = 1;

  while (true) {
    console.log(`  📄  Reading invitations page ${pageNum}…`);

    const rows     = page.locator('table tbody tr');
    const rowCount = await rows.count();

    for (let r = 0; r < rowCount; r++) {
      const text = (await rows.nth(r).innerText()).trim().toLowerCase();
      if (text) allRowTexts.push(text);
    }

    console.log(`      ${rowCount} row(s) read (total so far: ${allRowTexts.length}).`);

    // Look for an enabled Next / › button to paginate
    const nextBtn = page.locator(
      'button:has-text("Next"), a:has-text("Next"), ' +
      'button[aria-label="Next page"], a[aria-label="Next page"], ' +
      'button:has-text("›"), a:has-text("›")'
    ).first();

    const nextExists  = (await nextBtn.count()) > 0;
    const nextEnabled = nextExists && await nextBtn.isEnabled() && !(await nextBtn.getAttribute('disabled'));

    if (!nextEnabled) break;

    await nextBtn.click();
    await sleep(800);
    pageNum++;
  }

  console.log(`\n✅  Cached ${allRowTexts.length} pending invitation row(s) across ${pageNum} page(s).\n`);
  return allRowTexts;
}

/**
 * Returns true if the given name has a pending invitation in the cached rows.
 * Splits the name into words and checks that every word appears in the same row.
 */
function hasPendingInvite(searchName, invitationRows) {
  const parts = stripRole(searchName).toLowerCase().split(/\s+/).filter(Boolean);
  return invitationRows.some(row => parts.every(part => row.includes(part)));
}

// ─── Duplicate check ──────────────────────────────────────────────────────────

/**
 * Search the Duties list page for an existing duty matching the given member,
 * date (DD/MM/YYYY) and duty type.  Returns true if a duplicate is found.
 *
 * Strategy:
 *  1. Type the member's name in the "Search by member name…" box on the list page.
 *  2. Wait for the table to refresh.
 *  3. Read every visible row and check whether both the date and duty type match.
 *  4. Clear the search box before returning.
 */
async function checkForDuplicate(page, searchName, uiDate, dutyTypeLabel) {
  const listSearchBox = page.locator('input[placeholder*="Search by member name"]');
  await listSearchBox.fill(searchName);
  await sleep(1000);   // wait for table to update

  // Grab all table rows
  const rows     = page.locator('table tbody tr');
  const rowCount = await rows.count();

  let found = false;
  for (let r = 0; r < rowCount; r++) {
    const rowText = (await rows.nth(r).innerText()).toLowerCase();
    // Match the UI date format ("5 Aug 2026") and the duty type (case-insensitive)
    if (rowText.includes(uiDate.toLowerCase()) &&
        rowText.includes(dutyTypeLabel.toLowerCase())) {
      found = true;
      break;
    }
  }

  // Clear the search so the full list is restored for the next iteration
  await listSearchBox.fill('');
  await sleep(400);

  return found;
}

// ─── Main ─────────────────────────────────────────────────────────────────────

(async () => {

  // ── 1. Read spreadsheet ───────────────────────────────────────────────────
  const duties = readDuties();
  if (duties.length === 0) {
    console.error('❌  No Race Officer duties found. Check the spreadsheet.');
    process.exit(1);
  }

  // ── 2. Launch Chrome ──────────────────────────────────────────────────────
  const browser = await chromium.launch({ channel: 'chrome', headless: false, slowMo: 60 });
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  const page    = await context.newPage();

  // ── 3. Login ──────────────────────────────────────────────────────────────
  console.log(`🌐  Opening: ${LOGIN_URL}`);
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' });

  console.log('\n──────────────────────────────────────────────────────────');
  console.log('🔐  Enter your username & password in the browser window,');
  console.log('    then come back here once you are logged in.');
  console.log('──────────────────────────────────────────────────────────\n');
  await waitForEnter('   ▶  Press ENTER once you are logged in: ');
  console.log('\n✅  Login confirmed.\n');

  // ── 4. System Administrator → Admin Console ───────────────────────────────
  console.log('🖱️   Opening "System Administrator" dropdown...');
  await page.click('text=System Administrator');
  await sleep(600);

  console.log('🖱️   Clicking "Admin Console"...');
  await page.click('text=Admin Console');
  await page.waitForLoadState('domcontentloaded');
  console.log(`📍  Now at: ${page.url()}\n`);

  // ── 5. Duty Roster sidebar ────────────────────────────────────────────────
  console.log('🖱️   Clicking "Duty Roster" in the sidebar...');
  await page.click('text=Duty Roster');
  await page.waitForLoadState('domcontentloaded');
  await sleep(600);
  console.log(`📍  Now at: ${page.url()}\n`);

  // ── 6. Mode banner ────────────────────────────────────────────────────────
  if (LIVE_MODE) {
    console.log('🚀  LIVE MODE — duties WILL be created (Create button clicked).\n');
  } else {
    console.log('🔄  DRY RUN — forms will be filled but NOT submitted (Cancel clicked).\n');
  }
  console.log('══════════════════════════════════════════════════════════\n');

  // ── 7. Process each duty ──────────────────────────────────────────────────
  const succeeded    = [];
  const failed       = [];
  const missingUsers = [];
  const duplicates   = [];

  for (let i = 0; i < duties.length; i++) {
    const { rowNum, memberName, searchName, inputDate, displayDate, uiDate, dutyTypeLabel,
            isSaturday, startHH, startMM, endHH, endMM } = duties[i];
    console.log(`[${i + 1}/${duties.length}]  Row ${rowNum}: [${dutyTypeLabel}] "${memberName}"  Date: ${displayDate} (${isSaturday ? 'Saturday' : 'Weekday'} — ${startHH}:${startMM}–${endHH}:${endMM})`);

    try {
      // ── a. Duplicate check on the list page ───────────────────────────────
      console.log(`  🔎  Checking for existing duty (searching "${searchName}" / "${uiDate}")…`);
      const isDuplicate = await checkForDuplicate(page, searchName, uiDate, dutyTypeLabel);
      if (isDuplicate) {
        const msg = `Row ${rowNum} [${dutyTypeLabel}]: "${memberName}" on ${displayDate} already exists — skipped.`;
        console.log(`  ⏭️   DUPLICATE: ${msg}`);
        duplicates.push({ rowNum, memberName, dutyTypeLabel, displayDate });
        continue;
      }
      console.log('  ✅  No duplicate found, proceeding.');

      // ── b. Click "+ New Duty" ─────────────────────────────────────────────
      await page.click('button:has-text("New Duty"), a:has-text("New Duty")');
      // Wait for the modal to appear
      await page.waitForSelector('text=Create Duty', { timeout: 10000 });
      await sleep(400);

      // Scope to the modal: find a div that contains BOTH "Create Duty" text AND
      // at least one <select> element (rules out the plain heading div).
      // Fall back to [role="dialog"] if the app uses that attribute.
      let modal = page.locator('[role="dialog"]');
      if ((await modal.count()) === 0) {
        modal = page.locator('div')
          .filter({ hasText: 'Create Duty' })
          .filter({ has: page.locator('select') })
          .last();
      }
      console.log(`  📦  Modal located (role="${await modal.getAttribute('role') ?? 'div'}").`);

      // ── c. Select Duty Type ───────────────────────────────────────────────
      console.log(`  🔽  Selecting Duty Type: "${dutyTypeLabel}"...`);
      const dutyTypeSelect = modal.locator('select').first();
      await dutyTypeSelect.waitFor({ timeout: 5000 });
      const typeOptions = await dutyTypeSelect.locator('option').allInnerTexts();
      console.log(`      Options: ${typeOptions.join(' | ')}`);
      const matched = typeOptions.find(o => o.toLowerCase().includes(dutyTypeLabel.toLowerCase()));
      if (!matched) throw new Error(`"${dutyTypeLabel}" option not found. Available: ${typeOptions.join(', ')}`);
      await dutyTypeSelect.selectOption({ label: matched });
      console.log(`  ✅  Duty Type: "${matched}"`);

      // ── d. Set Date ───────────────────────────────────────────────────────
      console.log(`  📅  Setting date to ${displayDate} (${inputDate})...`);
      const dateInput = modal.locator('input[type="date"]');
      await dateInput.waitFor({ timeout: 5000 });
      // Use JS evaluate to set the value directly — avoids the native date-picker
      // popup which can intercept Playwright's fill() and cause a hang.
      await dateInput.evaluate((el, val) => {
        el.value = val;
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }, inputDate);
      console.log('  ✅  Date set.');

      // ── e. Set Start Time ─────────────────────────────────────────────────
      console.log(`  🕔  Setting Start Time ${startHH}:${startMM}...`);
      const allSelects  = modal.locator('select');
      const selectCount = await allSelects.count();
      console.log(`      Found ${selectCount} <select> element(s).`);
      // Select elements order: [0] Duty Type, [1] Start HH, [2] Start MM, [3] End HH, [4] End MM
      await allSelects.nth(1).selectOption({ value: startHH });
      await allSelects.nth(2).selectOption({ value: startMM });
      console.log(`  ✅  Start Time ${startHH}:${startMM} set.`);

      // ── f. Set End Time ───────────────────────────────────────────────────
      console.log(`  🕘  Setting End Time ${endHH}:${endMM}...`);
      await allSelects.nth(3).selectOption({ value: endHH });
      await allSelects.nth(4).selectOption({ value: endMM });
      console.log(`  ✅  End Time ${endHH}:${endMM} set.`);

      // ── g. Search for member ──────────────────────────────────────────────
      console.log(`  🔍  Searching for member: "${searchName}"...`);
      const searchBox = modal.locator('input[placeholder*="Search by name"], input[placeholder*="search"]').first();
      await searchBox.waitFor({ timeout: 5000 });
      await searchBox.fill(searchName);
      await sleep(800);   // wait for list to filter

      // Find the matching checkbox item within the modal
      const memberItems = modal.locator('text=' + searchName);
      const itemCount   = await memberItems.count();
      console.log(`      ${itemCount} result(s) found for "${searchName}".`);

      if (itemCount === 0) {
        const msg = `Row ${rowNum} [${dutyTypeLabel}]: member "${memberName}" not found in member list.`;
        console.warn(`  ⚠️  ${msg}`);
        missingUsers.push(msg);
        // Close modal and continue
        await modal.getByRole('button', { name: 'Cancel' }).click();
        await modal.waitFor({ state: 'hidden', timeout: 10000 });
        await sleep(400);
        failed.push({ rowNum, memberName, dutyTypeLabel, displayDate, reason: 'Member not found' });
        continue;
      }

      // Click the first matching result (checkbox or its container)
      await memberItems.first().click();
      console.log(`  ✅  Member "${searchName}" selected.`);

      // ── h. Pause for visual verification ─────────────────────────────────
      console.log(`  👀  Pausing ${PAUSE_MS / 1000}s for verification…`);
      await sleep(PAUSE_MS);

      // ── i. Cancel (dry run) or Create (live) — scoped to modal ───────────
      if (LIVE_MODE) {
        console.log('  ✔️   Clicking "Create"…');
        await modal.getByRole('button', { name: 'Create' }).click();
      } else {
        console.log('  ❌  Clicking "Cancel" (dry run)…');
        await modal.getByRole('button', { name: 'Cancel' }).click();
      }

      await modal.waitFor({ state: 'hidden', timeout: 15000 });
      await sleep(500);

      succeeded.push({ rowNum, memberName, dutyTypeLabel, displayDate });
      console.log(`  ✅  Done.\n`);

    } catch (err) {
      const reason = err.message.split('\n')[0];
      console.error(`  ❌  FAILED: ${reason}\n`);
      failed.push({ rowNum, memberName, dutyTypeLabel, displayDate, reason });

      // Try to dismiss the modal if it's still open — scoped to avoid background buttons
      try {
        let modal = page.locator('[role="dialog"]');
        if ((await modal.count()) === 0) {
          modal = page.locator('div')
            .filter({ hasText: 'Create Duty' })
            .filter({ has: page.locator('select') })
            .last();
        }
        const cancelBtn = modal.getByRole('button', { name: 'Cancel' });
        if (await cancelBtn.isVisible({ timeout: 2000 })) {
          await cancelBtn.click();
          await modal.waitFor({ state: 'hidden', timeout: 5000 });
        }
      } catch { /* ignore */ }
      await sleep(600);
    }
  }

  // ── 8. Fetch pending invitations for report enrichment ───────────────────
  let invitationRows = [];
  try {
    invitationRows = await fetchPendingInvitations(page);
  } catch (err) {
    console.warn(`⚠️  Could not fetch invitations: ${err.message.split('\n')[0]}`);
    console.warn('    "Member not found" entries will not show invite status.\n');
  }

  // Navigate back to Duty Roster so the page is in a clean state
  try {
    await page.click('text=Duty Roster');
    await page.waitForLoadState('domcontentloaded');
    await sleep(400);
  } catch { /* non-fatal */ }

  // ── 9. Final report ───────────────────────────────────────────────────────
  const modeLabel  = LIVE_MODE ? 'LIVE RUN' : 'DRY RUN';
  const lines      = [];

  lines.push('');
  lines.push('══════════════════════════════════════════════════════════════════');
  lines.push(`                  DUTY ROSTER — ${modeLabel} — FINAL REPORT`);
  lines.push('══════════════════════════════════════════════════════════════════');
  lines.push(`  Total duties processed : ${duties.length}`);
  lines.push(`  Succeeded              : ${succeeded.length}`);
  lines.push(`  Duplicates skipped     : ${duplicates.length}`);
  lines.push(`  Failed / skipped       : ${failed.length}`);
  lines.push(`  Missing users          : ${missingUsers.length}`);
  lines.push('');

  if (succeeded.length) {
    lines.push('──────────────────────────────────────────────────────────────────');
    lines.push(`  ✅  SUCCEEDED (${succeeded.length})`);
    lines.push('──────────────────────────────────────────────────────────────────');
    succeeded.forEach(({ rowNum, memberName, dutyTypeLabel, displayDate }) => {
      lines.push(`  Row ${String(rowNum).padStart(3)}  ${displayDate}  [${dutyTypeLabel.padEnd(18)}]  ${memberName}`);
    });
    lines.push('');
  }

  if (duplicates.length) {
    lines.push('──────────────────────────────────────────────────────────────────');
    lines.push(`  ⏭️   DUPLICATES SKIPPED (${duplicates.length})`);
    lines.push('──────────────────────────────────────────────────────────────────');
    duplicates.forEach(({ rowNum, memberName, dutyTypeLabel, displayDate }) => {
      lines.push(`  Row ${String(rowNum).padStart(3)}  ${displayDate}  [${dutyTypeLabel.padEnd(18)}]  ${memberName}`);
    });
    lines.push('');
  }

  if (failed.length) {
    lines.push('──────────────────────────────────────────────────────────────────');
    lines.push(`  ❌  FAILED / SKIPPED (${failed.length})`);
    lines.push('──────────────────────────────────────────────────────────────────');
    failed.forEach(({ rowNum, memberName, dutyTypeLabel, displayDate, reason }) => {
      const inviteTag = (reason === 'Member not found' && hasPendingInvite(memberName, invitationRows))
        ? '  [invite pending]'
        : '';
      lines.push(`  Row ${String(rowNum).padStart(3)}  ${displayDate}  [${dutyTypeLabel.padEnd(18)}]  ${memberName}  — ${reason}${inviteTag}`);
    });
    lines.push('');
  }

  lines.push('══════════════════════════════════════════════════════════════════');

  const report = lines.join('\n');
  console.log(report);

  const reportPath = path.join(__dirname, `duties-${LIVE_MODE ? 'live' : 'dry'}-run-report.txt`);
  fs.writeFileSync(reportPath, report, 'utf8');
  console.log(`\n📄  Report saved to: ${reportPath}`);

  // ── 10. Missing users file ────────────────────────────────────────────────
  if (missingUsers.length > 0) {
    const enriched = missingUsers.map(msg => {
      // Extract name from message format: "Row N [Type]: "Name" not found..."
      const nameMatch = msg.match(/"([^"]+)"/);
      if (nameMatch && hasPendingInvite(nameMatch[1], invitationRows)) {
        return msg + '  [invite pending]';
      }
      return msg;
    });
    const missingPath = path.join(__dirname, 'missing_users.txt');
    fs.writeFileSync(missingPath, enriched.join('\n') + '\n', 'utf8');
    console.log(`⚠️   Missing users saved to: ${missingPath}`);
  }

  await waitForEnter('\nPress ENTER to close the browser: ');
  await browser.close();
})();
