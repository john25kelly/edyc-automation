/**
 * EDYC Member Account Creation — DRY RUN (Cancel — no invitations sent)
 *
 * Spreadsheet:  members.xlsx  →  sheet "APRIL 26"
 *   Column B:  Membership type
 *   Column D:  Full name  (e.g. "Alan Bell")
 *   Column E:  Email address
 *
 * Flow:
 *  1.  Open Chrome → edyc.co.uk/login  (you log in manually).
 *  2.  Navigate: System Administrator → Admin Console.
 *  3.  Click Membership in the sidebar.
 *  4.  Click the Invitations tab.
 *  5.  For every member row:
 *        a. Click "Send Invite" to open the modal.
 *        b. Fill Email, Confirm Email, First Name, Last Name.
 *        c. Select Membership Type from the dropdown.
 *        d. Pause 5 s so you can verify the data.
 *        e. Click "Cancel" — nothing is submitted (DRY RUN).
 *        f. Record success or failure.
 *  6.  Print a full report at the end.
 */

const { chromium } = require('playwright');
const XLSX        = require('xlsx');
const path        = require('path');
const readline    = require('readline');
const fs          = require('fs');

// ─── Configuration ────────────────────────────────────────────────────────────

const LOGIN_URL   = 'https://edyc.co.uk/login';
const SPREADSHEET = path.join(__dirname, 'members.xlsx');
const SHEET_NAME  = 'APRIL 26';

// --live      : click "Send Invitation" (default is dry run → Cancel)
// --secondary : read name from col G, email from col H  (default → col D, col E)
//               Column B (membership type) is used in both modes.
//               IMPORTANT: --secondary only processes secondary members;
//               primary members (cols D/E) are completely ignored.
const LIVE_MODE      = process.argv.includes('--live');
const SECONDARY_MODE = process.argv.includes('--secondary');

// ─── Membership type aliases ───────────────────────────────────────────────────
// Maps spreadsheet values (lowercase) to the exact dropdown label on the form.
// Add more entries here if other codes in the sheet need remapping.
const MEMBERSHIP_ALIASES = {
  'cadet':  'Young Adult',
  'junior': 'Young Adult',
};

const PAUSE_MS = 1000;   // 1s pause after filling the form (both modes)

// ─── Helpers ──────────────────────────────────────────────────────────────────

function waitForEnter(prompt) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(prompt, () => { rl.close(); resolve(); });
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Read members from the "APRIL 26" sheet.
 *
 * Primary mode   (default):    col B = membership type, col D = name, col E = email
 * Secondary mode (--secondary): col B = membership type, col G = name, col H = email
 *
 * Rows where name OR email is blank are silently skipped.
 */
function readMembers() {
  const memberLabel = SECONDARY_MODE ? 'secondary' : 'primary';
  const nameCol     = SECONDARY_MODE ? 6 : 3;   // G=6, D=3  (0-based)
  const emailCol    = SECONDARY_MODE ? 7 : 4;   // H=7, E=4

  console.log(`\n📂  Reading: ${SPREADSHEET}  (sheet: "${SHEET_NAME}", ${memberLabel} members)`);

  const workbook = XLSX.readFile(SPREADSHEET);
  const sheet    = workbook.Sheets[SHEET_NAME];

  if (!sheet) {
    const available = workbook.SheetNames.join(', ');
    throw new Error(`Sheet "${SHEET_NAME}" not found. Available sheets: ${available}`);
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  if (rows.length < 2) throw new Error('The sheet has no data rows (only a header, or is empty).');

  const members = [];
  const skipped = [];

  for (let i = 1; i < rows.length; i++) {
    const row            = rows[i];
    const membershipType = (row[1]        || '').toString().trim();  // Column B
    const fullName       = (row[nameCol]  || '').toString().trim();
    const email          = (row[emailCol] || '').toString().trim();

    // Skip entirely blank rows
    if (!fullName && !email && !membershipType) continue;

    const errors = [];
    if (!fullName)       errors.push('name is blank');
    if (!email)          errors.push('email is blank');
    if (!membershipType) errors.push('membership type is blank');

    if (errors.length) {
      skipped.push({ rowNum: i + 1, fullName, email, membershipType, reason: errors.join('; ') });
      console.warn(`  ⚠️  Row ${i + 1} skipped — ${errors.join('; ')}`);
      continue;
    }

    // Split "Alan Bell" → firstName="Alan", lastName="Bell"
    const spaceIdx  = fullName.indexOf(' ');
    const firstName = spaceIdx === -1 ? fullName : fullName.slice(0, spaceIdx).trim();
    const lastName  = spaceIdx === -1 ? ''       : fullName.slice(spaceIdx + 1).trim();

    members.push({ firstName, lastName, fullName, email, membershipType, rowNum: i + 1 });
  }

  console.log(`✅  ${members.length} valid ${memberLabel} member(s) found; ${skipped.length} row(s) skipped.\n`);
  return { members, skipped };
}

// ─── Main ─────────────────────────────────────────────────────────────────────

(async () => {

  // ── 1. Read spreadsheet ───────────────────────────────────────────────────
  const { members, skipped: skippedAtRead } = readMembers();
  if (members.length === 0) {
    console.error('❌  No valid members to process. Please check the spreadsheet.');
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

  // ── 5. Membership sidebar ─────────────────────────────────────────────────
  console.log('🖱️   Clicking "Membership" in the sidebar...');
  await page.click('text=Membership');
  await page.waitForLoadState('domcontentloaded');
  console.log(`📍  Now at: ${page.url()}\n`);

  // ── 6. Invitations tab ────────────────────────────────────────────────────
  console.log('🖱️   Clicking "Invitations" tab...');
  await page.click('text=Invitations');
  await sleep(800);
  console.log('✅  Invitations tab active.\n');

  // ── 7. Process each member ────────────────────────────────────────────────
  const memberLabel = SECONDARY_MODE ? 'SECONDARY' : 'PRIMARY';
  if (LIVE_MODE) {
    console.log(`🚀  LIVE MODE [${memberLabel}] — invitations WILL be sent (Send Invitation clicked).\n`);
  } else {
    console.log(`🔄  DRY RUN [${memberLabel}] — forms will be filled but NOT submitted (Cancel clicked).\n`);
  }
  console.log('══════════════════════════════════════════════════════════\n');

  const succeeded = [];
  const failed    = [];

  for (let i = 0; i < members.length; i++) {
    const { firstName, lastName, fullName, email, membershipType, rowNum } = members[i];
    console.log(`[${i + 1}/${members.length}]  Row ${rowNum}: ${fullName} <${email}>  [${membershipType}]`);

    try {
      // ── a. Open the Send Membership Invitation modal ──────────────────────
      await page.click('button:has-text("Send Invite"), a:has-text("Send Invite")');
      await page.waitForSelector('text=Send Membership Invitation', { timeout: 10000 });
      await sleep(400);

      // ── b. Fill email fields ──────────────────────────────────────────────
      await page.locator('input[placeholder="member@example.com"]').fill(email);
      await page.locator('input[placeholder="Confirm email address"]').fill(email);

      // ── c. Fill name fields ───────────────────────────────────────────────
      await page.locator('input[placeholder="John"]').fill(firstName);
      await page.locator('input[placeholder="Doe"]').fill(lastName);

      // ── d. Select Membership Type ─────────────────────────────────────────
      console.log(`  🔽  Selecting membership type: "${membershipType}"`);

      const selectEl = page.locator('select').first();
      await selectEl.waitFor({ timeout: 5000 });

      const optionLabels = await selectEl.locator('option').allInnerTexts();
      console.log(`      Available options: ${optionLabels.join(' | ')}`);

      const needle = membershipType.toLowerCase().trim();

      // 0️⃣  Alias map (e.g. CADET → Young Adult)
      const aliasedLabel = MEMBERSHIP_ALIASES[needle];

      // 1️⃣  Exact match (case-insensitive)
      let matchedLabel = aliasedLabel || optionLabels.find(l => l.toLowerCase().trim() === needle);

      // 2️⃣  Partial: spreadsheet value contained in option label
      if (!matchedLabel) matchedLabel = optionLabels.find(l => l.toLowerCase().includes(needle));

      // 3️⃣  Partial: option label contained in spreadsheet value
      if (!matchedLabel) matchedLabel = optionLabels.find(l => needle.includes(l.toLowerCase().trim()));

      if (!matchedLabel) {
        throw new Error(
          `Membership type "${membershipType}" not found in dropdown. ` +
          `Available: ${optionLabels.join(', ')}`
        );
      }

      if (matchedLabel.toLowerCase().trim() !== needle) {
        console.log(`      ℹ️  Mapped "${membershipType}" → "${matchedLabel}"`);
      }

      await selectEl.selectOption({ label: matchedLabel });
      console.log(`  ✅  Membership type selected: "${matchedLabel}".`);

      // ── e. Pause ─────────────────────────────────────────────────────────
      console.log(`  👀  Pausing ${PAUSE_MS / 1000}s…`);
      await sleep(PAUSE_MS);

      // ── f. Submit or cancel ───────────────────────────────────────────────
      const formOverlay = page.locator('div').filter({ hasText: 'Send Membership Invitation' }).last();

      if (LIVE_MODE) {
        console.log('  📨  Clicking "Send Invitation"…');
        await formOverlay.getByRole('button', { name: 'Send Invitation' }).click();
      } else {
        console.log('  ❌  Clicking Cancel (dry run)…');
        await formOverlay.getByRole('button', { name: 'Cancel' }).click();
      }

      await page.waitForSelector('text=Send Membership Invitation', { state: 'hidden', timeout: 15000 });
      await sleep(500);

      succeeded.push({ rowNum, fullName, email, membershipType });
      console.log(`  ✅  Done.\n`);

    } catch (err) {
      const reason = err.message.split('\n')[0];
      console.error(`  ❌  FAILED: ${reason}\n`);
      failed.push({ rowNum, fullName, email, membershipType, reason });

      try {
        const overlay = page.locator('div').filter({ hasText: 'Send Membership Invitation' }).last();
        const cancelBtn = overlay.getByRole('button', { name: 'Cancel' });
        if (await cancelBtn.isVisible({ timeout: 2000 })) {
          await cancelBtn.click();
          await page.waitForSelector('text=Send Membership Invitation', { state: 'hidden', timeout: 5000 });
        }
      } catch { /* ignore */ }
      await sleep(600);
    }
  }

  // ── 8. Final report ───────────────────────────────────────────────────────
  const allIssues  = [...skippedAtRead, ...failed];
  const modeLabel  = `${LIVE_MODE ? 'LIVE' : 'DRY'} RUN — ${SECONDARY_MODE ? 'SECONDARY' : 'PRIMARY'} MEMBERS`;
  const reportFile = `${LIVE_MODE ? 'live' : 'dry'}-run-${SECONDARY_MODE ? 'secondary' : 'primary'}-report.txt`;
  const lines      = [];

  lines.push('');
  lines.push('══════════════════════════════════════════════════════════════════');
  lines.push(`                    ${modeLabel} — FINAL REPORT`);
  lines.push('══════════════════════════════════════════════════════════════════');
  lines.push(`  Total data rows in sheet : ${members.length + skippedAtRead.length}`);
  lines.push(`  ${LIVE_MODE ? 'Successfully sent      ' : 'Successfully previewed '}  : ${succeeded.length}`);
  lines.push(`  Automation failures      : ${failed.length}`);
  lines.push(`  Skipped (bad data)       : ${skippedAtRead.length}`);
  lines.push('');

  if (succeeded.length) {
    lines.push('──────────────────────────────────────────────────────────────────');
    lines.push(`  ✅  ${LIVE_MODE ? 'INVITATIONS SENT' : 'SUCCESSFULLY PREVIEWED'} (${succeeded.length})`);
    lines.push('──────────────────────────────────────────────────────────────────');
    succeeded.forEach(({ rowNum, fullName, email, membershipType }) => {
      lines.push(`  Row ${String(rowNum).padStart(3)}  ${fullName.padEnd(28)}  ${email.padEnd(38)}  [${membershipType}]`);
    });
    lines.push('');
  }

  if (allIssues.length) {
    lines.push('──────────────────────────────────────────────────────────────────');
    lines.push(`  ❌  ISSUES / SKIPPED (${allIssues.length})`);
    lines.push('──────────────────────────────────────────────────────────────────');
    allIssues.forEach(({ rowNum, fullName, email, membershipType, reason }) => {
      lines.push(`  Row ${String(rowNum).padStart(3)}  ${(fullName||'(blank)').padEnd(28)}  ${(email||'(blank)').padEnd(38)}  Reason: ${reason}`);
    });
    lines.push('');
  }

  lines.push('══════════════════════════════════════════════════════════════════');

  const report = lines.join('\n');
  console.log(report);

  const reportPath = path.join(__dirname, reportFile);
  fs.writeFileSync(reportPath, report, 'utf8');
  console.log(`\n📄  Report saved to: ${reportPath}\n`);

  await waitForEnter('Press ENTER to close the browser: ');
  await browser.close();
})();
