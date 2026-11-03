# edyc-automation

Playwright-based script that automates creation of user accounts on **edyc.co.uk** by reading member details from a spreadsheet.

---

## Prerequisites

- [Node.js](https://nodejs.org/) v18 or later
- Google Chrome installed on your machine

---

## Setup

```bash
# Install dependencies (first time only)
npm install
npx playwright install chromium
```

---

## Spreadsheet format

The script reads from **`members.xlsx`**, sheet **"APRIL 26"**.

| Column | Field |
|--------|-------|
| B | Membership Type (must match a dropdown option — see below) |
| D | Full name (e.g. `Alan Bell`) |
| E | Email address |

Row 1 is treated as a header and is ignored.

**Valid Membership Type values** (case-insensitive):
- Single Membership
- Family Membership
- Young Adult
- Temporary Crewing
- Outport Membership
- Senior Membership
- Honorary Member
- Temporary Member

---

## Running the script

### 🔄 Dry run — safe preview, nothing is sent

```bash
node automate.js
```

- Fills in every form field including the membership type dropdown
- Pauses **1 second** on each completed form so you can verify the details
- Clicks **Cancel** — no invitations are sent
- Produces a report in `dry-run-report.txt`

Use this to verify the spreadsheet data looks correct before going live.

---

### 🚀 Live run — sends the invitations for real

```bash
node automate.js --live
```

- Fills in every form field including the membership type dropdown
- Pauses **3 seconds** on each completed form
- Clicks **Send Invitation** — invitations are sent immediately
- Produces a report in `live-run-report.txt`

> ⚠️ **Only run this when you are ready.** Invitations are emailed to members as soon as each form is submitted.

---

### Step-by-step (both modes)

1. Chrome opens and navigates to `https://edyc.co.uk/login`.
2. **You** type your username and password and complete the login.
3. Switch back to the terminal and press **Enter**.
4. The script automatically:
   - Clicks **System Administrator → Admin Console**
   - Clicks **Membership** in the left sidebar
   - Clicks the **Invitations** tab
5. For each member row in the spreadsheet the script:
   - Clicks **Send Invite** to open the modal
   - Fills in **Email Address**, **Confirm Email**, **First Name**, **Last Name**, and **Membership Type**
   - Pauses so you can review the values (1 s dry run / 3 s live)
   - Clicks **Cancel** *(dry run)* or **Send Invitation** *(live)*
6. A summary report is printed to the terminal and saved to a `.txt` file.
7. Press **Enter** to close Chrome.

---

## Reports

| File | Created by |
|------|-----------|
| `dry-run-report.txt` | `node automate.js` |
| `live-run-report.txt` | `node automate.js --live` |

Both reports list every member that was successfully processed and any rows that were skipped or failed (with the reason).
