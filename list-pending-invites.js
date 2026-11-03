/**
 * EDYC — List Pending Membership Invitations
 *
 * 1. Prompts for username & password.
 * 2. Authenticates against the EDYC API to obtain a Bearer token.
 * 3. Fetches ALL pages of pending invitations.
 * 4. Prints a formatted table to the console.
 */

const readline = require('readline');
const https    = require('https');

const API_BASE = 'https://production-api-production-a02e.up.railway.app';

// ─── Helpers ──────────────────────────────────────────────────────────────────

// Single shared rl instance — closing it multiple times breaks piped input
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

function ask(prompt, hidden = false) {
  return new Promise((resolve) => {
    if (hidden && process.stdin.isTTY) {
      // Raw-mode hidden input for interactive TTY
      process.stdout.write(prompt);
      process.stdin.setRawMode(true);
      process.stdin.resume();
      process.stdin.setEncoding('utf8');

      let input = '';
      process.stdin.on('data', function handler(ch) {
        if (ch === '\n' || ch === '\r' || ch === '\u0004') {
          process.stdin.setRawMode(false);
          process.stdin.pause();
          process.stdin.removeListener('data', handler);
          process.stdout.write('\n');
          resolve(input);
        } else if (ch === '\u0003') {
          process.exit();
        } else if (ch === '\u007f') {
          if (input.length > 0) input = input.slice(0, -1);
        } else {
          input += ch;
        }
      });
    } else {
      rl.question(prompt, (answer) => resolve(answer));
    }
  });
}

function request(method, path, body, token) {
  return new Promise((resolve, reject) => {
    const payload = body ? JSON.stringify(body) : null;
    const options = {
      hostname: 'production-api-production-a02e.up.railway.app',
      path,
      method,
      headers: {
        'Content-Type':  'application/json',
        'Accept':        '*/*',
        'Origin':        'https://edyc.co.uk',
        'Referer':       'https://edyc.co.uk/',
        'User-Agent':    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36',
        ...(token        ? { 'Authorization': `Bearer ${token}` } : {}),
        ...(payload      ? { 'Content-Length': Buffer.byteLength(payload) } : {}),
      },
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, body: data ? JSON.parse(data) : null });
        } catch {
          resolve({ status: res.statusCode, body: data });
        }
      });
    });

    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

// ─── Auth ─────────────────────────────────────────────────────────────────────

async function login(username, password) {
  const endpoint = '/api/auth/login';
  console.log(`  🔑  POST ${endpoint} …`);

  // HAR confirms the body uses "username" (not "email")
  const res = await request('POST', endpoint, { username, password });

  if (res.status === 200 || res.status === 201) {
    const token =
      res.body?.token        ||
      res.body?.accessToken  ||
      res.body?.access_token ||
      res.body?.data?.token  ||
      res.body?.data?.accessToken;

    if (token) {
      console.log(`  ✅  Authenticated successfully\n`);
      return token;
    }
    throw new Error(
      `Auth endpoint returned HTTP ${res.status} but no token was found.\nResponse: ${JSON.stringify(res.body).slice(0, 300)}`
    );
  }

  throw new Error(
    `Authentication failed — HTTP ${res.status}.\nResponse: ${JSON.stringify(res.body).slice(0, 300)}`
  );
}

// ─── Fetch invitations (all pages) ───────────────────────────────────────────

async function fetchAllPendingInvitations(token) {
  const all  = [];
  let   page = 1;
  const limit = 25;

  while (true) {
    const path = `/api/membership/invitations?page=${page}&limit=${limit}&status=pending`;
    console.log(`  📥  Fetching page ${page} …`);
    const res = await request('GET', path, null, token);

    if (res.status !== 200) {
      throw new Error(`API returned HTTP ${res.status}: ${JSON.stringify(res.body).slice(0, 300)}`);
    }

    // Handle common pagination shapes
    // Actual API returns: { success: true, data: { invitations: [...], totalCount, ... } }
    const items =
      Array.isArray(res.body?.data?.invitations) ? res.body.data.invitations :
      Array.isArray(res.body)                    ? res.body                  :
      Array.isArray(res.body?.data)              ? res.body.data             :
      Array.isArray(res.body?.items)             ? res.body.items            :
      Array.isArray(res.body?.invitations)       ? res.body.invitations      :
      null;

    if (!items) {
      console.log('\n  ℹ️  Raw response (first 500 chars):', JSON.stringify(res.body).slice(0, 500));
      throw new Error('Unexpected response shape — could not find the invitations array.');
    }

    all.push(...items);

    const total =
      res.body?.data?.totalCount ??
      res.body?.data?.total      ??
      res.body?.total            ??
      res.body?.totalCount       ??
      res.body?.meta?.total;

    const hasMore = items.length === limit && (total == null || all.length < total);
    if (!hasMore) break;
    page++;
  }

  return all;
}

// ─── Print table ──────────────────────────────────────────────────────────────

function printTable(invitations) {
  if (invitations.length === 0) {
    console.log('\n  ℹ️  No pending invitations found.\n');
    return;
  }

  // Normalise field names across possible API shapes
  const rows = invitations.map((inv, i) => ({
    '#':    String(i + 1).padStart(3),
    Name:  [inv.firstName || inv.first_name || '', inv.lastName || inv.last_name || ''].filter(Boolean).join(' ') || inv.name || '—',
    Email: inv.email || '—',
    Type:  inv.membershipType || inv.membership_type || inv.type || '—',
    Sent:  inv.createdAt || inv.created_at || inv.sentAt || '—',
    Status:inv.status || 'pending',
  }));

  const cols = ['#', 'Name', 'Email', 'Type', 'Sent', 'Status'];
  const widths = {};
  cols.forEach(c => widths[c] = c.length);
  rows.forEach(r => cols.forEach(c => { widths[c] = Math.max(widths[c], String(r[c]).length); }));

  const sep  = '  ' + cols.map(c => '─'.repeat(widths[c] + 2)).join('┼') ;
  const head = '  ' + cols.map(c => ` ${c.padEnd(widths[c])} `).join('│');
  const divider = '  ' + cols.map(c => '─'.repeat(widths[c] + 2)).join('┬');
  const bottom  = '  ' + cols.map(c => '─'.repeat(widths[c] + 2)).join('┴');

  console.log('\n' + divider);
  console.log(head);
  console.log(sep);
  rows.forEach(r => {
    console.log('  ' + cols.map(c => ` ${String(r[c]).padEnd(widths[c])} `).join('│'));
  });
  console.log(bottom);
  console.log(`\n  Total pending invitations: ${invitations.length}\n`);
}

// ─── Main ─────────────────────────────────────────────────────────────────────

(async () => {
  console.log('\n══════════════════════════════════════════════════════════');
  console.log('          EDYC — Pending Membership Invitations');
  console.log('══════════════════════════════════════════════════════════\n');

  const username = await ask('  👤  Username / Email: ');
  const password = await ask('  🔒  Password: ', true);
  rl.close();

  console.log('\n  🔐  Authenticating …');

  let token;
  try {
    token = await login(username, password);
  } catch (err) {
    console.error(`\n  ❌  ${err.message}\n`);
    process.exit(1);
  }

  console.log('  📋  Fetching pending invitations …\n');

  let invitations;
  try {
    invitations = await fetchAllPendingInvitations(token);
  } catch (err) {
    console.error(`\n  ❌  ${err.message}\n`);
    process.exit(1);
  }

  printTable(invitations);
})();

