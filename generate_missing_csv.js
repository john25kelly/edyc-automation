/**
 * Converts missing_users.txt → missing_users.csv
 */

const fs   = require('fs');
const path = require('path');

const inputPath  = path.join(__dirname, 'missing_users.txt');
const outputPath = path.join(__dirname, 'missing_users.csv');

const lines = fs.readFileSync(inputPath, 'utf8').split('\n');

const csvRows = ['Duty,Date,Name,Invitation Sent'];

for (const raw of lines) {
  const line = raw.trim();
  if (!line) continue;

  // Date: DD/MM/YYYY
  const dateMatch = line.match(/\b(\d{2}\/\d{2}\/\d{4})\b/);
  if (!dateMatch) continue;
  const date = dateMatch[1];

  // Duty type: first [...] block
  const dutyMatch = line.match(/\[([^\]]+?)\]/);
  if (!dutyMatch) continue;
  const duty = dutyMatch[1].trim();

  // Name: text between the closing ] of duty type and ' — '
  const nameMatch = line.match(/\]\s{2,}(.+?)\s{2,}—/);
  if (!nameMatch) continue;
  const name = nameMatch[1].trim();

  // Invitation sent
  const inviteSent = line.includes('[invite pending]') ? 'Yes' : 'No';

  // Quote fields that might contain commas
  const quote = (s) => s.includes(',') ? `"${s}"` : s;
  csvRows.push(`${quote(duty)},${date},${quote(name)},${inviteSent}`);
}

fs.writeFileSync(outputPath, csvRows.join('\n') + '\n', 'utf8');
console.log(`✅  Written ${csvRows.length - 1} rows to: ${outputPath}`);

