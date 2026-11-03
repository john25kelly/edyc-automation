const XLSX = require('xlsx');
const wb = XLSX.readFile('members.xlsx');
console.log('Sheet names:', wb.SheetNames);
const sheet = wb.Sheets['APRIL 26'];
if (!sheet) { console.log('Sheet APRIL 26 not found'); process.exit(1); }
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
console.log('Total rows:', rows.length);
console.log('Row 1 (headers):', rows[0]);
console.log('Row 2 (first data):', rows[1]);
console.log('Row 3:', rows[2]);
console.log('Row 4:', rows[3]);

