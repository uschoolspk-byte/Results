const XLSX = require('xlsx');
const path = require('path');
const TOTAL_MARKS_PATH = path.join(__dirname, 'data', 'total.xlsx');

try {
  const wb = XLSX.readFile(TOTAL_MARKS_PATH);
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1});
  console.log('Headers:', rows[0]);
} catch (e) {
  console.error('Error:', e.message);
}
