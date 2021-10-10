const XLSX = require('xlsx');

const workbook = XLSX.readFile('ex.xlsx');
const sheet = workbook.SheetNames;
console.log(sheet);
console.log('\n');
console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet[0]]));
