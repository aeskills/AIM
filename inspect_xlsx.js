const fs = require('fs');
const path = require('path');

// Read the minified sheetjs library
const xlsxPath = path.join(__dirname, 'xlsx.full.min.js');
const xlsxContent = fs.readFileSync(xlsxPath, 'utf8');

// Evaluate sheetjs in a sandbox/global context
const sandbox = { module: { exports: {} } };
const initXLSX = new Function('module', xlsxContent);
initXLSX(sandbox.module);
const XLSX = sandbox.module.exports;

// Parse the excel file
const excelPath = path.join(__dirname, 'Implementation plan.xlsx');
const buf = fs.readFileSync(excelPath);
const workbook = XLSX.read(buf, { type: 'buffer' });

console.log('Sheet Names:', workbook.SheetNames);
for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    console.log(`Sheet "${sheetName}": ${data.length} rows`);
}
