#!/usr/bin/env node
// Bakes a spreadsheet into src/data/lastUpload.json, which the app bundles
// as the default "Toyota" replay so it's available on any computer that has
// this repo — no backend/account, but it means re-running this + committing
// + pulling on other machines is how the default gets updated.
//
// Usage: npm run bake-upload -- "/path/to/sheet.xlsx"

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { buildInitialMap } = require('../src/fieldAliases');

const filePath = process.argv[2];
if (!filePath) {
  console.error('Usage: npm run bake-upload -- "/path/to/sheet.xlsx"');
  process.exit(1);
}

const wb = XLSX.readFile(filePath);
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
if (!data || data.length < 1) {
  console.error('Sheet appears empty.');
  process.exit(1);
}

const headers  = data[0].map((h) => String(h));
const dataRows = data.slice(1).filter((r) => r.some((c) => c !== ''));
const fieldMap = buildInitialMap(headers);

const output = {
  fileName: path.basename(filePath),
  rawHeaders: headers,
  rawRows: dataRows,
  fieldMap,
  bakedAt: new Date().toISOString(),
};

const outPath = path.join(__dirname, '..', 'src', 'data', 'lastUpload.json');
fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, JSON.stringify(output, null, 2));

console.log(`Baked ${dataRows.length} rows from "${output.fileName}" into ${path.relative(process.cwd(), outPath)}`);
console.log('Commit + push this file (and have other computers pull) to update the shared "Toyota" default.');
