const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const STD_COLS = [
  'ID', 'Standard Number', 'Title', 'Scope/Abstract', 'Comment',
  'Web Links', 'First Year', 'Current Year', 'ICS', 'Type',
  'Is CEN?', 'CEN Committee', 'Is ISO?', 'ISO Committee',
  'Degree Centrality', 'In-Degree Centrality', 'Out-Degree Centrality', 'Eigencentrality'
];

const REL_COLS = [
  'Id', 'Source Standard Number', 'Target Current Number', 'Type', 'Link'
];

function parseSheet(wb, sheetName, cols) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  return rows.map(row => {
    const out = {};
    cols.forEach(c => { out[c] = row[c] ?? null; });
    return out;
  });
}

const excelPath = path.join(__dirname, '..', 'data', 'GIS-standards-landscape.xlsx');
const wb = XLSX.readFile(excelPath, { cellFormula: false, cellStyles: false });

const standards     = parseSheet(wb, 'GIS-Standards',     STD_COLS);
const relationships = parseSheet(wb, 'GIS-Relationships', REL_COLS);

const outDir = path.join(__dirname, '..', 'public', 'data');
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, 'standards.json'),     JSON.stringify(standards));
fs.writeFileSync(path.join(outDir, 'relationships.json'), JSON.stringify(relationships));

console.log(`Generated: ${standards.length} standards, ${relationships.length} relationships`);
