const express = require('express');
const XLSX    = require('xlsx');
const multer  = require('multer');
const path    = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Column definitions ────────────────────────────────────────────────────────

const STD_COLS = [
  'ID', 'Standard Number', 'Title', 'Scope/Abstract', 'Comment',
  'Web Links', 'First Year', 'Current Year', 'ICS', 'Type',
  'Is CEN?', 'CEN Committee', 'Is ISO?', 'ISO Committee',
  'Degree Centrality', 'In-Degree Centrality', 'Out-Degree Centrality', 'Eigencentrality'
];

const REL_COLS = [
  'Id', 'Source Standard Number', 'Target Current Number', 'Type', 'Link'
];

// ── Helpers ───────────────────────────────────────────────────────────────────

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

function parseWorkbook(wb) {
  const standards     = parseSheet(wb, 'GIS-Standards',     STD_COLS);
  const relationships = parseSheet(wb, 'GIS-Relationships', REL_COLS);
  return { standards, relationships, lastUpdated: new Date().toISOString() };
}

// ── Initial load ──────────────────────────────────────────────────────────────

const EXCEL_PATH = path.join(__dirname, 'data', 'GIS-standards-landscape.xlsx');

let data = { standards: [], relationships: [], lastUpdated: null };

try {
  const wb = XLSX.readFile(EXCEL_PATH, { cellFormula: false, cellStyles: false });
  data = parseWorkbook(wb);
  console.log(`✓  ${data.standards.length} standards, ${data.relationships.length} relationships loaded`);
} catch (err) {
  console.error('✗  Could not read Excel file:', err.message);
}

// ── Admin auth ────────────────────────────────────────────────────────────────

const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'changeme';

function requireAdmin(req, res, next) {
  const pwd = req.headers['x-admin-password'] || req.body?.password;
  if (pwd !== ADMIN_PASSWORD) return res.status(401).json({ error: 'Unauthorised' });
  next();
}

// ── File upload ───────────────────────────────────────────────────────────────

const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 20 * 1024 * 1024 },   // 20 MB cap
  fileFilter: (_req, file, cb) => {
    const ok = file.originalname.toLowerCase().endsWith('.xlsx');
    cb(ok ? null : new Error('Only .xlsx files are accepted'), ok);
  },
});

// ── Routes ────────────────────────────────────────────────────────────────────

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Public data API
app.get('/api/standards',     (_req, res) => res.json(data.standards));
app.get('/api/relationships', (_req, res) => res.json(data.relationships));
app.get('/api/status', (_req, res) => res.json({
  standards:     data.standards.length,
  relationships: data.relationships.length,
  lastUpdated:   data.lastUpdated,
}));

// Admin page
app.get('/admin', (_req, res) =>
  res.sendFile(path.join(__dirname, 'public', 'admin.html'))
);

// Excel upload → live reload (no restart needed)
app.post('/admin/upload',
  upload.single('excel'),
  requireAdmin,
  (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    try {
      const wb  = XLSX.read(req.file.buffer, { cellFormula: false, cellStyles: false });
      data      = parseWorkbook(wb);
      console.log(`↺  Data refreshed: ${data.standards.length} standards, ${data.relationships.length} relationships`);
      res.json({
        success:       true,
        standards:     data.standards.length,
        relationships: data.relationships.length,
        lastUpdated:   data.lastUpdated,
      });
    } catch (err) {
      res.status(400).json({ error: err.message });
    }
  }
);

// ── Start ─────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`\n  GIS Standards Landscape  →  http://localhost:${PORT}`);
  console.log(`  Admin panel              →  http://localhost:${PORT}/admin\n`);
});
