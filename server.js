const express  = require('express');
const cors     = require('cors');
const fs       = require('fs');
const path     = require('path');
const multer   = require('multer');
const XLSX     = require('xlsx');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Paths ──────────────────────────────────────────
const DATA_DIR    = path.join(__dirname, 'data');
const DATA_FILE   = path.join(DATA_DIR, 'sales_data.json');
const UPLOAD_DIR  = path.join(DATA_DIR, 'uploads');
const PUBLIC_DIR  = path.join(__dirname, 'public');

fs.mkdirSync(DATA_DIR,   { recursive: true });
fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// ── Middleware ────────────────────────────────────
app.use(cors());
app.use(express.json({ limit: '20mb' }));
app.use(express.static(PUBLIC_DIR));

const upload = multer({ dest: UPLOAD_DIR, limits: { fileSize: 10 * 1024 * 1024 } });

// ── Helpers ───────────────────────────────────────
function readStore() {
  try {
    if (fs.existsSync(DATA_FILE))
      return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  } catch(e) {}
  return { data: [], updatedAt: null, source: null };
}

function writeStore(payload) {
  payload.updatedAt = new Date().toISOString();
  fs.writeFileSync(DATA_FILE, JSON.stringify(payload), 'utf8');
  return payload.updatedAt;
}

// ── Column detection (matches frontend logic) ────
const COL_ALIASES = {
  region:   ['region','city','state','area'],
  promoter: ['promoter name','promoter','name','promoters name'],
  store:    ['store','store name','outlet'],
  zone:     ['zone'],
  emp_code: ['emp code','empl code','jwplcode','jwpl code','employee code'],
  dsr_id:   ['dsr id','dsr','dsrid'],
  c_target: ['canola target','c.target','c target','canola_target'],
  c_sale:   ['canola sale','c.sale','c sale','canola_sale','canola'],
  o_target: ['olive target','o.target','o target','olive_target'],
  o_sale:   ['olive sale','o.sale','o sale','olive_sale','olive'],
  so_name:  ['so name','soname','so'],
};

function findCol(headers, key) {
  const aliases = COL_ALIASES[key] || [];
  for (const a of aliases) {
    const idx = headers.findIndex(h => h.toLowerCase().trim() === a);
    if (idx >= 0) return idx;
  }
  return -1;
}

function sheetToRows(ws) {
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!json.length) return [];

  // Detect if this is GT MAY_Final format (region name in col0, then header row, then data)
  // vs simple format (single header row at top)
  const isGTFormat = json.slice(0, 5).some(row => {
    const first = String(row[0] || '').toLowerCase();
    return (first.includes('gujarat') || first.includes('haryana') ||
            first.includes('punjab') || first.includes('delhi') ||
            first.includes('mumbai') || first.includes('sale report'));
  });

  if (isGTFormat) {
    return parseGTFormat(json);
  }

  // Standard format — find single header row
  let hdrIdx = 0;
  for (let i = 0; i < Math.min(10, json.length); i++) {
    const row = json[i].map(c => String(c).toLowerCase());
    if (row.some(c => c.includes('promoter') || c.includes('canola'))) { hdrIdx = i; break; }
  }
  return parseStandardFormat(json, hdrIdx);
}

// Parse GT MAY_Final format: region row → header row → data rows → Total → blank → repeat
function parseGTFormat(json) {
  const num = v => { const n = parseFloat(String(v).replace(/[^\d.]/g,'')); return isNaN(n) ? 0 : n; };
  const rows = [];
  let currentRegion = '';
  let headers = null;
  let colMap = {};

  for (let i = 0; i < json.length; i++) {
    const raw = json[i];
    const first = String(raw[0] || '').trim();
    const firstLow = first.toLowerCase();

    // Blank row — reset headers so next region header is picked fresh
    if (raw.every(c => c === '' || c === null || c === undefined)) {
      headers = null;
      continue;
    }

    // Region header row — col0 has region name, rest empty
    const nonEmpty = raw.filter(c => c !== '' && c !== null && c !== undefined);
    if (nonEmpty.length <= 2 && (
      firstLow.includes('gujarat') || firstLow.includes('haryana') ||
      firstLow.includes('punjab') || firstLow.includes('delhi') ||
      firstLow.includes('mumbai') || firstLow.includes('kolkata') ||
      firstLow.includes('hyderabad') || firstLow.includes('bangalore') ||
      firstLow.includes('sale report') || firstLow.includes('report')
    )) {
      // Extract clean region name
      currentRegion = first
        .replace(/sale report/gi, '').replace(/report/gi, '')
        .trim().replace(/\s+/g, ' ');
      if (!currentRegion) currentRegion = first.trim();
      headers = null;
      continue;
    }

    // Header row — contains 'promoter' or 'canola'
    const rowLow = raw.map(c => String(c || '').toLowerCase());
    if (rowLow.some(c => c.includes('promoter') || c.includes('canola'))) {
      headers = raw.map(c => String(c || '').trim());
      // Build column map
      colMap = {};
      headers.forEach((h, idx) => {
        const hl = h.toLowerCase().trim();
        if (hl.includes('promoter') || hl === 'name') colMap.promoter = idx;
        else if (hl.includes('store') || hl.includes('outlet')) colMap.store = idx;
        else if (hl === 'zone' || hl === 'zone ') colMap.zone = idx;
        else if (hl.includes('emp') || hl.includes('jwpl')) colMap.emp_code = idx;
        else if (hl === 'dsr id' || hl === 'dsr') colMap.dsr_id = idx;
        else if (hl === 'canola target' || hl === 'c.target' || hl === 'c target') colMap.c_target = idx;
        else if (hl === 'canola' || hl === 'canola ' || hl === 'c.sale' || hl === 'c sale') colMap.c_sale = idx;
        else if (hl === 'olive target' || hl === 'o.target' || hl === 'o target') colMap.o_target = idx;
        else if (hl === 'olive' || hl === 'olive ' || hl === 'o.sale' || hl === 'o sale') colMap.o_sale = idx;
        else if (hl.includes('so name') || hl === 'so' || hl === 'soname') colMap.so_name = idx;
        else if (hl === 'dsr id' || hl.startsWith('dsr')) colMap.dsr_id = idx;
      });
      continue;
    }

    // Total / skip rows
    if (firstLow.startsWith('total') || firstLow === 's.no' || firstLow === 'sno') continue;

    // Data row — needs headers
    if (!headers || !currentRegion) continue;

    const g = key => colMap[key] !== undefined ? (raw[colMap[key]] ?? '') : '';
    const promoter = String(g('promoter')).trim();
    if (!promoter || promoter === '' || promoter.toLowerCase().startsWith('total')) continue;
    // Skip if looks like a number-only row (S.No column)
    if (/^\d+$/.test(promoter)) continue;

    rows.push({
      region:   currentRegion,
      promoter: promoter,
      store:    String(g('store')).trim(),
      zone:     String(g('zone')).trim(),
      emp_code: String(g('emp_code')).trim(),
      dsr_id:   String(g('dsr_id') !== undefined ? raw[colMap['dsr_id']] || '' : '').trim(),
      c_target: num(g('c_target')),
      c_sale:   num(g('c_sale')),
      o_target: num(g('o_target')),
      o_sale:   num(g('o_sale')),
      so_name:  String(g('so_name')).trim(),
    });
  }
  return rows;
}

// Standard format parser
function parseStandardFormat(json, hdrIdx) {
  const headers = json[hdrIdx].map(c => String(c).trim());
  const num = v => { const n = parseFloat(String(v).replace(/[^\d.]/g,'')); return isNaN(n) ? 0 : n; };
  const rows = [];
  let lastRegion = '';
  for (let i = hdrIdx + 1; i < json.length; i++) {
    const raw = json[i];
    if (raw.every(c => !c)) continue;
    const g = key => { const idx = findCol(headers, key); return idx >= 0 ? (raw[idx] ?? '') : ''; };
    const region = String(g('region') || lastRegion).trim();
    if (!region) continue;
    lastRegion = region;
    const promoter = String(g('promoter')).trim();
    if (!promoter || promoter.toLowerCase().startsWith('total')) continue;
    rows.push({
      region, promoter,
      store:    String(g('store')).trim(),
      zone:     String(g('zone')).trim(),
      emp_code: String(g('emp_code')).trim(),
      dsr_id:   String(g('dsr_id')).trim(),
      c_target: num(g('c_target')),
      c_sale:   num(g('c_sale')),
      o_target: num(g('o_target')),
      o_sale:   num(g('o_sale')),
      so_name:  String(g('so_name')).trim(),
    });
  }
  return rows;
}

// ── ROUTES ────────────────────────────────────────

// Ping — lightweight check for updates
app.get('/api/ping', (req, res) => {
  const s = readStore();
  res.json({ ok: true, updatedAt: s.updatedAt, count: s.data.length });
});

// GET all data
app.get('/api/data', (req, res) => {
  res.json(readStore());
});

// POST — save data (from frontend edits)
app.post('/api/data', (req, res) => {
  const { data, source } = req.body;
  if (!Array.isArray(data)) return res.status(400).json({ error: 'data must be array' });
  const store = readStore();
  const updatedAt = writeStore({ data, source: source || store.source });
  res.json({ ok: true, count: data.length, updatedAt });
});

// POST — upload Excel/CSV file
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file' });
  try {
    const ext = path.extname(req.file.originalname).toLowerCase();
    let rows = [];

    if (ext === '.csv') {
      const text = fs.readFileSync(req.file.path, 'utf8');
      const lines = text.split('\n').map(l => l.split(',').map(c => c.trim().replace(/^"|"$/g,'')));
      const headers = lines[0];
      rows = lines.slice(1).filter(l => l.some(c=>c)).map(raw => {
        const g = keys => { for(const k of keys){const i=headers.indexOf(k);if(i>=0)return raw[i]||'';} return ''; };
        const n = v => { const x=parseFloat(String(v||0).replace(/[^\d.]/g,'')); return isNaN(x)?0:x; };
        return {
          region:g(['region','Region']), promoter:g(['promoter name','Promoter Name']),
          store:g(['store','Store']), zone:g(['zone','Zone']),
          emp_code:g(['emp code','EMP CODE']), dsr_id:g(['dsr id','DSR ID']),
          c_target:n(g(['canola target','Canola Target'])), c_sale:n(g(['canola sale','Canola Sale'])),
          o_target:n(g(['olive target','Olive Target'])), o_sale:n(g(['olive sale','Olive Sale'])),
          so_name:g(['so name','SO NAME']),
        };
      });
    } else {
      const wb = XLSX.readFile(req.file.path);
      let sheetName = wb.SheetNames[0];
      for (const sn of wb.SheetNames) {
        const tmp = XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:''});
        const flat = tmp.slice(0,5).flat().map(c=>String(c).toLowerCase());
        if (flat.some(c=>c.includes('promoter')||c.includes('canola'))) { sheetName=sn; break; }
      }
      rows = sheetToRows(wb.Sheets[sheetName]);
    }

    fs.unlinkSync(req.file.path);
    if (!rows.length) return res.status(422).json({ error: 'No rows found' });

    const updatedAt = writeStore({ data: rows, source: req.file.originalname });
    res.json({ ok: true, count: rows.length, updatedAt, data: rows });
  } catch(e) {
    console.error(e);
    try { fs.unlinkSync(req.file.path); } catch(_) {}
    res.status(500).json({ error: e.message });
  }
});

// DELETE — clear all data
app.delete('/api/data', (req, res) => {
  writeStore({ data: [], source: null });
  res.json({ ok: true });
});

// GET summary by region
app.get('/api/summary', (req, res) => {
  const { data } = readStore();
  const regions = [...new Set(data.map(r=>r.region))];
  const summary = regions.map(region => {
    const rows = data.filter(r=>r.region===region);
    const ct=rows.reduce((s,r)=>s+r.c_target,0), cs=rows.reduce((s,r)=>s+r.c_sale,0);
    const ot=rows.reduce((s,r)=>s+r.o_target,0), os=rows.reduce((s,r)=>s+r.o_sale,0);
    return { region, promoters:rows.length, c_target:ct, c_sale:cs,
             c_pct: ct>0?+(cs/ct*100).toFixed(1):0,
             o_target:ot, o_sale:os, o_pct: ot>0?+(os/ot*100).toFixed(1):0 };
  });
  res.json({ summary, total: data.length });
});

// Catch-all → serve SPA
app.get('*', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`\n🟢  JIVO Dashboard  →  http://localhost:${PORT}`);
  console.log(`    Data: ${DATA_FILE}\n`);
});
