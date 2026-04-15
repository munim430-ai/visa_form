const express = require('express');
const session = require('express-session');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const os = require('os');

const { parseExcel } = require('./src/excel');
const { generateForms } = require('./src/generate');


const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Paths — works both inside Electron (packaged/dev) and plain Node (tests)
const isPackaged = !process.defaultApp; // undefined (packaged) → true; true (dev) → false

const resourcesPath = (isPackaged && process.resourcesPath)
  ? process.resourcesPath
  : __dirname;

const TEMPLATES_DIR = path.join(resourcesPath, 'templates');
const ASSETS_DIR    = path.join(resourcesPath, 'assets');
const PHOTOS_DIR    = (isPackaged && process.execPath)
  ? path.join(path.dirname(process.execPath), 'photos')
  : path.join(__dirname, 'photos');
const OUTPUT_DIR = path.join(os.homedir(), 'Documents', 'Korea Visa Forms');

const PASSWORD = 'visaform2024';

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: 'kv-secret-2024',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }
}));

// Serve renderer
app.use(express.static(path.join(__dirname, 'renderer')));

// Auth middleware
function requireAuth(req, res, next) {
  if (req.session && req.session.authenticated) return next();
  res.status(401).json({ error: 'Unauthorized' });
}

// --- Auth routes ---
app.post('/login', (req, res) => {
  const { password } = req.body;
  if (password === PASSWORD) {
    req.session.authenticated = true;
    res.json({ ok: true });
  } else {
    res.status(401).json({ error: 'Invalid password' });
  }
});

app.get('/logout', (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// --- Excel template download (single blank template) ---
app.get('/excel-template', requireAuth, (req, res) => {
  const baseFile = path.join(ASSETS_DIR, 'visa_template.xlsx');
  if (!fs.existsSync(baseFile)) {
    return res.status(404).json({ error: 'Template file not found' });
  }
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="visa_template.xlsx"');
  res.sendFile(path.resolve(baseFile));
});

// --- Generate forms ---
app.post('/generate', requireAuth, upload.single('excel'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No Excel file uploaded' });

    const university = req.body.university || 'Unknown';
    const templatePath    = path.join(TEMPLATES_DIR, 'visa.docx');
    const nocTemplatePath = path.join(TEMPLATES_DIR, 'noc.docx');

    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: 'Visa template not found. Please ensure templates/visa.docx exists.' });
    }
    if (!fs.existsSync(nocTemplatePath)) {
      return res.status(500).json({ error: 'NOC template not found. Please ensure templates/noc.docx exists.' });
    }

    // Parse Excel
    const students = parseExcel(req.file.buffer);
    const valid = students.filter(s => s._valid);
    const invalid = students.filter(s => !s._valid);

    if (valid.length === 0) {
      return res.status(400).json({
        error: 'No valid rows found in Excel file',
        invalidRows: invalid.map(s => ({
          row: s._rowIndex,
          missing: s._missingFields
        }))
      });
    }

    // Ensure output dir exists
    if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

    // Generate
    const results = await generateForms({
      students: valid,
      templatePath,
      nocTemplatePath,
      photosDir: PHOTOS_DIR,
      outputDir: OUTPUT_DIR,
      university
    });

    res.json({
      ok: true,
      generated: results.success,
      failed: results.failed,
      outputDir: OUTPUT_DIR,
      invalidRows: invalid.map(s => ({
        row: s._rowIndex,
        missing: s._missingFields
      }))
    });

  } catch (err) {
    console.error('Generate error:', err);
    res.status(500).json({ error: err.message });
  }
});

// --- Download generated file ---
app.get('/download/:filename', requireAuth, (req, res) => {
  const safeName = path.basename(req.params.filename);
  const filePath = path.join(OUTPUT_DIR, safeName);
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found' });
  }
  res.download(filePath);
});

module.exports = app;
