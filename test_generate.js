/**
 * test_generate.js
 * Quick headless smoke-test: parse the sample Excel and generate one DOCX.
 * Run with:  node test_generate.js
 */
const path = require('path');
const fs   = require('fs');
const os   = require('os');

const { parseExcel }   = require('./src/excel');
const { generateForms } = require('./src/generate');

const TEMPLATE   = path.join(__dirname, 'templates', 'visa.docx');
const XLSX_FILE  = path.join(__dirname, 'assets', 'visa_template.xlsx');
const PHOTOS_DIR = path.join(__dirname, 'photos');
const OUTPUT_DIR = path.join(os.homedir(), 'Documents', 'Korea Visa Forms');

async function main() {
  console.log('=== Visa Form Docx — Generation Test ===\n');

  // 1. Check template
  if (!fs.existsSync(TEMPLATE)) {
    console.error('ERROR: templates/visa.docx not found. Run tag_template.py first.');
    process.exit(1);
  }
  console.log('✓ Template found:', TEMPLATE);

  // 2. Parse Excel
  if (!fs.existsSync(XLSX_FILE)) {
    console.error('ERROR: assets/visa_template.xlsx not found.');
    process.exit(1);
  }
  const buffer   = fs.readFileSync(XLSX_FILE);
  const students = parseExcel(buffer);

  console.log(`✓ Excel parsed: ${students.length} total row(s)`);
  const valid   = students.filter(s => s._valid);
  const invalid = students.filter(s => !s._valid);
  console.log(`  Valid: ${valid.length}  |  Invalid: ${invalid.length}`);
  if (invalid.length) {
    invalid.forEach(s => console.log(`  ✗ Row ${s._rowIndex}: missing ${s._missingFields.join(', ')}`));
  }
  if (!valid.length) {
    console.error('No valid rows — aborting.');
    process.exit(1);
  }

  // Print first student fields for inspection
  console.log('\nFirst student:');
  const s = valid[0];
  Object.entries(s)
    .filter(([k]) => !k.startsWith('_'))
    .forEach(([k, v]) => v && console.log(`  ${k.padEnd(24)} = ${v}`));

  // 3. Ensure output dir
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  console.log(`\n✓ Output dir: ${OUTPUT_DIR}`);

  // 4. Generate
  console.log('\nGenerating DOCX files...');
  const results = await generateForms({
    students: valid,
    templatePath: TEMPLATE,
    nocTemplatePath: path.join(__dirname, 'templates', 'noc.docx'),
    photosDir: PHOTOS_DIR,
    outputDir: OUTPUT_DIR,
    university: 'Test University',
  });

  console.log(`\n=== Results ===`);
  results.success.forEach(r =>
    console.log(`  ✓ Row ${r.row}  ${r.name}  →  ${r.file}`)
  );
  results.failed.forEach(r =>
    console.error(`  ✗ Row ${r.row}  ${r.name}  ERROR: ${r.error}`)
  );

  console.log(`\nDone. Generated ${results.success.length} / ${valid.length} file(s).`);

  // 5. Spot-check: verify markers were actually replaced in output
  if (results.success.length > 0) {
    const JSZip = require('jszip');
    const outPath = path.join(OUTPUT_DIR, results.success[0].file);
    const buf = fs.readFileSync(outPath);
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('word/document.xml').async('string');

    const MARKERS = [
      'VISAFAMILYNAME','VISAGIVENNAME','VISAPASSNUM','VISAUNIV',
      'VISACOURSE','VISASTAY','VISAENTRYDATE','VISAAPPYEAR'
    ];
    const remaining = MARKERS.filter(m => xml.includes(m));
    if (remaining.length === 0) {
      console.log('\n✓ Marker check passed — no raw markers left in output DOCX.');
    } else {
      console.warn('\n⚠ Unreplaced markers found in output:', remaining.join(', '));
    }
  }
}

main().catch(err => {
  console.error('\nFATAL:', err);
  process.exit(1);
});
