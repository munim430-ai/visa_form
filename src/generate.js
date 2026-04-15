const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

// Photo dimensions: 3.5 cm × 4.5 cm in EMUs (1 cm = 360000 EMU)
const PHOTO_CX = 1260000; // 3.5 cm
const PHOTO_CY = 1620000; // 4.5 cm

// ---------------------------------------------------------------------------
// Marker replacement — sort by key length descending so longer markers
// (e.g. VISAUNIVADDR) are replaced before their prefixes (VISAUNIV).
// ---------------------------------------------------------------------------
function replaceMarkers(xml, markers) {
  let result = xml;
  const entries = Object.entries(markers)
    .sort((a, b) => b[0].length - a[0].length);
  for (const [marker, value] of entries) {
    result = result.split(marker).join(value || '');
  }
  return result;
}

// ---------------------------------------------------------------------------
// Education checkbox line builder
// ---------------------------------------------------------------------------
function buildEducationLine(edu) {
  const choices = [
    ['Master',      "석사/박사 Master's/Doctoral Degree"],
    ['Bachelor',    "대졸 Bachelor's Degree"],
    ['High School', '고졸 High School Diploma'],
    ['Others',      '기타 Other'],
  ];
  return choices.map(([key, label]) =>
    label + (edu === key ? ' [ \u2713 ]' : ' []')
  ).join('');
}

// ---------------------------------------------------------------------------
// Photo helpers
// ---------------------------------------------------------------------------
function findPhotoFile(photosDir, familyName, givenName) {
  if (!photosDir || !fs.existsSync(photosDir)) return null;
  const fullName = `${familyName}_${givenName}`.replace(/\s+/g, '_');
  const candidates = [
    fullName,
    familyName,
    `${familyName} ${givenName}`,
  ];
  const exts = ['.jpg', '.jpeg', '.png'];
  try {
    const files = fs.readdirSync(photosDir);
    for (const name of candidates) {
      for (const ext of exts) {
        const target = (name + ext).toLowerCase();
        const match = files.find(f => f.toLowerCase() === target);
        if (match) return path.join(photosDir, match);
      }
    }
  } catch (_) {}
  return null;
}

function buildPhotoRel(relId) {
  return `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/photo_student.jpg"/>`;
}

function buildDrawingXml(relId) {
  return (
    `<w:p><w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
    `<wp:extent cx="${PHOTO_CX}" cy="${PHOTO_CY}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="100" name="StudentPhoto"/>` +
    `<wp:cNvGraphicFramePr>` +
    `<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>` +
    `</wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="0" name="StudentPhoto"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill>` +
    `<a:blip r:embed="${relId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
    `<a:stretch><a:fillRect/></a:stretch>` +
    `</pic:blipFill>` +
    `<pic:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${PHOTO_CX}" cy="${PHOTO_CY}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r></w:p>`
  );
}

async function injectPhoto(zip, photoPath, docXml) {
  const photoBuffer = fs.readFileSync(photoPath);
  zip.file('word/media/photo_student.jpg', photoBuffer);

  const relsPath = 'word/_rels/document.xml.rels';
  let relsXml = await zip.file(relsPath).async('string');
  const relId = 'rIdPhoto100';

  if (!relsXml.includes(relId)) {
    relsXml = relsXml.replace('</Relationships>', buildPhotoRel(relId) + '</Relationships>');
    zip.file(relsPath, relsXml);
  }

  // Replace the VISAPHOTO marker paragraph with the drawing XML
  const photoParaRegex = /<w:p\b[^>]*>(?:(?!<\/w:p>).)*?VISAPHOTO(?:(?!<\/w:p>).)*?<\/w:p>/s;
  docXml = docXml.replace(photoParaRegex, buildDrawingXml(relId));
  return docXml;
}

// ---------------------------------------------------------------------------
// Filename sanitisation
// ---------------------------------------------------------------------------
function safeFilename(str) {
  return str
    .replace(/[^a-zA-Z0-9_\- ]/g, '')
    .replace(/\s+/g, '_')
    .slice(0, 40)
    .replace(/^_+|_+$/g, '');
}

// ---------------------------------------------------------------------------
// Build the contact string: "Phone: xxx  |  Email: yyy"
// ---------------------------------------------------------------------------
function buildContact(student) {
  const parts = [];
  if (student.phone)  parts.push(student.phone);
  if (student.email)  parts.push(student.email);
  return parts.join('  |  ');
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------
/**
 * @param {Object} opts
 * @param {Array}  opts.students     Parsed student objects from parseExcel()
 * @param {string} opts.templatePath Absolute path to templates/visa.docx
 * @param {string} opts.nocTemplatePath Absolute path to templates/noc.docx
 * @param {string} opts.photosDir    Directory for student photos
 * @param {string} opts.outputDir    Output directory for generated DOCX files
 * @param {string} opts.university   Institution name (used as fallback for VISAUNIV)
 */
async function generateForms({ students, templatePath, nocTemplatePath, photosDir, outputDir, university }) {
  const templateBuffer    = fs.readFileSync(templatePath);
  const nocTemplateBuffer = fs.readFileSync(nocTemplatePath);
  const success = [];
  const failed  = [];

  for (let i = 0; i < students.length; i++) {
    const s   = students[i];
    const seq = String(i + 1).padStart(3, '0');

    // ------------------------------------------------------------------
    // Visa Application DOCX
    // ------------------------------------------------------------------
    try {
      const zip    = await JSZip.loadAsync(templateBuffer);
      let docXml   = await zip.file('word/document.xml').async('string');

      // Build marker → value map (must match FIELD_MAP in tag_template.py)
      const markers = {
        // Personal
        VISAFAMILYNAME:    s.family_name,
        VISAGIVENNAME:     s.given_name,
        VISADOB:           [s.birth_year, s.birth_month, s.birth_day].filter(Boolean).join('-') || '',
        VISANATIONALID:    s.national_id        || '',

        // 1.3 Gender checkboxes
        VISAGENDERMALE:    s.gender === 'M' ? '\u2713' : '',
        VISAGENDERFEMALE:  s.gender === 'F' ? '\u2713' : '',

        // 2.2 Status of Stay
        VISASTAYTYPE:      s.status_of_stay     || '',

        // Passport
        VISAPASSNUM:       s.passport_number,
        VISAPASSISSUEPLACE: s.passport_issue_place || 'DIP/DHAKA',
        VISAPASSISSUEDATE: s.passport_issue_date   || '',
        VISAPASSEXPIRY:    s.passport_expiry_date  || '',

        // Contact
        VISAHOMEADDR:      s.home_address       || '',
        VISACONTACT:       (s.phone || '').replace(/^\+/, ''),
        VISAEMAIL:         s.email              || '',
        VISAEMERGNAME:     s.emergency_contact_name         || '',
        VISAEMERGREL:      s.emergency_contact_relationship || '',

        // Education
        VISASCHOOLNAME:    s.school_name     || '',
        VISASCHOOLLOC:     s.school_location || '',
        // 6.1 Education checkboxes — ✓ if selected, space if not (gives [ ] not [])
        VISAEDUMASTER: s.education === 'Master'      ? '\u2713' : ' ',
        VISAEDUBACH:   s.education === 'Bachelor'    ? '\u2713' : ' ',
        VISAEDUHIGH:   s.education === 'High School' ? '\u2713' : ' ',
        VISAEDUOTHER:  s.education === 'Others'      ? '\u2713' : ' ',

        // University / Study
        VISAUNIV:          s.university_name    || university,
        VISACOURSE:        s.course_name        || '',
        VISAUNIVADDR:      s.university_address || '',
        VISAUNIVPHONE:     s.university_phone   || '',

        // Visit
        VISASTAY:          '365 DAYS',
        VISAENTRYDATE:     s.entry_date         || '',
        VISAKOREAADDR:     s.korea_address      || '',
        VISAKOREATELL:     s.korea_contact      || '',

        // Inviter
        VISAINVNAME:       s.inviter_name       || s.university_name || university,
        VISAINVREG:        s.inviter_reg_no     || '',
        VISAINVADDR:       s.inviter_address    || '',
        VISAINVPHONE:      s.inviter_phone      || '',

        // Funding / Sponsor
        VISACOSTS:         s.travel_costs       || '',
        VISASPONSORNAME:   s.sponsor_name         || '',
        VISASPONSORREL:    s.sponsor_relationship || '',
        VISASPONSORPHONE:  (s.sponsor_contact || '').replace(/^\+/, ''),

        // Section 11 — static
        VISASEC11NAME:     'HANGEUL KOREAN LANGUAGE AND VISA',
        VISASEC11PHONE:    '821071571779',
        VISASEC11REL:      'AGENCY',

        // Declaration
        VISASIGNATURE:     `${s.family_name} ${s.given_name}`,

        // Photo (handled separately below)
        VISAPHOTO:         '',
      };

      docXml = replaceMarkers(docXml, markers);

      // Photo injection
      let photoPath = null;
      if (s.photo_path && fs.existsSync(s.photo_path)) {
        photoPath = s.photo_path;
      } else {
        photoPath = findPhotoFile(photosDir, s.family_name, s.given_name);
      }
      if (photoPath) {
        docXml = await injectPhoto(zip, photoPath, docXml);
      }

      zip.file('word/document.xml', docXml);

      const visaName   = `STU${seq}_${safeFilename(s.family_name)}_${safeFilename(s.given_name)}.docx`;
      const outBuffer  = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      fs.writeFileSync(path.join(outputDir, visaName), outBuffer);

      // ------------------------------------------------------------------
      // NOC (Parents No Objection Certificate) DOCX
      // ------------------------------------------------------------------
      const isFemale   = (s.gender || '').toUpperCase() === 'F';
      const hisHer     = isFemale ? 'her'      : 'his';
      const himHer     = isFemale ? 'her'      : 'him';
      const sonDaughter = isFemale ? 'daughter' : 'son';
      const studentName = `${s.family_name} ${s.given_name}`;
      const fatherName  = s.noc_father_name || s.sponsor_name || '';
      const motherName  = s.noc_mother_name || '';
      const nocPhone    = (s.sponsor_contact || '').replace(/^\+/, '');

      const nocZip = await JSZip.loadAsync(nocTemplateBuffer);
      let nocXml   = await nocZip.file('word/document.xml').async('string');

      const nocMarkers = {
        NOCUNIV:       s.university_name || university,
        NOCSTUDENT:    studentName,
        NOCFATHERNAME: fatherName,
        NOCMOTHERNAME: motherName,
        NOCADDR:       s.home_address   || '',
        NOCPHONE:      nocPhone,
        NOCHISHER:     hisHer,
        NOCHIMHER:     himHer,
        NOCSON:        sonDaughter,
      };

      nocXml = replaceMarkers(nocXml, nocMarkers);
      nocZip.file('word/document.xml', nocXml);

      const nocName   = `NOC${seq}_${safeFilename(s.family_name)}_${safeFilename(s.given_name)}.docx`;
      const nocBuffer = await nocZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      fs.writeFileSync(path.join(outputDir, nocName), nocBuffer);

      success.push({
        row:     s._rowIndex,
        name:    studentName,
        file:    visaName,
        nocFile: nocName,
      });

    } catch (err) {
      failed.push({ row: s._rowIndex, name: `${s.family_name} ${s.given_name}`, error: err.message });
    }
  }

  return { success, failed };
}

module.exports = { generateForms };
