const XLSX = require('xlsx');

// ---------------------------------------------------------------------------
// COLUMNS — one entry per Excel column, in order A → AH
// Column letter must match create_excel_template.py exactly.
// ---------------------------------------------------------------------------
const COLUMNS = [
  { col: 'A',  key: 'family_name' },           // Family Name (English)
  { col: 'B',  key: 'given_name' },            // Given Names (English)
  { col: 'C',  key: 'korean_name' },           // 漢字姓名 / Korean characters (optional)
  { col: 'D',  key: 'birth_year' },            // Birth Year  (YYYY)
  { col: 'E',  key: 'birth_month' },           // Birth Month (MM)
  { col: 'F',  key: 'birth_day' },             // Birth Day   (DD)
  { col: 'G',  key: 'gender' },                // M or F
  { col: 'H',  key: 'national_id' },           // National Identity No.
  { col: 'I',  key: 'passport_number' },       // Passport Number
  { col: 'J',  key: 'passport_issue_place' },  // Place of Issue
  { col: 'K',  key: 'passport_issue_date' },   // Issue Date (YYYY-MM-DD or YYYY/MM/DD)
  { col: 'L',  key: 'passport_expiry_date' },  // Expiry Date
  { col: 'M',  key: 'university_name' },       // University / Institution name in Korea
  { col: 'N',  key: 'course_name' },           // Course / Program name
  { col: 'O',  key: 'university_address' },    // University address
  { col: 'P',  key: 'university_phone' },      // University phone
  // Q: stay_duration removed — hardcoded as "365 DAYS" in generate.js
  { col: 'Q',  key: 'entry_date' },            // Intended entry date (YYYY/MM/DD)
  { col: 'R',  key: 'korea_address' },         // Address in Korea
  { col: 'S',  key: 'korea_contact' },         // Contact no. in Korea
  { col: 'T',  key: 'inviter_name' },          // Inviting organisation name
  { col: 'U',  key: 'inviter_reg_no' },        // Business registration no. of inviter
  { col: 'V',  key: 'inviter_address' },       // Inviter address
  { col: 'W',  key: 'inviter_phone' },         // Inviter phone
  { col: 'X',  key: 'travel_costs' },          // Estimated travel costs (USD)
  { col: 'Y',  key: 'home_address' },          // Home country address
  { col: 'Z',  key: 'current_address' },       // Current residential address (if different)
  { col: 'AA', key: 'phone' },                 // Applicant cell phone no.
  { col: 'AB', key: 'email' },                 // Applicant email
  { col: 'AC', key: 'emergency_contact_name' },         // 4.5 Emergency contact full name
  { col: 'AD', key: 'emergency_contact_relationship' }, // 4.5 Emergency contact relationship
  { col: 'AE', key: 'photo_path' },            // Full path to photo file (optional)
  { col: 'AF', key: 'status_of_stay' },        // 2.2 Status of Stay (e.g. D-4-1)
  { col: 'AG', key: 'education' },             // 6.1 Education (Master/Bachelor/High School/Others)
  { col: 'AH', key: 'sponsor_name' },          // Section 10: sponsor/payer name
  { col: 'AI', key: 'sponsor_relationship' },  // Section 10: sponsor relationship to applicant
  { col: 'AJ', key: 'sponsor_contact' },       // Section 10: sponsor contact no.
  { col: 'AK', key: 'school_name' },          // 6.2 School/College Name
  { col: 'AL', key: 'school_location' },      // 6.3 School/College Location (city/country)
  // NOC-specific columns
  { col: 'AM', key: 'noc_father_name' },      // NOC: Father's full name
  { col: 'AN', key: 'noc_mother_name' },      // NOC: Mother's full name
];

const REQUIRED_FIELDS = ['family_name', 'given_name', 'passport_number'];

/**
 * Parse an Excel file buffer into an array of student objects.
 * @param {Buffer} buffer
 * @returns {Array<Object>}
 */
function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw: false,
  });

  if (rows.length < 2) return [];

  const students = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every(cell => cell === '' || cell == null)) continue;

    const student = { _rowIndex: i + 1, _valid: true, _missingFields: [] };

    COLUMNS.forEach(({ col, key }) => {
      const colIndex = colLetterToIndex(col);
      const value = row[colIndex] !== undefined ? String(row[colIndex]).trim() : '';
      student[key] = value;
    });

    // Validate required fields
    REQUIRED_FIELDS.forEach(field => {
      if (!student[field]) {
        student._valid = false;
        student._missingFields.push(field);
      }
    });

    students.push(student);
  }

  return students;
}

function colLetterToIndex(col) {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1;
}

module.exports = { parseExcel, COLUMNS };
