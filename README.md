# Visa Form

**Korea D-4-1 Student Visa Application Form Generator**

Batch-generate filled Korean D-4 visa application DOCX files and Parents' No Objection Certificates (NOC) from a single Excel spreadsheet. Built for visa agencies and language schools processing multiple student applications.

---

## Features

- Upload one Excel file with one row per student
- Generates a filled **Visa Application Form** (DOCX) per student
- Generates a **Parents' No Objection Certificate** (NOC DOCX) per student
- Automatic photo embedding (JPG/PNG) into visa form
- Gender-aware pronoun handling in NOC (his/him/son vs her/her/daughter)
- Education level checkboxes auto-filled
- Separate download button for each file

---

## Requirements

- Windows 10 or later (64-bit)
- No installation required — portable EXE

---

## How to Use

1. **Download** `KoreaVisaForm.exe` from the [Releases](https://github.com/munim430-ai/visa_form/releases) page
2. **Run** the EXE — the app opens in a window (no browser needed)
3. **Log in** with the agency password
4. **Download** the Excel template from inside the app
5. **Fill** one row per student in the Excel file and save as `.xlsx`
6. **Upload** the Excel file in the app
7. **Click Generate** — two DOCX files are created per student
8. **Download** each file with the individual download buttons

---

## Excel Columns (A – AN)

| Col | Field | Notes |
|-----|-------|-------|
| A | Family Name | English, uppercase |
| B | Given Name | English, uppercase |
| C | Korean Name | Optional (漢字) |
| D | Birth Year | YYYY |
| E | Birth Month | MM |
| F | Birth Day | DD |
| G | Gender | M or F |
| H | National ID | Passport-page ID number |
| I | Passport Number | |
| J | Passport Place of Issue | e.g. DIP/DHAKA |
| K | Passport Issue Date | YYYY-MM-DD |
| L | Passport Expiry Date | YYYY-MM-DD |
| M | University / Institution in Korea | |
| N | Course / Program Name | |
| O | University Address | Korean address |
| P | University Phone | |
| Q | Intended Entry Date | YYYY/MM/DD |
| R | Address in Korea | Dormitory / residence |
| S | Contact No. in Korea | |
| T | Inviting Organisation Name | Usually same as university |
| U | Inviter Business Reg. No. | |
| V | Inviter Address | |
| W | Inviter Phone | |
| X | Estimated Travel Costs | e.g. 22,000$ |
| Y | Home Country Address | |
| Z | Current Address | If different from home |
| AA | Applicant Cell Phone | |
| AB | Applicant Email | |
| AC | Emergency Contact Full Name | |
| AD | Emergency Contact Relationship | |
| AE | Photo File Path | Full path to JPG/PNG (optional) |
| AF | Status of Stay | e.g. D-4-1 |
| AG | Education | Master / Bachelor / High School / Others |
| AH | Sponsor / Payer Name | |
| AI | Sponsor Relationship | |
| AJ | Sponsor Contact No. | |
| AK | School Name | Previous school (section 6.2) |
| AL | School Location | City/Country |
| AM | NOC — Father Full Name | Used in NOC document |
| AN | NOC — Mother Full Name | Used in NOC document |

---

## Static Fields (hardcoded — no Excel column needed)

| Field | Value |
|-------|-------|
| Period of Stay | 365 DAYS |
| Passport Type | Regular ✓ |
| Marital Status | Single ✓ |
| Occupation | Student ✓ |
| Nationality | BANGLADESHI |
| Country of Birth | BANGLADESH |
| Relationship to Inviter | ADMITTED STUDENT |
| Section 11 Assistance | HANGEUL KOREAN LANGUAGE AND VISA / AGENCY |

---

## Photo Lookup

Photos are matched automatically from a `photos/` folder next to the EXE:

1. `FAMILYNAME_GIVENNAME.jpg` (or `.jpeg`, `.png`)
2. `FAMILYNAME.jpg`

Or provide a full absolute path in Excel column **AE**.

---

## Output Files

Generated in `Documents\Korea Visa Forms\`:

- `STU001_FAMILYNAME_GIVENNAME.docx` — Visa application form
- `NOC001_FAMILYNAME_GIVENNAME.docx` — Parents' No Objection Certificate

---

## Password

Default login password: `visaform2024`

---

## Credits

**Developer:** Hasibul Munim
**Motivation:** Farhabi Shikder

© 2024 Hasibul Munim. All rights reserved.
