# Korea D-4 Visa Form Generator — Developer Context

## What It Does
Electron v28 desktop app (Express backend, port 3721). Operator uploads an Excel file with one row per student → app generates filled Korean D-4 visa application DOCX files, one per student.

## Key Files
| File | Role |
|------|------|
| `main.js` | Electron entry point — launches Express server + BrowserWindow |
| `server.js` | Express routes: `/login`, `/generate`, `/download/:filename`, `/excel-template` |
| `src/excel.js` | Parses Excel buffer → array of student objects (cols A–AL, 38 columns) |
| `src/generate.js` | Takes student objects + template DOCX → fills markers → writes output DOCX |
| `scripts/tag_template.py` | One-time: reads `assets/AFRIN SHANJIDA HOSSAIN.docx`, inserts marker strings at paraIds, writes `templates/visa.docx` |
| `assets/create_excel_template.py` | Generates `assets/visa_template.xlsx` (blank input template for operators) |
| `renderer/index.html` | Single-page UI: login, file drop, generate button, download buttons |

## Tagging Pipeline
```
assets/AFRIN SHANJIDA HOSSAIN.docx   ← source (has w14:paraId attributes)
        │
        ▼  python scripts/tag_template.py
templates/visa.docx                  ← tagged template (has VISA* marker strings)
        │
        ▼  node src/generate.js (via server.js /generate)
Documents/Korea Visa Forms/STU001_NAME.docx  ← filled student forms
```

**Why AFRIN's doc?** The blank form (`VisaapplicationForm_EN.docx`) has zero `w14:paraId` attributes (old Word version). AFRIN's filled doc has them. All her personal data is overwritten by markers.

## Marker Modes (tag_template.py)
- `replace` — overwrites all `<w:t>` text in the paragraph
- `inject` — clears runs, adds a single run with the marker (for empty cells)
- `text_sub` — in-place substitution: `marker = 'OLD|||NEW'`
- `insert_after` — inserts a new paragraph after the target paraId

## Column Layout (A–AL, 38 columns)
A=family_name, B=given_name, C=korean_name, D=birth_year, E=birth_month, F=birth_day,
G=gender, H=national_id, I=passport_number, J=passport_issue_place, K=passport_issue_date,
L=passport_expiry_date, M=university_name, N=course_name, O=university_address, P=university_phone,
Q=entry_date, R=korea_address, S=korea_contact, T=inviter_name, U=inviter_reg_no,
V=inviter_address, W=inviter_phone, X=travel_costs, Y=home_address, Z=current_address,
AA=phone, AB=email, AC=emergency_contact_name, AD=emergency_contact_relationship,
AE=photo_path, AF=status_of_stay, AG=education, AH=sponsor_name, AI=sponsor_relationship,
AJ=sponsor_contact, AK=school_name, AL=school_location

## Static Fields (hardcoded, no Excel column)
- Period of Stay → `365 DAYS`
- Section 11 assistance → `HANGEUL KOREAN LANGUAGE AND VISA` / `AGENCY`
- Section 11.1 Yes/No → always `Yes [✓]`
- Section 9.1 Yes/No (has inviter) → always `Yes [✓]`

## Run / Build
```bash
npm start                            # dev: Electron + Express
python scripts/tag_template.py       # regenerate templates/visa.docx after form changes
python assets/create_excel_template.py  # regenerate Excel input template
node test_generate.js                # smoke-test: generates 2 sample docs
```

## Auth
- Password: `visaform2024`
- Session cookie (8 h), secret: `kv-secret-2024`

## Output
`C:\Users\HP\Documents\Korea Visa Forms\` — `STU001_FAMILYNAME_GIVENNAME.docx`

## Photos
Looked up in `photos/` dir by `FAMILYNAME_GIVENNAME.jpg` (or `.jpeg`, `.png`). Falls back to `FAMILYNAME.jpg`. Or set full path in Excel col AE.
