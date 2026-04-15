"""
tag_noc_template.py
-------------------
Reads  assets/NOC_SOURCE.docx  (a filled sample NOC letter)
Replaces real student/parent values with NOC* marker strings
Writes templates/noc.docx — the blank template used by generate.js

Markers injected
----------------
NOCUNIV        institution name (from university_name / col M)
NOCSTUDENT     student full name  (family_name + " " + given_name)
NOCFATHERNAME  father's full name (col AM)
NOCMOTHERNAME  mother's full name (col AN)
NOCADDR        home address (col Y — father's address)
NOCPHONE       sponsor contact phone (col AJ)
NOCHISHER      his / her  (gender-driven in generate.js)
NOCHIMHER      him / her
NOCSON         son / daughter
"""

import os, sys, shutil, zipfile
from lxml import etree

# ---------------------------------------------------------------------------
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR    = os.path.join(SCRIPT_DIR, '..')
SOURCE_DOCX = os.path.join(ROOT_DIR, 'assets', 'NOC_SOURCE.docx')
OUTPUT_DOCX = os.path.join(ROOT_DIR, 'templates', 'noc.docx')
TEMP_DIR    = os.path.join(ROOT_DIR, 'assets', '_noc_tmp')

W        = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14      = 'http://schemas.microsoft.com/office/word/2010/wordml'
XML_SPACE = 'http://www.w3.org/XML/1998/namespace'

# ---------------------------------------------------------------------------
# Exact run-text replacements  (key = exact text of a <w:t> node)
# ---------------------------------------------------------------------------
EXACT = {
    'ANWAR HOSSAIN KHANDAKAR,':    'NOCFATHERNAME,',
    'ANWAR HOSSAIN KHANDAKAR (FATHER)': 'NOCFATHERNAME (FATHER)',
    'MD TAMIM KHANDAKAR':          'NOCSTUDENT',
    'TANIA AKTER (MOTHER)':        'NOCMOTHERNAME (MOTHER)',
    'his':                         'NOCHISHER',
    'his ':                        'NOCHISHER ',
}

# Substring replacements applied to every <w:t> text value
SUBS = [
    # order matters — longer/more-specific first
    ("Subject: Parent's No Objection Letter for MOKWON UNIVERSITY.",
     "Subject: Parent's No Objection Letter for NOCUNIV."),
    ('MOKWON UNIVERSITY',  'NOCUNIV'),
    ('Mokwon University',  'NOCUNIV'),
    ('my son ',            'my NOCSON '),
    (' him ',              ' NOCHIMHER '),
    (' his ',              ' NOCHISHER '),
    ("AMODPUR,FARIDPUR,KULIARCHAR,NALBIDE-2340,KISHOREGANJ", 'NOCADDR'),
    ('+ 88016051178899',   'NOCPHONE'),
]

# ---------------------------------------------------------------------------
def process_xml(xml_bytes):
    tree  = etree.fromstring(xml_bytes)
    count = 0

    for t_node in tree.iter(f'{{{W}}}t'):
        text = t_node.text or ''
        if not text:
            continue

        new_text = text

        # Exact match first
        if new_text in EXACT:
            new_text = EXACT[new_text]
        else:
            # Substring replacements
            for old, new in SUBS:
                new_text = new_text.replace(old, new)

        if new_text != text:
            t_node.text = new_text
            t_node.set(f'{{{XML_SPACE}}}space', 'preserve')
            count += 1

    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True), count

# ---------------------------------------------------------------------------
def main():
    if not os.path.exists(SOURCE_DOCX):
        print(f'ERROR: source not found: {SOURCE_DOCX}')
        print('Copy your filled NOC sample to assets/NOC_SOURCE.docx first.')
        sys.exit(1)

    os.makedirs(os.path.join(ROOT_DIR, 'templates'), exist_ok=True)

    # Unpack
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)
    with zipfile.ZipFile(SOURCE_DOCX) as z:
        z.extractall(TEMP_DIR)

    doc_path = os.path.join(TEMP_DIR, 'word', 'document.xml')
    with open(doc_path, 'rb') as f:
        xml_bytes = f.read()

    xml_bytes, count = process_xml(xml_bytes)

    with open(doc_path, 'wb') as f:
        f.write(xml_bytes)

    # Repack
    tmp_out = OUTPUT_DOCX + '.tmp'
    with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(TEMP_DIR):
            for file in files:
                abs_path = os.path.join(root, file)
                arc_name = os.path.relpath(abs_path, TEMP_DIR)
                zout.write(abs_path, arc_name)

    if os.path.exists(OUTPUT_DOCX):
        try:
            os.remove(OUTPUT_DOCX)
        except OSError:
            pass
    os.replace(tmp_out, OUTPUT_DOCX)
    shutil.rmtree(TEMP_DIR, ignore_errors=True)

    print(f'Done. Tagged NOC template -> {OUTPUT_DOCX}')
    print(f'Text nodes replaced: {count}')

if __name__ == '__main__':
    main()
