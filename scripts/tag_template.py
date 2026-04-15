"""
tag_template.py
---------------
Reads assets/blank_visa.docx, injects/replaces marker strings into
specific paragraphs identified by their w14:paraId, and writes the
result to templates/visa.docx.

Usage:
    python scripts/tag_template.py

Prerequisites:
    pip install lxml
"""

import zipfile, shutil, os, sys, io, re
from lxml import etree

# Force UTF-8 stdout for Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BLANK_DOCX  = os.path.join(os.path.dirname(__file__), '..', 'assets', 'AFRIN SHANJIDA HOSSAIN.docx')
OUTPUT_DOCX = os.path.join(os.path.dirname(__file__), '..', 'templates', 'visa.docx')
TEMP_DIR    = os.path.join(os.path.dirname(__file__), '..', '_tmp_tag')

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
XML_SPACE = 'http://www.w3.org/XML/1998/namespace'

# ---------------------------------------------------------------------------
# FIELD MAP  — paraId values confirmed from structural analysis
# ---------------------------------------------------------------------------
# Modes:
#   'inject'       — clear all runs in existing empty <w:p>, add single run with marker
#   'replace'      — replace text content of existing paragraph (keep formatting)
#   'insert_after' — insert a brand-new paragraph immediately after the target paraId
#   'text_sub'     — in-place text substitution; marker = 'OLD_TEXT|||NEW_TEXT'
# ---------------------------------------------------------------------------
FIELD_MAP = [
    # --- 1. Personal Details ---
    # All paraIds verified against AFRIN SHANJIDA HOSSAIN.docx (same form version)
    ('261EA86B', 'VISAFAMILYNAME',     'replace'),   # Family name cell
    ('20A91AB6', 'VISAGIVENNAME',      'replace'),   # Given names cell
    # 1.2 漢字姓名 — intentionally left blank (no marker)
    # 1.3 Gender — single paragraph; markers sit inside the brackets
    ('3D23D906', '남성/Male[VISAGENDERMALE]  여성/Female[VISAGENDERFEMALE]', 'replace'),
    # 1.4 Date of Birth — single combined cell "YYYY-MM-DD"
    ('486C99AC', 'VISADOB',            'replace'),
    # 1.7 National Identity No.
    ('5130E3D2', 'VISANATIONALID',     'replace'),

    # --- 2. Application Details ---
    ('568D8F9B', 'VISASTAYTYPE',       'replace'),   # 2.2 Status of Stay

    # --- Photo ---
    ('24579310', 'VISAPHOTO',          'replace'),   # Photo placeholder cell

    # --- 3. Passport Information ---
    ('77F5831A', 'VISAPASSNUM',        'replace'),
    ('51FF809A', 'VISAPASSISSUEPLACE', 'replace'),
    ('7A092376', 'VISAPASSISSUEDATE',  'replace'),
    ('30F049C0', 'VISAPASSEXPIRY',     'replace'),

    # --- 4. Contact Information ---
    ('3FE3A0B4', 'VISAHOMEADDR',       'replace'),   # 4.1 Home country address
    # 4.2 Current Residential Address — intentionally left blank (no marker)
    ('1FC5EBA2', 'VISACONTACT',        'replace'),   # 4.3 Phone (col break embedded — preserved)
    ('5AF8D824', 'VISAEMAIL',          'inject'),    # 4.4 Email (empty cell in AFRIN doc)
    # 4.5 Emergency Contact — label+value combined paragraphs
    ('7ADD2597', 'a) 성명 Full Name in English VISAEMERGNAME',         'replace'),
    ('5133A0CF', 'd) 관계 Relationship to the applicant VISAEMERGREL', 'replace'),

    # --- 6. Education ---
    # Handled by tag_education_para() — NOT in the generic loop.
    # 4 individual markers are injected into the bracket slots:
    #   VISAEDUMASTER / VISAEDUBACH / VISAEDUHIGH / VISAEDUOTHER

    # --- 7. University / Study ---
    ('4BD49403', 'VISAUNIV',           'replace'),
    ('4F5391CD', 'VISACOURSE',         'replace'),
    ('35253A36', 'VISAUNIVADDR',       'replace'),
    ('05CC5C60', 'VISAUNIVPHONE',      'replace'),

    # --- 8. Visit Details ---
    ('08C35CEE', 'VISASTAY',           'replace'),   # 8.2 Period of Stay — always "365 DAYS"
    ('6BDC96F7', 'VISAENTRYDATE',      'replace'),
    ('304363B1', 'VISAKOREAADDR',      'replace'),
    ('4EEB3C6D', 'VISAKOREATELL',      'replace'),

    # --- 9. Invitation Details ---
    # 9.1 Yes/No — AFRIN has "No[✓]" → swap to "Yes[✓]" (always has an inviter)
    ('0400E826', '[ \u2713]예 Yes []|||[]예 Yes [ \u2713]', 'text_sub'),
    ('14FEA69B', 'VISAINVNAME',        'replace'),
    ('23E15450', 'VISAINVREG',         'replace'),
    # Inviter relationship `533C005E` = "ADMITTED STUDENT" — static, leave as-is
    # Inviter address — label+value combined paragraph
    ('1CC1E9DB', 'd) \uc8fc\uc18c Address VISAINVADDR',  'replace'),
    ('535D8D20', 'VISAINVPHONE',       'replace'),

    # --- 10. Funding ---
    ('6AD5923A', 'VISACOSTS',          'replace'),   # Travel costs (col break embedded — preserved)
    ('7F961417', 'VISASPONSORNAME',    'replace'),
    ('4875E8B1', 'VISASPONSORREL',     'replace'),
    ('1C3223C1', 'VISASPONSORPHONE',   'replace'),

    # --- Section 10.2 floating 2x2 grid overlaps Section 11 header ---
    # The entire section 10.2 grid (sponsor name/relation/support/contact) lives
    # inside ONE floating <wp:anchor>. Its posV is -2138512 EMU (above anchor
    # para 50892618) with cy=623570 (~0.68" tall). Its bottom edge crashes into
    # Section 11's grey header bar. Move it UP (negative delta) so it clears.
    # Any paraId inside the grid targets the same anchor; use 7F961417.
    ('7F961417', '-110000',            'shift_anchor'),

    # --- 6.2 / 6.3 School Details (previous education before Korea) ---
    ('0B673656', 'VISASCHOOLNAME',     'replace'),   # 6.2 School/College name data cell
    ('57F4333B', 'VISASCHOOLLOC',      'replace'),   # 6.3 School location (city/country) data cell

    # --- 11. Assistance with Form (static — always the agency) ---
    ('4C1B4E15', 'VISASEC11NAME',      'inject'),   # Full Name cell (empty in AFRIN doc)
    ('6E808F3A', 'VISASEC11PHONE',     'inject'),   # Telephone No. cell (empty)
    ('550385DA', 'VISASEC11REL',       'inject'),   # Relationship cell (empty)
    # DOB cell `4E8F6773` — intentionally left blank

    # --- 11.1 Assistance Yes/No — always "Yes" (HANGEUL agency helped fill form) ---
    ('3107B95E', '[ \u2713] \uc608 Yes [ ]|||[ ] \uc608 Yes [ \u2713]', 'text_sub'),

    # --- Page 1 overflow fix ---
    # Remove one empty paragraph between the stamp area and the 210×297 footer text
    # so the footer stays on page 1 (email injection adds one line of height that
    # pushes content over). 77D55B58 is an empty para OUTSIDE any table cell — safe
    # to delete (unlike paras inside <w:tc> which must contain at least one <w:p>).
    ('77D55B58', '',                   'delete'),

    # --- 12. Declaration ---
    # Application year cell — clear it (leave blank for manual completion)
    ('64EFA135', '',                   'replace'),
    # Signature
    ('416DF00E', 'VISASIGNATURE',      'replace'),
]

# ---------------------------------------------------------------------------
# Helpers — operate on xml_bytes (UTF-8 encoded bytes), return xml_bytes
# ---------------------------------------------------------------------------

def _find_para(tree, para_id):
    attr = f'{{{W14}}}paraId'
    for p in tree.iter(f'{{{W}}}p'):
        if p.get(attr, '').upper() == para_id.upper():
            return p
    return None


def _make_run(text):
    """Build a <w:r><w:t xml:space="preserve">text</w:t></w:r> element."""
    r = etree.Element(f'{{{W}}}r')
    t = etree.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set(f'{{{XML_SPACE}}}space', 'preserve')
    return r


def inject_marker(xml_bytes, para_id, marker):
    """Clear all runs in an existing empty paragraph and inject a marker run."""
    tree = etree.fromstring(xml_bytes)
    p = _find_para(tree, para_id)
    if p is None:
        print(f'  WARNING: paraId {para_id} not found (inject)')
        return xml_bytes
    # Remove existing runs
    for r in list(p.findall(f'{{{W}}}r')):
        p.remove(r)
    p.append(_make_run(marker))
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def replace_para_text(xml_bytes, para_id, new_text):
    """Replace all <w:t> content in a paragraph with new_text (first run gets it)."""
    tree = etree.fromstring(xml_bytes)
    p = _find_para(tree, para_id)
    if p is None:
        print(f'  WARNING: paraId {para_id} not found (replace)')
        return xml_bytes
    first = True
    for t in p.iter(f'{{{W}}}t'):
        if first:
            t.text = new_text
            t.set(f'{{{XML_SPACE}}}space', 'preserve')
            first = False
        else:
            t.text = ''
    if first:
        # No runs at all — add one
        p.append(_make_run(new_text))
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def delete_para(xml_bytes, para_id):
    """Remove a paragraph entirely from the document."""
    tree = etree.fromstring(xml_bytes)
    p = _find_para(tree, para_id)
    if p is None:
        print(f'  WARNING: paraId {para_id} not found (delete)')
        return xml_bytes
    parent = p.getparent()
    parent.remove(p)
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def insert_para_after(xml_bytes, para_id, marker):
    """Insert a new paragraph containing marker immediately after the target paragraph."""
    tree = etree.fromstring(xml_bytes)
    p = _find_para(tree, para_id)
    if p is None:
        print(f'  WARNING: paraId {para_id} not found (insert_after)')
        return xml_bytes
    parent = p.getparent()
    idx = list(parent).index(p)
    new_p = etree.Element(f'{{{W}}}p')
    new_p.append(_make_run(marker))
    parent.insert(idx + 1, new_p)
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def shift_anchor(xml_bytes, para_id, delta_emu):
    """Find the floating drawing anchor that contains the given paraId and add
    delta_emu to its <wp:positionV><wp:posOffset>. Positive delta moves the
    textbox DOWN the page. 914400 EMU = 1 inch."""
    WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    tree = etree.fromstring(xml_bytes)
    p = _find_para(tree, para_id)
    if p is None:
        print(f'  WARNING: paraId {para_id} not found (shift_anchor)')
        return xml_bytes
    # Walk up to the enclosing <wp:anchor>
    cur = p
    while cur is not None and etree.QName(cur.tag).localname != 'anchor':
        cur = cur.getparent()
    if cur is None:
        print(f'  WARNING: no anchor ancestor for {para_id}')
        return xml_bytes
    pv = cur.find(f'{{{WP}}}positionV')
    if pv is None:
        print(f'  WARNING: no positionV on anchor for {para_id}')
        return xml_bytes
    off = pv.find(f'{{{WP}}}posOffset')
    if off is None or not (off.text or '').strip():
        print(f'  WARNING: no posOffset on positionV for {para_id}')
        return xml_bytes
    old_val = int(off.text.strip())
    off.text = str(old_val + delta_emu)
    print(f'    posOffset {old_val} -> {off.text}')
    # Also update the Fallback <v:shape> style "margin-top" if present, so
    # older Word renderers pick up the same shift.
    for shape in cur.getparent().iter():
        ln = etree.QName(shape.tag).localname
        if ln in ('shape',) and shape.get('style'):
            style = shape.get('style')
            # style contains "margin-top:X.Ypt" — convert delta_emu to pt (1pt=12700 EMU)
            delta_pt = delta_emu / 12700.0
            def repl(m):
                return f'margin-top:{float(m.group(1)) + delta_pt:.2f}pt'
            new_style = re.sub(r'margin-top:([-\d.]+)pt', repl, style)
            if new_style != style:
                shape.set('style', new_style)
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def tag_education_para(xml_bytes):
    """
    Inject 4 individual markers into the education checkbox bracket slots of
    paragraph 2F902749, preserving the original Korean/English labels and tab layout.

    Slot order (left to right in the form):
      VISAEDUMASTER   석사/박사 Master's/Doctoral Degree
      VISAEDUBACH     대졸 Bachelor's Degree
      VISAEDUHIGH     고졸 High School Diploma
      VISAEDUOTHER    기타 Other

    In the source AFRIN.docx the slots are:
      Master/Bachelor/Other  → run with <w:tab/> inside brackets  (empty)
      High School            → run with text '✓' inside brackets  (checked)
    """
    EDU_MARKERS = ['VISAEDUMASTER', 'VISAEDUBACH', 'VISAEDUHIGH', 'VISAEDUOTHER']
    tree = etree.fromstring(xml_bytes)
    p = _find_para(tree, '2F902749')
    if p is None:
        print('  WARNING: paraId 2F902749 not found (tag_education_para)')
        return xml_bytes

    runs = p.findall(f'{{{W}}}r')
    slot = 0
    in_bracket = False

    for r in runs:
        t_nodes = r.findall(f'{{{W}}}t')
        tab_nodes = r.findall(f'{{{W}}}tab')
        text = ''.join(t.text or '' for t in t_nodes)

        if not in_bracket:
            if text == '[':
                in_bracket = True
            continue

        # Inside bracket
        if text.strip() == ']':
            in_bracket = False
            slot = min(slot + 1, len(EDU_MARKERS) - 1)
            continue

        if text == '\u2713':
            # ✓ checkmark run — replace with marker
            for t in t_nodes:
                t.text = EDU_MARKERS[slot]
            if t_nodes:
                t_nodes[0].set(f'{{{XML_SPACE}}}space', 'preserve')
            continue

        if tab_nodes and not text.strip():
            # Tab-only slot run — replace tab element with text marker
            for tab in list(tab_nodes):
                r.remove(tab)
            if t_nodes:
                t_nodes[0].text = EDU_MARKERS[slot]
                t_nodes[0].set(f'{{{XML_SPACE}}}space', 'preserve')
            else:
                t = etree.SubElement(r, f'{{{W}}}t')
                t.text = EDU_MARKERS[slot]
                t.set(f'{{{XML_SPACE}}}space', 'preserve')
            continue
        # else: space padding run or column-separator tab — leave untouched

    # Issue 1 fix: add a visible blank space before '기타' (Other) so it is
    # clearly separated from the 3rd option (High School) when they wrap.
    for r in runs:
        t_nodes = r.findall(f'{{{W}}}t')
        text = ''.join(t.text or '' for t in t_nodes)
        if text == '\uae30\ud0c0':          # 기타
            if t_nodes:
                t_nodes[0].text = ' \uae30\ud0c0'   # prepend a space
                t_nodes[0].set(f'{{{XML_SPACE}}}space', 'preserve')
            break

    fixed = sum(1 for m in EDU_MARKERS if m.encode() in
                etree.tostring(tree, encoding='UTF-8'))
    print(f'  [edu_markers] injected {fixed}/4 education markers')
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def fix_clipped_runs(xml_bytes):
    """
    Fix text that clips at the top in generated files due to JSZip font-metric
    differences interacting with very tight line heights and raised text positions.

    Para 250AB605 (Visiting Family/Relatives/Friends, section 8.1):
      - All runs carry w:position=11 (text raised 5.5pt above baseline)
      - w:line=84 lineRule=auto (0.35× single) — line box too narrow for raised text
      Fix: set w:line=240 (single spacing) so the raised text has room.

    Paras 319806B8 / 2383F94B (국적 Nationality header, sections 8.8 / 8.9):
      - w:before=103/116 + w:line=96/98 lineRule=auto (≈0.4× single) — similarly tight
      Fix: set w:before=0 and w:line=240 so the header fits the row without clipping.
    """
    tree = etree.fromstring(xml_bytes)
    fixed = 0

    # Para 250AB605 — widen line height only
    p = _find_para(tree, '250AB605')
    if p is not None:
        pPr = p.find(f'{{{W}}}pPr')
        if pPr is not None:
            spacing = pPr.find(f'{{{W}}}spacing')
            if spacing is not None:
                spacing.set(f'{{{W}}}line',     '240')
                spacing.set(f'{{{W}}}lineRule', 'auto')
        fixed += 1

    # Paras 319806B8 and 2383F94B — remove before-spacing, widen line height
    for pid in ('319806B8', '2383F94B'):
        p = _find_para(tree, pid)
        if p is None:
            continue
        pPr = p.find(f'{{{W}}}pPr')
        if pPr is not None:
            spacing = pPr.find(f'{{{W}}}spacing')
            if spacing is not None:
                spacing.attrib.pop(f'{{{W}}}before', None)
                spacing.set(f'{{{W}}}line',     '240')
                spacing.set(f'{{{W}}}lineRule', 'auto')
        fixed += 1

    print(f'  [fix_clip] fixed line-height clipping on {fixed} paragraph(s)')
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def text_sub_para(xml_bytes, para_id, old_text, new_text):
    """Replace old_text with new_text across all <w:t> runs in the target paragraph."""
    tree = etree.fromstring(xml_bytes)
    p = _find_para(tree, para_id)
    if p is None:
        print(f'  WARNING: paraId {para_id} not found (text_sub)')
        return xml_bytes
    # Collect all <w:t> nodes and their text in order
    t_nodes = list(p.iter(f'{{{W}}}t'))
    # Build combined text, replace, then redistribute back into the first <w:t>
    combined = ''.join(t.text or '' for t in t_nodes)
    new_combined = combined.replace(old_text, new_text)
    if new_combined == combined:
        print(f'  WARNING: text_sub pattern not found in paraId {para_id}')
        return xml_bytes
    # Put all new text in first node, clear the rest
    if t_nodes:
        t_nodes[0].text = new_combined
        t_nodes[0].set(f'{{{XML_SPACE}}}space', 'preserve')
        for t in t_nodes[1:]:
            t.text = ''
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not os.path.exists(BLANK_DOCX):
        print(f'ERROR: blank DOCX not found at {BLANK_DOCX}')
        return

    os.makedirs(TEMP_DIR, exist_ok=True)
    os.makedirs(os.path.dirname(OUTPUT_DOCX), exist_ok=True)

    with zipfile.ZipFile(BLANK_DOCX, 'r') as z:
        z.extractall(TEMP_DIR)

    doc_path = os.path.join(TEMP_DIR, 'word', 'document.xml')
    with open(doc_path, 'rb') as f:
        xml_bytes = f.read()

    print(f'Loaded document.xml ({len(xml_bytes):,} bytes)')

    # --- Education checkboxes: inject 4 individual markers before main loop ---
    xml_bytes = tag_education_para(xml_bytes)

    print(f'Applying {len(FIELD_MAP)} field mappings...\n')

    for entry in FIELD_MAP:
        para_id, marker, mode = entry
        if mode == 'inject':
            xml_bytes = inject_marker(xml_bytes, para_id, marker)
        elif mode == 'replace':
            xml_bytes = replace_para_text(xml_bytes, para_id, marker)
        elif mode == 'insert_after':
            xml_bytes = insert_para_after(xml_bytes, para_id, marker)
        elif mode == 'text_sub':
            old_text, new_text = marker.split('|||', 1)
            xml_bytes = text_sub_para(xml_bytes, para_id, old_text, new_text)
        elif mode == 'delete':
            xml_bytes = delete_para(xml_bytes, para_id)
        elif mode == 'shift_anchor':
            # marker is the integer EMU delta (positive = push down)
            xml_bytes = shift_anchor(xml_bytes, para_id, int(marker))
        else:
            print(f'  WARNING: unknown mode "{mode}" for {para_id}')
            continue
        print(f'  [{mode:<12}]  {para_id}  ->  {marker}')

    # --- Fix clipped text (Visiting Family/Relatives/Friends, Nationality) ---
    xml_bytes = fix_clipped_runs(xml_bytes)

    # --- Procedure arrows: replace 'è' with ▶ (U+25B6, solid black right arrow) ---
    # Strip w:w (190% expansion) and w:spacing from these runs so the glyph
    # stays at normal width — that keeps all 5 boxes on one row.
    tree = etree.fromstring(xml_bytes)
    arrow_fixed = 0
    for r in list(tree.iter(f'{{{W}}}r')):
        t_nodes = list(r.findall(f'{{{W}}}t'))
        if not t_nodes:
            continue
        combined = ''.join(t.text or '' for t in t_nodes)
        if not combined or set(combined) != {'\u00e8'}:
            continue
        # Replace each 'è' with one ▶ arrow
        t_nodes[0].text = '\u25b6' * len(combined)
        t_nodes[0].set(f'{{{XML_SPACE}}}space', 'preserve')
        for extra in t_nodes[1:]:
            extra.text = ''
        # Remove width-expansion and letter-spacing that were tuned for 'è'
        rpr = r.find(f'{{{W}}}rPr')
        if rpr is not None:
            for tag in ('w', 'spacing', 'kern'):
                elt = rpr.find(f'{{{W}}}{tag}')
                if elt is not None:
                    rpr.remove(elt)
        arrow_fixed += 1
    xml_bytes = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    print(f'\n  [arrow_fill] replaced {arrow_fixed} \u00e8 run(s) with \u25b6')

    with open(doc_path, 'wb') as f:
        f.write(xml_bytes)

    # Repack DOCX — write to a temp file first, then replace (handles file-lock on Windows)
    tmp_out = OUTPUT_DOCX + '.tmp'
    with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(TEMP_DIR):
            for file in files:
                abs_path = os.path.join(root, file)
                arc_name = os.path.relpath(abs_path, TEMP_DIR)
                zout.write(abs_path, arc_name)
    # os.replace is atomic on Windows and overwrites the destination
    if os.path.exists(OUTPUT_DOCX):
        try:
            os.remove(OUTPUT_DOCX)
        except OSError:
            pass  # If still locked, os.replace below may still succeed
    os.replace(tmp_out, OUTPUT_DOCX)

    shutil.rmtree(TEMP_DIR, ignore_errors=True)
    print(f'\nDone. Tagged template -> {OUTPUT_DOCX}')
    print(f'Total markers injected: {len(FIELD_MAP)}')


if __name__ == '__main__':
    main()
