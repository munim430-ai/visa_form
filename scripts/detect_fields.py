"""
detect_fields.py
----------------
Diffs a blank and a filled visa DOCX to discover which XML paragraphs
(identified by w14:paraId) changed.  Run this ONCE to determine the
paraId values you need in tag_template.py.

Usage:
    python scripts/detect_fields.py blank_visa.docx filled_visa.docx

Output: a table printed to stdout showing each changed paragraph's
        paraId, nearby label text, and the filled value.

Prerequisites:
    pip install lxml
"""

import sys
import zipfile
from lxml import etree
from collections import OrderedDict

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'

def extract_doc_xml(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            return f.read()

def get_paragraphs(xml_bytes):
    """Return OrderedDict: paraId -> full text content of the paragraph."""
    tree = etree.fromstring(xml_bytes)
    result = OrderedDict()
    attr = f'{{{W14}}}paraId'
    for para in tree.iter(f'{{{W}}}p'):
        pid = para.get(attr)
        if not pid:
            continue
        text = ''.join(t.text or '' for t in para.iter(f'{{{W}}}t'))
        result[pid] = text
    return result

def get_paragraphs_with_context(xml_bytes):
    """Return dict: paraId -> {text, prev_text, next_text}"""
    tree = etree.fromstring(xml_bytes)
    attr = f'{{{W14}}}paraId'
    paras = []
    for para in tree.iter(f'{{{W}}}p'):
        pid = para.get(attr)
        if not pid:
            continue
        text = ''.join(t.text or '' for t in para.iter(f'{{{W}}}t'))
        paras.append((pid, text))

    result = {}
    for i, (pid, text) in enumerate(paras):
        prev_text = paras[i-1][1] if i > 0 else ''
        next_text = paras[i+1][1] if i < len(paras)-1 else ''
        result[pid] = {
            'text': text,
            'prev': prev_text,
            'next': next_text,
        }
    return result

def main():
    if len(sys.argv) < 3:
        print("Usage: python detect_fields.py <blank.docx> <filled.docx>")
        sys.exit(1)

    blank_path  = sys.argv[1]
    filled_path = sys.argv[2]

    print(f"Blank : {blank_path}")
    print(f"Filled: {filled_path}")
    print()

    blank_xml  = extract_doc_xml(blank_path)
    filled_xml = extract_doc_xml(filled_path)

    blank_paras  = get_paragraphs(blank_xml)
    filled_paras = get_paragraphs_with_context(filled_xml)

    changed = []
    for pid, blank_text in blank_paras.items():
        if pid not in filled_paras:
            continue
        filled_info = filled_paras[pid]
        filled_text = filled_info['text']
        # Detect: was empty/whitespace in blank, now has content in filled
        if blank_text.strip() == '' and filled_text.strip() != '':
            changed.append({
                'paraId':      pid,
                'blank':       blank_text,
                'filled':      filled_text,
                'label_before': filled_info['prev'][:60],
                'label_after':  filled_info['next'][:60],
            })
        # Detect: content changed (non-checkbox)
        elif blank_text.strip() != filled_text.strip() and blank_text.strip() != '':
            changed.append({
                'paraId':      pid,
                'blank':       blank_text[:40],
                'filled':      filled_text[:40],
                'label_before': filled_info['prev'][:60],
                'label_after':  filled_info['next'][:60],
            })

    if not changed:
        print("No differences found. Check that the correct files were passed.")
        return

    # Print results
    col_w = [10, 20, 20, 30, 30]
    header = ['paraId', 'blank_text', 'filled_value', 'label_before', 'label_after']
    sep = '  '.join('-' * w for w in col_w)

    def row(vals):
        return '  '.join(str(v)[:w].ljust(w) for v, w in zip(vals, col_w))

    print(row(header))
    print(sep)
    for c in changed:
        print(row([
            c['paraId'],
            c['blank'],
            c['filled'],
            c['label_before'],
            c['label_after'],
        ]))

    print()
    print(f"Total changed paragraphs: {len(changed)}")
    print()
    print("--- Copy-paste ready for tag_template.py FIELD_MAP ---")
    for c in changed:
        label = (c['label_before'] or c['label_after']).strip()[:30]
        print(f"    # {label}")
        print(f"    ('{c['paraId']}', 'VISA???', 'inject'),")

if __name__ == '__main__':
    main()
