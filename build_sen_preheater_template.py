"""
build_sen_preheater_template.py  — v3
Builds SEN_Preheater_Offer_Template.docx from Burner_Offer_Template.docx.

Strategy:
  1. Keep cover, client tables, T&C table exactly as-is.
  2. Update TOC and Annexures tables.
  3. Replace Tech Spec table rows with SEN rows.
  4. Replace price-schedule row description with SEN Preheater.
  5. Replace the scope section (paragraphs 54-78) with clean SEN content.
  6. Rename Annexure I heading.

Run once:  python build_sen_preheater_template.py
"""
from __future__ import annotations
import os, shutil
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

BASE   = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(BASE, "Burner_Offer_Template.docx")
TARGET = os.path.join(BASE, "SEN_Preheater_Offer_Template.docx")

shutil.copy2(SOURCE, TARGET)
doc = Document(TARGET)

# ── Helper ────────────────────────────────────────────────────────────────────
def _set_cell(cell, text, bold=False, size=9, bg=None):
    for p in cell.paragraphs:
        for r in p.runs: r.text = ""
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    r = p.add_run(text); r.bold = bold; r.font.size = Pt(size)
    if bg:
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), bg)
        tcPr.append(shd)

def _remove_para(p):
    el = p._element; el.getparent().remove(el)

def _new_para(text, bold=False, underline=False, size=9, italic=False):
    """Create a bare w:p element with one run."""
    p_el = OxmlElement("w:p")
    pPr  = OxmlElement("w:pPr")
    rPr  = OxmlElement("w:rPr")
    if bold:      rPr.append(OxmlElement("w:b"))
    if underline:
        u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    if italic:    rPr.append(OxmlElement("w:i"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(size * 2)); rPr.append(sz)
    szCs = OxmlElement("w:szCs"); szCs.set(qn("w:val"), str(size * 2)); rPr.append(szCs)
    r = OxmlElement("w:r"); r.append(rPr)
    t = OxmlElement("w:t"); t.text = text
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t); p_el.append(r)
    return p_el

# ── TABLE 1: TOC ───────────────────────────────────────────────────────────────
toc = doc.tables[1]
rows = toc.rows
if len(rows) >= 2:
    _set_cell(rows[0].cells[0], "1.", bold=True)
    _set_cell(rows[0].cells[1], "COMPANY PROFILE")
    _set_cell(rows[1].cells[0], "2.", bold=True)
    _set_cell(rows[1].cells[1], "ABOUT THE CLIENT")
    r3 = toc.add_row()
    _set_cell(r3.cells[0], "3.", bold=True); _set_cell(r3.cells[1], "TECHNICAL SPECIFICATIONS")
    r4 = toc.add_row()
    _set_cell(r4.cells[0], "4.", bold=True); _set_cell(r4.cells[1], "GENERAL INFORMATION")

# ── TABLE 2: Annexures ─────────────────────────────────────────────────────────
ann = doc.tables[2]
ann_data = [
    ("Annexure - I",   "EXCLUSIONS",        "Attached"),
    ("Annexure \u2013 II",  "PRICE SCHEDULE",     "Attached"),
    ("Annexure - III", "TERMS & CONDITIONS", "Attached"),
    ("Annexure - IV",  "REFERENCE LIST",     "Attached"),
]
for i, (a, b, c_) in enumerate(ann_data):
    if i + 1 < len(ann.rows):
        _set_cell(ann.rows[i+1].cells[0], a)
        _set_cell(ann.rows[i+1].cells[1], b)
        _set_cell(ann.rows[i+1].cells[2], c_)
    else:
        nr = ann.add_row()
        _set_cell(nr.cells[0], a); _set_cell(nr.cells[1], b); _set_cell(nr.cells[2], c_)

# ── TABLE 5: Tech Spec ─────────────────────────────────────────────────────────
spec = doc.tables[5]
SEN_ROWS = [
    ("Heating Medium",                  "{{ fuel_name }} Fired Burners"),
    ("Quantity",                        "{{ sen_quantity }}"),
    ("Minimum Temperature",             "Ambient to 1200\u00B0C"),
    ("Heating Time",                    "{{ heating_time }}"),
    ("{{ gas_line }} Burner Capacity",  "{{ burner_capacity_sen }}"),
    ("No. of Burners",                  "{{ num_burners_sen }}"),
    ("Combustion Air Blower",           "{{ blower_spec_sen }}"),
    ("Ignition of Burner",              "Manual"),
]
# Header
if spec.rows:
    hdr = spec.rows[0]
    _set_cell(hdr.cells[0], "SEN PREHEATER", bold=True, size=10, bg="1A3A5C")
    _set_cell(hdr.cells[1], "TECHNICAL DATA", bold=True, size=10, bg="1A3A5C")

# Data rows
existing = list(spec.rows[1:])
for i, (lbl, val) in enumerate(SEN_ROWS):
    if i < len(existing):
        _set_cell(existing[i].cells[0], lbl, bold=True, size=9, bg="EEF2F7")
        _set_cell(existing[i].cells[1], val, size=9)
    else:
        nr = spec.add_row()
        _set_cell(nr.cells[0], lbl, bold=True, size=9, bg="EEF2F7")
        _set_cell(nr.cells[1], val, size=9)
# Remove excess rows
for row in existing[len(SEN_ROWS):]:
    row._element.getparent().remove(row._element)

# ── TABLE 6: Price Schedule — update item description ─────────────────────────
price_tbl = doc.tables[6]
for row in price_tbl.rows:
    if len(row.cells) >= 2 and "price_heading" in row.cells[1].text:
        _set_cell(row.cells[1], "SEN Preheater")
        break

# ── Paragraphs: replace scope section ─────────────────────────────────────────
# Find the range: from para with "scope_intro" to para with "ANNEXURE II"
paras = list(doc.paragraphs)
scope_start = None
ann2_idx    = None
for i, p in enumerate(paras):
    if "scope_intro" in p.text and scope_start is None:
        scope_start = i
    if "ANNEXURE II" in p.text.upper() and "PRICE" in p.text.upper():
        ann2_idx = i
        break

if scope_start is not None and ann2_idx is not None:
    body = doc.element.body
    # Collect elements to remove
    to_remove = [p._element for p in paras[scope_start:ann2_idx]]
    
    # Build replacement paragraphs (inserted before ANNEXURE II element)
    ann2_el = paras[ann2_idx]._element

    new_paras = [
        # Scope intro
        _new_para("{{ pipeline_scope_text }}", size=9),
        _new_para(""),  # spacer
        # OBJECTIVE
        _new_para("OBJECTIVE", bold=True, underline=True, size=10),
        _new_para("{{ sen_objective }}", size=9),
        _new_para(""),
        # SCOPE OF SUPPLY
        _new_para("SCOPE OF SUPPLY", bold=True, underline=True, size=10),
        _new_para("Our scope of supply will cover design, engineering, manufacture supply, supervision of commissioning & erection of the following main components:", size=9),
        _new_para(""),
        # ENCON BURNERS
        _new_para("{{ sen_burners_heading }}", bold=True, underline=True, size=9),
        _new_para("{{ sen_burners_body }}", size=9),
        _new_para(""),
        # GAS LINE
        _new_para("{{ gas_line }} LINE FOR BURNERS", bold=True, underline=True, size=9),
        _new_para("The {{ gas_line }} line for main burners shall be routed from the main gas train and will consist of the following main components:", size=9),
        # gas line items loop
        _new_para("{%p for x in fuel1_line_items %}"),
        _new_para("\u00b7  {{ x.item }}", size=9),
        _new_para("{%p endfor %}"),
        _new_para(""),
        # AIR LINE
        _new_para("AIR LINE FOR MAIN BURNERS:", bold=True, underline=True, size=9),
        # air line items loop
        _new_para("{%p for x in air_pipeline_items %}"),
        _new_para("\u00b7  {{ x.item }}", size=9),
        _new_para("{%p endfor %}"),
        _new_para(""),
        # COMBUSTION AIR BLOWER
        _new_para("COMBUSTION AIR BLOWER", bold=True, underline=True, size=9),
        _new_para("{{ sen_combustion_blower_text }}", size=9),
        _new_para(""),
        # TROLLEY
        _new_para("TROLLEY", bold=True, underline=True, size=9),
        _new_para("{{ sen_trolley_text }}", size=9),
        _new_para(""),
        # ROLLER RACK
        _new_para("ROLLER SUPPORTED RACK WITH HANDLE", bold=True, underline=True, size=9),
        _new_para("{{ sen_roller_rack_text }}", size=9),
        _new_para(""),
        # NOTE
        _new_para("Note: - {{ sen_note }}", size=9, italic=True),
        _new_para(""),
    ]

    # Insert new paragraphs before ann2_el (in forward order)
    for np_el in new_paras:
        ann2_el.addprevious(np_el)

    # Remove old scope paragraphs
    for el in to_remove:
        try: el.getparent().remove(el)
        except: pass

# ── Rename ANNEXURE I heading ──────────────────────────────────────────────────
for p in doc.paragraphs:
    if "ANNEXURE I" in p.text.upper() and ("SCOPE" in p.text.upper() or "BURNER" in p.text.upper()):
        for run in p.runs: run.text = ""
        p.add_run("ANNEXURE I \u2014 EXCLUSIONS")
        break

# ── Save ───────────────────────────────────────────────────────────────────────
doc.save(TARGET)
print(f"Saved: {TARGET}")

# ── Verify ─────────────────────────────────────────────────────────────────────
from docxtpl import DocxTemplate
tpl = DocxTemplate(TARGET)
try:
    v = sorted(tpl.get_undeclared_template_variables())
    print(f"Variables ({len(v)}): {v}")
except Exception as e:
    print(f"Jinja error: {e}")
