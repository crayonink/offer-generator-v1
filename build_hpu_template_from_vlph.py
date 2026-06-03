"""Build HPU_Offer_Template.docx by cloning the VLPH offer template
(Offer_Template.docx) and surgically replacing the ladle-preheater
content with stand-alone Hydraulic Pumping Unit content.

Modelled on build_recup_template_from_vlph.py — same shell, just a
different equipment.

What is kept from the VLPH template:
  * Cover page, ENCON banner, cover letter
  * Table of Contents + List of Annexures
  * Section 1 Company Profile
  * Section 2 About the Client + Client Details + Marketing Person
  * Section heading "3. TECHNICAL SPECIFICATIONS"
  * ANNEXURE I  (Scope of Supply table shell — content swapped)
  * ANNEXURE II (Exclusions — universal)
  * ANNEXURE III (Price Schedule — single-line)
  * Supervision Charges sub-table
  * ANNEXURE IV (T&C 12-row table)
  * ANNEXURE V  (Reference List — left as ladle list for now;
                 user edits in Word until HPU clients are curated)
  * ANNEXURE VI (Make List — universal)

What is replaced:
  * "VERTICAL LADLE PREHEATERS AND DRYERS" -> "HYDRAULIC PUMPING UNIT"
  * 24-row Technical Specs table -> 8-row HPU Tech Specs
  * VLPH scope-of-supply body (200+ paragraphs on hood/burner/blower
    /gas-train/PLC) -> short HPU description
  * Annexure I 12-row Scope table -> HPU standard accessories list
  * Annexure I intro paragraph -> HPU-specific wording

Template placeholders introduced (filled by /api/generate-hpu-quote):
  hpu_variant, hpu_kw, hpu_lph, hpu_qty
"""
from __future__ import annotations

import os
import shutil
from copy import deepcopy

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.enum.table import WD_ROW_HEIGHT_RULE

SOURCE = "Offer_Template.docx"
TARGET = "HPU_Offer_Template.docx"


# ── HPU Tech Specs table (replaces VLPH's 24-row preheater spec) ──────────
_HPU_SPECS = [
    ('Equipment',           '{{ equipment_name }}'),
    ('Variant',             '{{ hpu_variant }}'),
    ('Motor Capacity',      '{{ hpu_kw }} kW'),
    ('Oil Flow Rate',       '{{ hpu_lph }} LPH'),
    ('Quantity',            '{{ hpu_qty }} No.'),
    ('Construction',        'MS skid-mounted with pre-wired control panel'),
    ('Electrical Supply',   '415 V / 3 Phase / 50 Hz, 4-Wire'),
    ('Operating Pressure',  'Up to 7 kg/cm² (adjustable)'),
]


# ── Annexure I — Scope of Supply (HPU accessories) ─────────────────────────
# The S.No. column is auto-numbered in _replace_scope_table_with_hpu().
# Each entry is (description, quantity). Pump/motor qty switches between
# Simplex (1 No.) and Duplex (2 Nos.) via the {% if %} Jinja gate; the
# rest of the rows are constant.
_HPU_SCOPE_ITEMS = [
    ('Oil tank — MS construction with cleaning door, breather, level indicator, drain plug',
     '1 No.'),
    ('Gear pump suitable for fuel-oil service',
     '{% if "duplex" in hpu_variant.lower() %}2 Nos.{% else %}1 No.{% endif %}'),
    ('TEFC motor — {{ hpu_kw }} kW × 1440 RPM × 415V / 3-Ph / 50 Hz',
     '{% if "duplex" in hpu_variant.lower() %}2 Nos.{% else %}1 No.{% endif %}'),
    ('Coupling with guard (pump-motor)',
     '{% if "duplex" in hpu_variant.lower() %}2 Nos.{% else %}1 No.{% endif %}'),
    ('Suction strainer 150 micron',
     '{% if "duplex" in hpu_variant.lower() %}2 Nos.{% else %}1 No.{% endif %}'),
    ('Pressure relief valve',
     '{% if "duplex" in hpu_variant.lower() %}2 Nos.{% else %}1 No.{% endif %}'),
    ('Pressure gauge with isolation valve (0–10 kg/cm²)',
     '1 No.'),
    ('Duplex line strainer 25 micron (changeover type) / Inline strainer 25 micron',
     '1 No.'),
    ('Non-return valve',
     '{% if "duplex" in hpu_variant.lower() %}2 Nos.{% else %}1 No.{% endif %}'),
    ('Pressure switch (low-pressure cut-off)',                       '1 No.'),
    ('Float-type oil level switch (low-level cut-off)',              '1 No.'),
    ('Inter-connecting piping inside the unit',                      'Lot'),
    ('Common skid base frame with mounting bolts',                   '1 No.'),
    ('Control panel — start / stop / interlock / indication',        '1 No.'),
    ('Set of foundation bolts & matching flanges',                   '1 Lot'),
]


def _para(text: str = '', *, bold: bool = False, underline: bool = False) -> OxmlElement:
    p = OxmlElement('w:p')
    if text:
        r = OxmlElement('w:r')
        if bold or underline:
            rPr = OxmlElement('w:rPr')
            if bold:
                rPr.append(OxmlElement('w:b'))
            if underline:
                u = OxmlElement('w:u')
                u.set(qn('w:val'), 'single')
                rPr.append(u)
            r.append(rPr)
        t = OxmlElement('w:t')
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        r.append(t)
        p.append(r)
    return p


def _replace_spec_table_with_hpu(table) -> None:
    """Wipe the VLPH 24-row tech-data table and re-fill with HPU specs.
    Keep the header row from VLPH; relabel it 'HPU Specifications'."""
    tbl_xml = table._element
    for tr in list(tbl_xml.findall(qn('w:tr')))[1:]:
        tbl_xml.remove(tr)
    head = table.rows[0]
    head.cells[0].text = 'HPU Specifications'
    head.cells[1].text = ''
    head.cells[0].merge(head.cells[1])
    for p in head.cells[0].paragraphs:
        for r in p.runs:
            r.bold = True
    for label, value in _HPU_SPECS:
        row = table.add_row()
        row.cells[0].text = label
        row.cells[1].text = value


def _replace_scope_table_with_hpu(table) -> None:
    """Replace the Annexure I 12-row Scope of Supply with HPU accessories.
    VLPH layout: header row, banner row (merged), 10 scope rows.
    HPU layout : header row, banner row (merged), 15 scope rows."""
    tbl_xml = table._element
    for tr in list(tbl_xml.findall(qn('w:tr')))[1:]:
        tbl_xml.remove(tr)

    # Banner row (merged) — short label
    banner = table.add_row()
    banner.cells[0].merge(banner.cells[1])
    banner.cells[0].text = 'Hydraulic Pumping Unit'
    for p in banner.cells[0].paragraphs:
        for r in p.runs:
            r.bold = True

    # Numbered scope rows. The table is 2-col 'S. No. | Description'; the
    # original VLPH layout has no Qty column. To keep table-width parity,
    # fold quantity into the description: "Item — Qty".
    for idx, (desc, qty) in enumerate(_HPU_SCOPE_ITEMS, start=1):
        row = table.add_row()
        row.cells[0].text = str(idx)
        row.cells[1].text = f'{desc}  ({qty})'


def _delete_vlph_scope_body(doc: Document) -> None:
    """Walk the body looking for the inline 'SCOPE OF SUPPLY' paragraph
    (NOT the Annexure heading) and delete every element from there up
    to (but not including) the 'ANNEXURE I' heading. Replace with a
    short HPU description block."""
    body = doc.element.body
    elements = list(body)

    start_idx, end_idx = None, None
    for i, el in enumerate(elements):
        if el.tag.split('}')[-1] != 'p':
            continue
        txt = ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()
        if start_idx is None and txt == 'SCOPE OF SUPPLY':
            start_idx = i
        elif start_idx is not None and txt.startswith('ANNEXURE I'):
            end_idx = i
            break

    if start_idx is None or end_idx is None:
        print('Could not locate VLPH scope body markers — skipping delete.')
        return

    anchor = elements[end_idx]
    for el in elements[start_idx:end_idx]:
        body.remove(el)

    hpu_body = [
        _para('HYDRAULIC PUMPING UNIT', bold=True),
        _para(
            'The offered Hydraulic Pumping Unit is an oil pumping skid '
            'designed to deliver fuel oil at the required pressure and '
            'flow rate to combustion equipment such as preheater burners, '
            'reheating furnaces and bath heaters. The unit is offered in '
            'three configurations — Simplex (single pump-motor), '
            'Duplex 1 and Duplex 2 (twin pump-motor arrangements with '
            'a standby line) — and is built on a common MS skid base '
            'with a pre-wired control panel for start / stop, '
            'interlocks and indication.'
        ),
        _para(
            'Refer to the HPU Specifications table above for the '
            'offered configuration and motor capacity. The complete '
            'scope of standard accessories shipped with each unit is '
            'detailed in Annexure I — Scope of Supply.'
        ),
    ]
    for el in hpu_body:
        anchor.addprevious(el)

    # Flip the section title. In the source template this heading is a
    # Jinja conditional:
    #   {% if is_tundish %}TUNDISH{% elif is_horizontal %}HORIZONTAL LADLE
    #   {% else %}VERTICAL LADLE{% endif %} PREHEATERS AND DRYERS
    # We match on the trailing 'PREHEATERS AND DRYERS' which is stable,
    # then wipe all runs and stamp the HPU heading.
    for el in body.iter(qn('w:p')):
        t_els = list(el.iter(qn('w:t')))
        joined = ''.join((t.text or '') for t in t_els).upper()
        if 'PREHEATERS AND DRYERS' in joined:
            for t in t_els:
                t.text = ''
            if t_els:
                t_els[0].text = 'HYDRAULIC PUMPING UNIT'


def _inject_scope_intro_paragraph(doc: Document) -> None:
    """Rewrite Annexure I's intro paragraph to HPU-specific wording."""
    body = doc.element.body
    elements = list(body)

    heading_idx = None
    for i, el in enumerate(elements):
        if el.tag.split('}')[-1] != 'p':
            continue
        txt = ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()
        if 'ANNEXURE I' in txt.upper() and 'SCOPE OF SUPPLY' in txt.upper():
            heading_idx = i
            break
    if heading_idx is None:
        return

    intro_p = None
    for j in range(heading_idx + 1, len(elements)):
        if elements[j].tag.split('}')[-1] == 'p':
            intro_p = elements[j]
            break
    if intro_p is None:
        return

    NEW_INTRO = (
        'The Scope of supply covers Design, manufacturing and supply '
        'of the Hydraulic Pumping Unit as per the specifications '
        'listed below. The standard accessories shipped with each '
        'unit are itemised in the table that follows.'
    )
    t_els = list(intro_p.iter(qn('w:t')))
    if t_els and t_els[0].text and t_els[0].text.startswith(NEW_INTRO[:40]):
        return  # idempotent
    for t in t_els:
        t.text = ''
    if t_els:
        t_els[0].text = NEW_INTRO
    else:
        r = OxmlElement('w:r')
        t = OxmlElement('w:t'); t.text = NEW_INTRO; r.append(t)
        intro_p.append(r)


def _clear_reference_list(doc: Document) -> None:
    """Wipe the 50-row VLPH Reference List (Annexure V). The HPU offer
    has no curated client list yet, so we strip rows and leave a single
    placeholder row the user can fill in Word. Header (S. No. | Client
    | Application) is preserved.

    Also rewrites the intro paragraph above the table from the
    'ladle preheating systems' line to a neutral HPU equivalent."""
    table = None
    for t in doc.tables:
        head = [c.text.strip() for c in t.rows[0].cells] if t.rows else []
        if (len(head) >= 3 and head[0].upper() == 'S. NO.'
                and head[1].upper() == 'CLIENT' and 'APPLICATION' in head[2].upper()):
            table = t
            break
    if table is None:
        print('Reference List table not found — skipping clear.')
        return

    tbl_xml = table._element
    for tr in list(tbl_xml.findall(qn('w:tr')))[1:]:
        tbl_xml.remove(tr)

    # Single placeholder row.
    row = table.add_row()
    row.cells[0].text = '1'
    row.cells[1].text = '— (to be added)'
    row.cells[2].text = 'Hydraulic Pumping Unit'

    # Rewrite intro paragraph.
    for p in doc.paragraphs:
        if 'ladle preheating systems' in p.text.lower():
            for run in p.runs:
                run.text = ''
            if p.runs:
                p.runs[0].text = (
                    'We have supplied Hydraulic Pumping Units to various '
                    'clients across India and overseas; a representative '
                    'list will be furnished on request.'
                )
            else:
                p.add_run(
                    'We have supplied Hydraulic Pumping Units to various '
                    'clients across India and overseas; a representative '
                    'list will be furnished on request.'
                )
            break


def _pad_table_rows(doc: Document, min_height_cm: float = 0.75) -> None:
    """Give the Price Schedule and Supervision sub-table some breathing
    room so the printed HPU offer doesn't look cramped."""
    targets = []
    for t in doc.tables:
        if not t.rows or len(t.rows[0].cells) < 2:
            continue
        head_r = t.rows[0].cells[1].text.strip().upper() if len(t.rows[0].cells) > 1 else ''
        head_l = t.rows[0].cells[0].text.strip()
        if 'ITEM DESCRIPTION' in head_r:
            targets.append(t)
        elif head_l.startswith('Supervision Charges'):
            targets.append(t)

    height = Cm(min_height_cm)
    for tbl in targets:
        for row in tbl.rows:
            first_cell_text = row.cells[0].text.strip() if row.cells else ''
            if first_cell_text.startswith('{%tr') or first_cell_text.startswith('{%p'):
                continue
            row.height = height
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def main() -> None:
    if not os.path.exists(SOURCE):
        raise SystemExit(f"missing: {SOURCE}")

    if os.path.exists(TARGET) and not os.path.exists(TARGET + '.bak'):
        shutil.copy(TARGET, TARGET + '.bak')

    shutil.copy(SOURCE, TARGET)
    doc = Document(TARGET)
    print(f'Cloned {SOURCE} -> {TARGET}: '
          f'{len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables')

    # 1. Delete VLPH scope-of-supply body; insert HPU description block.
    _delete_vlph_scope_body(doc)

    # 2. Replace T5 (Tech Specs 24-row) with HPU specs.
    designing_table = doc.tables[5]
    _replace_spec_table_with_hpu(designing_table)

    # 3. Replace Annexure I scope table with HPU accessories. Locate by
    #    row count (10..15 rows, 'S. No.' header, not the Reference List
    #    50-row table and not the Make List).
    scope_table = None
    for t in doc.tables:
        if not t.rows or len(t.rows[0].cells) < 2:
            continue
        head_l = t.rows[0].cells[0].text.strip().upper()
        head_r = t.rows[0].cells[1].text.strip().upper()
        if head_l == 'S. NO.' and 'ITEM DESCRIPTION' not in head_r:
            if 10 <= len(t.rows) <= 15:
                scope_table = t
                break
    if scope_table is not None:
        _replace_scope_table_with_hpu(scope_table)
    else:
        print('Annexure I Scope of Supply table not found — leaving as-is.')

    # 4. Rewrite the intro paragraph above the Scope of Supply table.
    _inject_scope_intro_paragraph(doc)

    # 5. Wipe the 50-row Reference List; leave a placeholder row.
    _clear_reference_list(doc)

    # 6. Give the Price Schedule + Supervision rows some breathing room.
    _pad_table_rows(doc, min_height_cm=0.75)

    doc.save(TARGET)
    print(f'Saved -> {TARGET}')


if __name__ == '__main__':
    main()
