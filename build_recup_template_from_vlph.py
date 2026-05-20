"""Build Recup_Offer_Template.docx by cloning the VLPH offer template
(Offer_Template.docx) and surgically replacing the ladle-preheater
content with recuperator content.

What this script keeps from the VLPH template (the "shell"):
  * Cover page, ENCON banner, cover letter
  * Table of Contents + List of Annexures
  * Section 1 Company Profile (universal)
  * Section 2 About the Client + Client Details + Marketing Person tables
  * Section heading "3. TECHNICAL SPECIFICATIONS"
  * ANNEXURE I (Scope of Supply table shell — content swapped)
  * ANNEXURE II (Exclusions paragraphs — universal)
  * ANNEXURE III (Price Schedule table — rebuilt for single/full toggle)
  * Supervision Charges sub-table
  * ANNEXURE IV (T&C 12-row table — universal)
  * ANNEXURE V (Reference List — universal)
  * ANNEXURE VI (Make List — universal)

What this script replaces:
  * "VERTICAL LADLE PREHEATERS AND DRYERS" -> "RECUPERATOR (Waste Heat
    Recovery System)"
  * 24-row Technical Specs table  -> 15-row Recup Designing Params
  * Material of Construction sub-table inserted ABOVE the Designing
    Params table
  * 3D image placeholders inserted AFTER the Designing Params table
  * VLPH scope-of-supply body (~200 paragraphs of fuel/oil/PLC/hood
    conditionals) -> short recup operation description
  * Annexure I 12-row Scope table content -> recup component scope rows
  * Annexure III Price Schedule -> single/full toggle (matches
    build_recup_template_v2.py output)

After this script the template still expects all VLPH-style
placeholders to be supplied by the backend:
  equipment_name, poc_designation, technical_person, technical_phone,
  technical_email, supervision_mech, supervision_plc, tnc_* etc.

Plus the recup-specific keys introduced here:
  hot_tube_plate_material, hot_tube_material, cold_tube_plate_material,
  cold_tube_material, flue_*, air_*, surface_area_m2, pipe_*,
  image_recup_side/front/top, bom_rows, bought_out_total, encon_total,
  grand_total, grand_total_in_words, price_schedule_style,
  supervision_include/rate/note.
"""
from __future__ import annotations

import os
import shutil
from copy import deepcopy

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SOURCE = "Offer_Template.docx"
TARGET = "Recup_Offer_Template.docx"


# ── Material of Construction layout (cold/hot/others) ──────────────────────
# Mirrors build_recup_template_v2.py's _MOC_ROWS.
_MOC_ROWS = [
    ('header',  'Material of Construction',  None),
    ('section', 'COLD BANK',                 None),
    ('item',    'Tube Plate',                '{{ cold_tube_plate_material }}'),
    ('item',    'Tube',                      '{{ cold_tube_material }}'),
    ('item',    'Duct Inlet',                'Mild Steel'),
    ('item',    'Inlet Collar',              'Mild Steel'),
    ('section', 'HOT BANK',                  None),
    ('item',    'Tube Plate',                '{{ hot_tube_plate_material }}'),
    ('item',    'Tube',                      '{{ hot_tube_material }}'),
    ('item',    'Duct Outlet',               'Mild Steel'),
    ('item',    'Outlet Collar',             'Mild Steel'),
    ('item',    'Supporting Frame',          'Mild Steel'),
    ('item',    'Bottom Duct',               'Mild Steel'),
    ('item',    'Flange',                    'Mild Steel'),
    ('section', 'OTHERS',                    None),
    ('item',    'Flanges',                   'Mild Steel'),
    ('item',    'Air Inlet Duct',            'Mild Steel'),
    ('item',    'Air Outlet Duct',           'Mild Steel'),
    ('item',    'Bottom Air Receiving Box',  'Mild Steel'),
    ('item',    'Matching Flange',           'Mild Steel'),
    ('item',    'Nut, Bolt & Washer',        'Mild Steel'),
    ('item',    'Gasket',                    'Heat Resistant'),
    ('item',    'Supporting Structure',      'Mild Steel'),
]

# ── Recup Designing Parameters layout (replaces VLPH's 24-row spec) ────────
_DESIGNING_PARAMS = [
    ('Equipment',                 '{{ equipment_name }}'),
    ('Application',               '{{ application }}'),
    ('No. of Units',              '{{ recup_qty }}'),
    ('Flue Gas Flow',             '{{ flue_flow_nm3hr }} Nm³/hr'),
    ('Flue Gas Temperature (In)', '{{ flue_temp_in_C }} °C'),
    ('Flue Gas Temperature (Out)','{{ flue_temp_out_C }} °C'),
    ('Air Volume',                '{{ air_volume_nm3hr }} Nm³/hr'),
    ('Air Temperature (In)',      '{{ air_temp_in_C }} °C'),
    ('Air Temperature (Out)',     '{{ air_temp_out_C }} °C'),
    ('Heat Transfer Area',        '{{ surface_area_m2 }} m²'),
    ('Pipe Outside Diameter',     '{{ pipe_dia_mm }} mm'),
    ('Pipe Length (per bank)',    '{{ pipe_length_m }} m'),
    ('Pipe Wall Thickness',       '{{ pipe_thick_mm }} mm'),
]

# ── Annexure I — Scope of Supply (recup) ────────────────────────────────────
_RECUP_SCOPE_ITEMS = [
    'Recuperator (Hot & Cold Bank) with tube plate, tubes, and supporting frame',
    'Material of Construction as per Annexure I (Material of Construction table above)',
    'Air inlet & outlet ducts with matching flanges and gaskets',
    'Combustion air receiving box with all internal baffles',
    'Bottom duct with expansion joints',
    'Pipe holding plates and supporting structure',
    'MS Side Hood (with insulation provision)',
    'Thermocouple (Type-K) for flue gas temperature monitoring',
    'Nut, bolt, washer, and matching flanges',
    'Painting and Packing (heat-resistant outer paint)',
]


def _make_moc_row(table, kind: str, label: str | None, value: str | None):
    new_row = table.add_row()
    cells = new_row.cells
    if kind in ('header', 'section'):
        cells[0].merge(cells[1])
        cells[0].text = label
        for p in cells[0].paragraphs:
            for r in p.runs:
                r.bold = True
    elif kind == 'item':
        cells[0].text = label
        cells[1].text = value or ''
    return new_row._element


def _replace_spec_table_with_designing_params(table) -> None:
    """Wipe table rows (except header) and re-fill with recup Designing
    Parameters. The first row is kept as the 'Parameter | Specification'
    header from VLPH."""
    tbl_xml = table._element
    # Keep header row (index 0), strip the rest.
    for tr in list(tbl_xml.findall(qn('w:tr')))[1:]:
        tbl_xml.remove(tr)
    # Replace header text with a recup-friendly banner.
    head = table.rows[0]
    head.cells[0].text = 'Recuperator Designing Parameters'
    head.cells[1].text = ''
    head.cells[0].merge(head.cells[1])
    for p in head.cells[0].paragraphs:
        for r in p.runs:
            r.bold = True
    # Add new rows.
    for label, value in _DESIGNING_PARAMS:
        row = table.add_row()
        row.cells[0].text = label
        row.cells[1].text = value


def _insert_moc_table_before(designing_table) -> None:
    """Build a brand-new MoC table (cloning the Designing Params table's
    style) and insert it BEFORE designing_table in the document body."""
    src_xml = designing_table._element
    moc_xml = deepcopy(src_xml)
    # Wipe all rows in clone; we'll add fresh.
    for tr in list(moc_xml.findall(qn('w:tr'))):
        moc_xml.remove(tr)
    # Insert the clone before designing_table.
    src_xml.addprevious(moc_xml)
    # Add a separating paragraph BETWEEN the MoC and Designing Params.
    sep = OxmlElement('w:p')
    src_xml.addprevious(sep)
    # Re-wrap the clone via a Table object so .add_row works.
    from docx.table import Table
    moc_table = Table(moc_xml, designing_table._parent)
    # Header row first
    moc_table.add_row()
    moc_table.add_row()  # we'll merge-fix in _make_moc_row
    # Actually clearer: build rows one at a time
    for tr in list(moc_xml.findall(qn('w:tr'))):
        moc_xml.remove(tr)
    for kind, label, value in _MOC_ROWS:
        _make_moc_row(moc_table, kind, label, value)


def _replace_scope_table_with_recup(table) -> None:
    """Replace the Annexure I 12-row Scope of Supply rows with recup
    components. The first row (header) stays untouched.

    VLPH structure for this table was:
      row 0: 'S. No. | Item Description'  (header)
      row 1: equipment name banner row (merged)
      rows 2-11: numbered scope items
    """
    tbl_xml = table._element
    # Drop rows 1+ (keep the 'S. No. | Item Description' header)
    for tr in list(tbl_xml.findall(qn('w:tr')))[1:]:
        tbl_xml.remove(tr)
    # Add equipment banner row (merged across both cols)
    banner = table.add_row()
    banner.cells[0].merge(banner.cells[1])
    banner.cells[0].text = '{{ equipment_name }}  —  {{ recup_qty }}'
    for p in banner.cells[0].paragraphs:
        for r in p.runs:
            r.bold = True
    # Numbered scope items
    for idx, item in enumerate(_RECUP_SCOPE_ITEMS, start=1):
        row = table.add_row()
        row.cells[0].text = str(idx)
        row.cells[1].text = item


def _delete_vlph_scope_body(doc: Document) -> None:
    """Walk the body looking for the 'SCOPE OF SUPPLY' paragraph (the
    inline one — i.e. NOT the 'ANNEXURE I — SCOPE OF SUPPLY' heading)
    and delete every element from there up to (but not including) the
    'ANNEXURE I' heading. Replace with a short recup operation block."""
    body = doc.element.body
    elements = list(body)

    start_idx = None
    end_idx = None
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

    # Remember the anchor (the ANNEXURE I element); we'll add replacement
    # paragraphs before it.
    anchor = elements[end_idx]

    # Delete elements [start_idx .. end_idx)
    for el in elements[start_idx:end_idx]:
        body.remove(el)

    # Insert recup operation block right before the ANNEXURE I anchor.
    def _para(text: str = '', bold: bool = False) -> OxmlElement:
        p = OxmlElement('w:p')
        if text:
            r = OxmlElement('w:r')
            if bold:
                rPr = OxmlElement('w:rPr')
                b = OxmlElement('w:b'); rPr.append(b); r.append(rPr)
            t = OxmlElement('w:t')
            t.text = text
            r.append(t)
            p.append(r)
        return p

    recup_body = [
        _para('SCOPE OF SUPPLY', bold=True),
        _para('Our scope of supply will cover design, engineering, '
              'manufacturing, supply, and supervision for erection & '
              'commissioning of the proposed Recuperator (Waste Heat '
              'Recovery System) for {{ application }}.'),
        _para('OPERATION', bold=True),
        _para('The Recuperator is a heat-recovery exchanger that uses '
              'the hot flue gases leaving the furnace to preheat the '
              'combustion air being supplied to the burner. The hot '
              'flue gas passes through the hot bank tubes, while '
              'ambient combustion air flows around them in a counter-'
              'cross-flow arrangement. The preheated air reduces fuel '
              'consumption and improves overall furnace efficiency.'),
        _para('Design and material of construction details are listed '
              'in the Material of Construction sub-table above. '
              'Designing parameters and process flows are listed in '
              'the Recuperator Designing Parameters table.'),
        _para('3D Image of the Proposed Recuperator', bold=True),
        _para('{{ image_recup_side }}'),
        _para(),
        _para('{{ image_recup_front }}'),
        _para(),
        _para('{{ image_recup_top }}'),
        _para(),
    ]
    # addprevious inserts BEFORE the anchor — iterate in normal order so
    # the first item lands earliest in the document.
    for el in recup_body:
        anchor.addprevious(el)

    # Also flip the section title "VERTICAL LADLE PREHEATERS AND DRYERS"
    for el in body.iter(qn('w:p')):
        for t in el.iter(qn('w:t')):
            if t.text and 'VERTICAL LADLE PREHEATERS' in t.text:
                t.text = 'RECUPERATOR (Waste Heat Recovery System)'


def _rebuild_price_schedule(doc: Document) -> None:
    """Find the 3-row Price Schedule (target by 'ITEM DESCRIPTION' header)
    and expand into single/full toggled rows + supervision + summary."""
    table = None
    for t in doc.tables:
        if not t.rows or len(t.rows[0].cells) < 2:
            continue
        if 'ITEM DESCRIPTION' in t.rows[0].cells[1].text.upper():
            table = t
            break
    if table is None:
        print('Price Schedule table not found — skipping rebuild.')
        return

    tbl_xml = table._element
    # Strip rows 1+ (keep the header 'S. No. | Item Description | Qty | Unit Price | Total Price')
    for tr in list(tbl_xml.findall(qn('w:tr')))[1:]:
        tbl_xml.remove(tr)

    def add(c1, c2, c3, c4, c5, *, bold=False):
        r = table.add_row()
        r.cells[0].text = c1
        r.cells[1].text = c2
        r.cells[2].text = c3
        r.cells[3].text = c4
        r.cells[4].text = c5
        if bold:
            for cell in r.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.bold = True

    # ── style: single ─────────────────────────────────────────────────
    add("{%tr if price_schedule_style == 'single' %}", "", "", "", "")
    add("1.", "Recuperator for {{ application }}",
        "{{ recup_qty }}", "{{ recup_unit_price }}", "{{ recup_total_price }}")
    add("{%tr endif %}", "", "", "", "")

    # ── style: full ────────────────────────────────────────────────────
    add("{%tr if price_schedule_style == 'full' %}", "", "", "", "")
    add("{%tr for r in bom_rows %}", "", "", "", "")
    add("{{ r.sno }}", "{{ r.item }}", "{{ r.qty }}",
        "{{ r.unit_price }}", "{{ r.total }}")
    add("{%tr endfor %}", "", "", "", "")
    add("{%tr if supervision_include %}", "", "", "", "")
    add("", "Supervision Charges for Erection & Commissioning "
        "(Erection by Client, Supervision by ENCON)", "",
        "{{ supervision_rate }}", "{{ supervision_note }}")
    add("{%tr endif %}", "", "", "", "")
    add("", "Bought-out Items Total", "", "", "{{ bought_out_total }}")
    add("", "ENCON Items Total",      "", "", "{{ encon_total }}")
    add("", "GRAND TOTAL",            "", "", "{{ grand_total }}", bold=True)
    add("{%tr endif %}", "", "", "", "")

    # Footer: amount-in-words spanning 4 cells
    footer = table.add_row()
    f = footer.cells
    merged = f[0].merge(f[1]).merge(f[2]).merge(f[3])
    merged.text = '{{ grand_total_in_words }}'
    f[4].text = '{{ grand_total }}'
    for p in merged.paragraphs:
        for run in p.runs:
            run.bold = True


def main() -> None:
    if not os.path.exists(SOURCE):
        raise SystemExit(f"missing: {SOURCE}")

    # Backup current target if it exists (caller already does this via the
    # _v2_backup file, but be defensive).
    if os.path.exists(TARGET) and not os.path.exists(TARGET + '.bak'):
        shutil.copy(TARGET, TARGET + '.bak')

    shutil.copy(SOURCE, TARGET)
    doc = Document(TARGET)
    print(f'Cloned {SOURCE} -> {TARGET}: '
          f'{len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables')

    # 1. Delete VLPH scope-of-supply body, replace with recup operation
    #    paragraphs + 3D image placeholders.
    _delete_vlph_scope_body(doc)

    # 2. Replace T5 (Tech Specs 24-row) with recup Designing Parameters.
    #    After the previous step, table indices haven't shifted (we only
    #    touched paragraphs), so doc.tables[5] is still the spec table.
    designing_table = doc.tables[5]
    _replace_spec_table_with_designing_params(designing_table)

    # 3. Insert Material of Construction sub-table BEFORE the
    #    Designing Parameters table.
    _insert_moc_table_before(designing_table)

    # 4. Replace Annexure I scope table content with recup items.
    #    After step 3, MoC took position 5 and Designing Params is now
    #    table 6, so the Scope of Supply table is at 7.
    # Locate by content to be robust:
    scope_table = None
    for t in doc.tables:
        if not t.rows or len(t.rows[0].cells) < 2:
            continue
        head = t.rows[0].cells[0].text.strip().upper()
        if head == 'S. NO.' and 'ITEM DESCRIPTION' not in t.rows[0].cells[1].text.upper():
            # candidate — the Annexure I Scope of Supply uses 'S. No.' header,
            # but so does the Reference List (50 rows) and Make List (1 row).
            # Distinguish by row count: scope is 12 rows.
            if 10 <= len(t.rows) <= 15:
                scope_table = t
                break
    if scope_table is not None:
        _replace_scope_table_with_recup(scope_table)
    else:
        print('Annexure I Scope of Supply table not found — leaving as-is.')

    # 5. Rebuild the Price Schedule (Annexure III) for single/full toggle.
    _rebuild_price_schedule(doc)

    doc.save(TARGET)
    print(f'Saved -> {TARGET}')


if __name__ == '__main__':
    main()
