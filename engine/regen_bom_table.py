"""Populate the regen offer's BILL OF MATERIAL table dynamically from the
actual BOM produced by ``bom.regen_builder.build_regen_df``.

The offer template ships with a static placeholder BOM table (header row
"Description | Qty."). At generation time we clear its body and re-emit it
faithfully from the computed BOM: one merged/bold row per BOM section, then an
item row (description + spec) and its real total quantity for every line.
"""
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH as WA
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DARK = "1F2937"     # body text colour (matches the rest of the offer)
SEC  = "EEF2F7"     # section-row shading


def _set_fill(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexc)
    tcPr.append(shd)


def _style_para(p, align):
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(2); pf.space_after = Pt(2); pf.line_spacing = 1.0


def _qty_label(q):
    try:
        q = int(round(float(q)))
    except (TypeError, ValueError):
        return str(q)
    return f"{q} No." if q == 1 else f"{q} Nos."


def _desc(name, spec):
    name = (name or "").strip()
    spec = (spec or "").strip()
    return f"{name} — {spec}" if spec else name


def _add_section_row(tbl):
    row = tbl.add_row()
    cell = row.cells[0].merge(row.cells[1])
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_fill(cell, SEC)
    p = cell.paragraphs[0]; _style_para(p, WA.LEFT)
    return cell, p


def _add_item_row(tbl):
    row = tbl.add_row()
    for c in row.cells:
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return row.cells[0], row.cells[1]


def _run(p, text, *, bold=False, size=12):
    r = p.add_run(text)
    r.bold = bold; r.font.size = Pt(size); r.font.color.rgb = RGBColor.from_string(DARK)
    return r


def fill_bom_table(doc_path, df):
    """Rewrite the BILL OF MATERIAL table in ``doc_path`` from BOM DataFrame ``df``.

    Returns True if the table was found and filled, else False.
    """
    d = Document(doc_path)

    tbl = None
    for t in d.tables:
        if (len(t.columns) == 2 and t.rows
                and t.rows[0].cells[0].text.strip().lower().startswith("description")):
            tbl = t
            break
    if tbl is None:
        return False

    # keep the styled header row, drop every existing body row
    for r in list(tbl.rows[1:]):
        r._tr.getparent().remove(r._tr)

    # faithful dump: section header row, then its item rows, in BOM order
    for section in df["SECTION"].drop_duplicates().tolist():
        _, sp = _add_section_row(tbl)
        _run(sp, str(section), bold=True)
        sub = df[df["SECTION"] == section]
        for _, row in sub.iterrows():
            dc, qc = _add_item_row(tbl)
            dp = dc.paragraphs[0]; _style_para(dp, WA.LEFT)
            _run(dp, _desc(row["ITEM NAME"], row.get("SPECIFICATION", "")))
            qp = qc.paragraphs[0]; _style_para(qp, WA.CENTER)
            _run(qp, _qty_label(row["QTY"]))

    d.save(doc_path)
    return True
