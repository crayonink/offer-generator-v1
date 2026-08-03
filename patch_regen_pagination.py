"""
Targeted pagination and figure fixes for the regen offer templates.

Each one was asked for specifically — this is deliberately NOT the
break-before-every-heading treatment, which was tried and reverted:

1. Page break before "To," — the cover block and the letter ran together, so
   "To, / <client> / Kind Attn / address" sat crammed under the letterhead at
   the foot of page 1 and the letter resumed on page 2 mid-address at
   "Contact No.:". The whole letter now fits on one page.

2. Page break before "TECHNICAL OFFER" — it and REGENERATIVE BURNERS were
   trailing at the bottom of the letter page, stranded from their body text on
   the next page.

3. Close the gap between "TECHNICAL OFFER" and "REGENERATIVE BURNERS" — the
   second heading carried 14 pt space-before, leaving a blank line between the
   two headings.

4. Rule under the Fig 1 diagram — it ran straight into the "Fig 2" caption
   with nothing separating the two figures.

   The figure captions sit ABOVE their diagram, so both are set keep-with-next:
   "Fig 2: Parts of Regenerative Burner" was stranding at the foot of a page
   with its diagram overleaf.

5. Scale both figure diagrams to 80%. They were 6.20 in wide (92% of the text
   width) and 4.30 / 3.33 in tall, so the pair alone overran a page.

Run once against the gas and oil templates, then rebuild the dual template
(it is derived from the gas one):

    python patch_regen_pagination.py
    python build_regen_dual_template.py

Idempotent.
"""

import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = ["Regen_Offer_Template.docx", "Regen_Oil_Offer_Template.docx"]

PAGE_BREAK_BEFORE = ["To,", "TECHNICAL OFFER"]
CLOSE_GAP_BEFORE = "REGENERATIVE BURNERS"
FIG1_CAPTION = "Fig 1:"
# Captions sit ABOVE their diagram, so they must keep-with-next — otherwise a
# caption strands at the foot of a page with its figure overleaf.
FIG_CAPTIONS = ("Fig 1:", "Fig 2:")
# An absolute target (80% of the original 6.20 in) rather than a scale factor,
# so re-running the script doesn't shrink the diagrams again. Aspect ratio kept.
FIGURE_TARGET_WIDTH_IN = 4.96
FIGURE_MIN_WIDTH_IN = 4.0    # only the big diagrams, not the logo/award images


def _set_page_break_before(p):
    pPr = p._element.get_or_add_pPr()
    for old in pPr.findall(qn("w:pageBreakBefore")):
        pPr.remove(old)
    pPr.append(OxmlElement("w:pageBreakBefore"))


def _rule_below(p):
    """Draw a horizontal line under this paragraph."""
    pPr = p._element.get_or_add_pPr()
    for old in pPr.findall(qn("w:pBdr")):
        pPr.remove(old)
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")          # eighths of a point -> 0.75 pt
    bottom.set(qn("w:space"), "6")       # gap between the image and the rule
    bottom.set(qn("w:color"), "808080")
    bdr.append(bottom)
    pPr.append(bdr)


def patch(path):
    d = Document(path)
    done, gap, rule, kept = [], False, False, 0
    after_fig1 = False
    for p in d.paragraphs:
        t = p.text.strip()
        if t in PAGE_BREAK_BEFORE and t not in done:
            _set_page_break_before(p)
            done.append(t)
        elif t == CLOSE_GAP_BEFORE and not gap:
            # Sits directly under TECHNICAL OFFER — no blank line between them.
            p.paragraph_format.space_before = Pt(0)
            gap = True
        # The Fig 1 diagram is the paragraph right after its caption; put the
        # rule under it so it doesn't run straight into the Fig 2 caption.
        if t.startswith(FIG_CAPTIONS):
            p.paragraph_format.keep_with_next = True
            kept += 1
        if after_fig1 and p._element.findall(".//" + qn("w:drawing")):
            _rule_below(p)
            rule = True
            after_fig1 = False
        elif t.startswith(FIG1_CAPTION):
            after_fig1 = True

    scaled = 0
    target = Inches(FIGURE_TARGET_WIDTH_IN)
    for shape in d.inline_shapes:
        if Emu(shape.width).inches < FIGURE_MIN_WIDTH_IN:
            continue                      # logo / award image — leave alone
        ratio = shape.height / shape.width
        shape.width = target
        shape.height = Emu(int(target * ratio))
        scaled += 1

    missing = [x for x in PAGE_BREAK_BEFORE if x not in done]
    note = f"  (not found: {', '.join(missing)})" if missing else ""
    print(f"  OK  {os.path.basename(path)} — breaks: {', '.join(done) or 'none'};"
          f" gap closed: {gap}; rule after Fig 1: {rule};"
          f" captions kept with figure: {kept}; figures scaled: {scaled}{note}")
    d.save(path)


if __name__ == "__main__":
    for name in TEMPLATES:
        patch(os.path.join(BASE_DIR, name))
    print("Now run: python build_regen_dual_template.py")
