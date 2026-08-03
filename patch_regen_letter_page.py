"""
Start the regen cover letter on its own page.

The cover block (logo, OFFER FOR, project table, letterhead) and the letter ran
together, so "To, / <client> / Kind Attn / address" sat crammed under the
letterhead at the bottom of page 1 and the letter continued on page 2 starting
mid-address at "Contact No.:". Breaking before "To," puts the whole letter on
one page.

Run once against the gas and oil templates, then rebuild the dual template
(it is derived from the gas one):

    python patch_regen_letter_page.py
    python build_regen_dual_template.py

Idempotent.
"""

import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = ["Regen_Offer_Template.docx", "Regen_Oil_Offer_Template.docx"]

LETTER_START = "To,"


def patch(path):
    d = Document(path)
    hit = None
    for p in d.paragraphs:
        if p.text.strip() == LETTER_START:
            hit = p
            break
    if hit is None:
        print(f"  SKIP {os.path.basename(path)} — no '{LETTER_START}' paragraph")
        return
    pPr = hit._element.get_or_add_pPr()
    for old in pPr.findall(qn("w:pageBreakBefore")):
        pPr.remove(old)
    pPr.append(OxmlElement("w:pageBreakBefore"))
    d.save(path)
    print(f"  OK  {os.path.basename(path)} — letter starts on its own page")


if __name__ == "__main__":
    for name in TEMPLATES:
        patch(os.path.join(BASE_DIR, name))
    print("Now run: python build_regen_dual_template.py")
