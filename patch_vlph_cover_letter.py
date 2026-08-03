"""
Drop the "COVER LETTER" heading from the vertical (VLPH) offer template.

The heading sat directly above "To," and added nothing — the letter is self
evidently a letter, and no entry in TABLE OF CONTENTS / LIST OF ANNEXURES
refers to it, so removing it leaves nothing dangling.

Run:  python patch_vlph_cover_letter.py

Idempotent — a second run finds nothing to remove.
"""

import os

from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = ["Offer_Template.docx"]
HEADING = "COVER LETTER"


def patch(path):
    d = Document(path)
    removed = 0
    for p in list(d.paragraphs):
        if p.text.strip().upper() == HEADING:
            p._element.getparent().remove(p._element)
            removed += 1
    if removed:
        d.save(path)
    print(f"  {os.path.basename(path)}: removed {removed} '{HEADING}' heading(s)"
          f"{' (already gone)' if not removed else ''}")


if __name__ == "__main__":
    for name in TEMPLATES:
        patch(os.path.join(BASE_DIR, name))
