"""
Start the recuperator's 3D image block on a fresh page.

At 120 mm wide the three views are 98–105 mm tall, so the first one already
does not fit under the technical-specification table it follows. Word pushes it
to the next page and leaves the "3D Image of the Proposed Recuperator" caption
stranded at the top of a mostly empty one — the gap in the offer.

Breaking the page before the caption keeps the caption with its images, and the
narrower images set in main.py (70 mm, giving 57–61 mm tall) let all three sit
on that page together instead of spilling onto a third.

Run:  python patch_recup_3d_block.py

Idempotent — a second run reports there is nothing to do.
"""

import os

from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TARGET = "Recup_Offer_Template.docx"
CAPTION = "3D Image of the Proposed Recuperator"


def patch():
    path = os.path.join(BASE_DIR, TARGET)
    doc = Document(path)

    caption = next((p for p in doc.paragraphs if CAPTION.lower() in p.text.lower()), None)
    if caption is None:
        print(f"  {TARGET}: no {CAPTION!r} caption found")
        return
    if caption.paragraph_format.page_break_before:
        print(f"  {TARGET}: the 3D block already starts a page")
        return

    caption.paragraph_format.page_break_before = True
    doc.save(path)
    print(f"  {TARGET}: the 3D block now starts on its own page")


if __name__ == "__main__":
    patch()
