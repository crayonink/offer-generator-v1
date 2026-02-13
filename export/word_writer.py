# word writer
"""
Word offer writer utility
Uses docxtpl to render the final offer document
"""

from docxtpl import DocxTemplate
import io
from typing import Dict


def generate_word_offer(
    template_path: str,
    context: Dict[str, str],
) -> io.BytesIO:
    """
    Renders a Word offer document from a template.

    Parameters
    ----------
    template_path : str
        Path to the .docx template file
    context : dict
        Placeholder context for docxtpl

    Returns
    -------
    BytesIO
        In-memory Word file buffer
    """

    doc = DocxTemplate(template_path)
    doc.render(context)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer
