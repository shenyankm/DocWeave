"""Minimal Word fixtures for save, rendering and input-boundary tests."""
import dockiln as dk
from dockiln import light_document_model as ldm


def word_document(text):
    if isinstance(text, bytes):
        text = text.decode("utf-8")
    doc = dk.Document()
    doc.first_section.body.children = [
        ldm.Paragraph(children=[ldm.Run(text=line)]) for line in text.split("\n")
    ]
    return doc


def word_bytes(text):
    return word_document(text).to_bytes(dk.SaveFormat.DOCX)
