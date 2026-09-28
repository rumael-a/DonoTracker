"""Text extraction from PDF attachments."""
from __future__ import annotations

import logging
from io import BytesIO

from pypdf import PdfReader

from .models import Attachment

logging.getLogger("pypdf").setLevel(logging.ERROR)


def pdf_to_text(data: bytes) -> str | None:
    try:
        reader = PdfReader(BytesIO(data))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception:
        return None
    return text if text.strip() else None


def attachments_text(attachments: list[Attachment]) -> tuple[str, bool]:
    texts = []
    textless = False
    for attachment in attachments:
        if not attachment.is_pdf:
            continue
        text = pdf_to_text(attachment.data)
        if text is None:
            textless = True
        else:
            texts.append(text)
    return "\n".join(texts), textless
