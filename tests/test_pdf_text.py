from donation_tracker.models import Attachment
from donation_tracker.pdf_text import attachments_text, pdf_to_text
from tests.helpers import make_text_pdf


def test_pdf_to_text_reads_text_layer():
    assert "Official Receipt 555" in pdf_to_text(make_text_pdf("Official Receipt 555"))


def test_pdf_to_text_none_for_empty_or_garbage():
    assert pdf_to_text(make_text_pdf("")) is None
    assert pdf_to_text(b"not a pdf") is None


def test_attachments_text_joins_pdfs_and_flags_textless():
    atts = [
        Attachment("a.pdf", "application/pdf", make_text_pdf("Amount 10.00")),
        Attachment("scan.pdf", "application/pdf", make_text_pdf("")),
        Attachment("logo.png", "image/png", b"\x89PNG"),
    ]
    text, textless = attachments_text(atts)
    assert "Amount 10.00" in text
    assert textless is True


def test_attachments_text_no_pdfs():
    assert attachments_text([]) == ("", False)
