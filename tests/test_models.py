from datetime import datetime, timezone
from decimal import Decimal

from donation_tracker.models import NEEDS_REVIEW, OK, Attachment, Donation
from tests.helpers import email_record, rule_record


def test_attachment_is_pdf_by_content_type_or_extension():
    assert Attachment("x.bin", "application/pdf", b"").is_pdf
    assert Attachment("Receipt.PDF", "application/octet-stream", b"").is_pdf
    assert not Attachment("logo.png", "image/png", b"").is_pdf


def test_donation_countable_only_for_ok_receipts():
    em = email_record()
    ok_receipt = Donation(em, rule_record(role="receipt"), Decimal("10"), None, None, OK)
    review = Donation(em, rule_record(role="receipt"), None, None, None, NEEDS_REVIEW)
    confirmation = Donation(em, rule_record(role="confirmation"), Decimal("10"), None, None, OK)
    assert ok_receipt.is_countable
    assert not review.is_countable
    assert not confirmation.is_countable


def test_helpers_build_parseable_email_bytes():
    from tests.helpers import make_email
    raw = make_email(subject="Hello", text="Body", attachments=[("r.pdf", "application/pdf", b"%PDF")])
    assert b"Subject: Hello" in raw
    assert b"r.pdf" in raw
