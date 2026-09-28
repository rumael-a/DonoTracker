from datetime import timezone

from donation_tracker.mbox_reader import decode_str, html_to_text, iter_emails, parse_date
from tests.helpers import make_email, make_mbox, make_text_pdf


def test_decode_str_handles_encoded_words_and_unknown_8bit():
    assert decode_str("=?utf-8?b?RG9uYXRpb24g4pyU?=") == "Donation ✔"
    assert decode_str("=?unknown-8bit?q?Re=E7u?=").startswith("Re")
    assert decode_str(None) == ""


def test_parse_date_returns_aware_datetime_or_none():
    dt = parse_date("Fri, 14 Mar 2025 10:00:00 -0400")
    assert dt.year == 2025 and dt.tzinfo is not None
    naive = parse_date("Fri, 14 Mar 2025 10:00:00 -0000")
    assert naive.tzinfo is not None
    assert parse_date("garbage") is None
    assert parse_date("") is None


def test_html_to_text_drops_style_and_tags():
    markup = "<style>.a{color:red}</style><p>Amount:&nbsp;<b>$50.00</b></p><br>Thanks"
    text = html_to_text(markup)
    assert "color" not in text
    assert "Amount: $50.00" in text
    assert "Thanks" in text


def test_iter_emails_filters_years_and_parses_fields(tmp_path):
    pdf = make_text_pdf("Receipt 123")
    raws = [
        make_email(subject="Tax receipt 2025", html="<p>Gift $25.00</p>", message_id="<a@x>",
                   attachments=[("receipt.pdf", "application/pdf", pdf)]),
        make_email(subject="Old", date="Mon, 01 Jan 2024 09:00:00 +0000", message_id="<b@x>", text="old"),
        make_email(subject="Next year", date="Mon, 02 Feb 2026 09:00:00 +0000", message_id="<c@x>", text="n"),
    ]
    box = make_mbox(tmp_path / "t.mbox", raws)
    emails = list(iter_emails(box, "abbasirumael@gmail.com", {2025, 2026}))
    assert [e.message_id for e in emails] == ["<a@x>", "<c@x>"]
    first = emails[0]
    assert first.account == "abbasirumael@gmail.com"
    assert first.sender == "SickKids Foundation <receipts@sickkids.ca>"
    assert first.subject == "Tax receipt 2025"
    assert "Gift $25.00" in first.body_text
    assert first.body_html is not None
    assert [a.filename for a in first.attachments] == ["receipt.pdf"]
    assert first.attachments[0].data == pdf
    assert first.raw.startswith(b"From:") or b"Subject: Tax receipt 2025" in first.raw


def test_iter_emails_synthesizes_missing_message_id(tmp_path):
    box = make_mbox(tmp_path / "t.mbox", [make_email(message_id=None, text="x")])
    (email,) = list(iter_emails(box, "a@gmail.com", {2025}))
    assert email.message_id.startswith("<no-id-")


def test_iter_emails_reports_dateless_messages(tmp_path):
    box = make_mbox(tmp_path / "t.mbox", [make_email(date=None, text="x")])
    errors = []
    assert list(iter_emails(box, "a@gmail.com", {2025}, errors.append)) == []
    assert len(errors) == 1 and "Date" in errors[0]


def test_iter_emails_skips_inline_images_but_keeps_named_attachments(tmp_path):
    raw = make_email(text="x", attachments=[("image001.png", "image/png", b"\x89PNG")])
    box = make_mbox(tmp_path / "t.mbox", [raw])
    (email,) = list(iter_emails(box, "a@gmail.com", {2025}))
    # add_attachment marks disposition=attachment, so a named PNG attachment is kept
    assert [a.filename for a in email.attachments] == ["image001.png"]
    assert not email.attachments[0].is_pdf
