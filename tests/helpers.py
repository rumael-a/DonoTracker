"""Builders for test emails, mboxes, PDFs and records."""
from __future__ import annotations

import mailbox
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

from donation_tracker.models import Attachment, Email, Rule


def make_email(
    *,
    sender: str = "SickKids Foundation <receipts@sickkids.ca>",
    to: str = "abbasirumael@gmail.com",
    subject: str = "Your tax receipt",
    date: str | None = "Fri, 14 Mar 2025 10:00:00 -0400",
    text: str | None = None,
    html: str | None = None,
    message_id: str | None = "<abc123@sickkids.ca>",
    attachments: tuple | list = (),
) -> bytes:
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    if date is not None:
        msg["Date"] = date
    if message_id:
        msg["Message-ID"] = message_id
    if text is not None:
        msg.set_content(text)
        if html is not None:
            msg.add_alternative(html, subtype="html")
    elif html is not None:
        msg.set_content(html, subtype="html")
    else:
        msg.set_content("")
    for filename, ctype, data in attachments:
        maintype, subtype = ctype.split("/")
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
    return msg.as_bytes()


def make_mbox(path: Path, raws: list[bytes]) -> Path:
    box = mailbox.mbox(str(path))
    for raw in raws:
        box.add(raw)
    box.flush()
    box.close()
    return path


def make_text_pdf(text: str) -> bytes:
    """Minimal one-page PDF whose text layer is `text` (no parentheses/backslashes)."""
    stream = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode("latin-1") if text else b""
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for number, body in enumerate(objects, 1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % number + body + b"\nendobj\n"
    xref_at = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    for offset in offsets:
        out += b"%010d 00000 n \n" % offset
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objects) + 1, xref_at)
    return bytes(out)


def email_record(
    *,
    account: str = "abbasirumael@gmail.com",
    message_id: str = "<m1@sickkids.ca>",
    date: datetime = datetime(2025, 3, 14, 10, 0, tzinfo=timezone.utc),
    sender: str = "SickKids Foundation <receipts@sickkids.ca>",
    subject: str = "Your tax receipt",
    body_text: str = "",
    body_html: str | None = None,
    attachments: list[Attachment] | None = None,
    raw: bytes = b"From: x\r\n\r\nraw",
) -> Email:
    return Email(account, message_id, date, sender, subject, body_text, body_html, attachments or [], raw)


def rule_record(**overrides) -> Rule:
    values = dict(
        name="SickKids Foundation", country="CA", currency="CAD", role="receipt", match_from=["sickkids.ca"]
    )
    values.update(overrides)
    return Rule(**values)
