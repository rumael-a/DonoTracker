"""Stream Email records out of a Google Takeout .mbox file.

decode_str / body extraction / attachment walking are adapted from
archived_attempts/attempt2 (parse_donations_by_year.py, extract_attachments.py).
"""
from __future__ import annotations

import email
import hashlib
import html
import mailbox
import re
from collections.abc import Callable, Iterator
from datetime import datetime, timezone, tzinfo
from email.header import decode_header
from email.message import Message
from email.parser import BytesHeaderParser
from email.utils import parsedate_to_datetime
from pathlib import Path

from .models import Attachment, Email

_HEADER_PARSER = BytesHeaderParser()
_EXT_BY_TYPE = {
    "application/pdf": ".pdf",
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
}
_MONTHS = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()


def decode_str(value) -> str:
    if value is None:
        return ""
    try:
        parts = decode_header(str(value))
    except Exception:
        return str(value)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            safe_enc = enc if enc and enc.lower() not in ("unknown-8bit", "unknown") else "utf-8"
            try:
                result.append(part.decode(safe_enc, errors="replace"))
            except LookupError:
                result.append(part.decode("utf-8", errors="replace"))
        else:
            result.append(part)
    return " ".join("".join(result).split())


def parse_date(raw: str, tz: tzinfo | None = None) -> datetime | None:
    """Parse a Date header; the result is in `tz` (None = this computer's local time) so the
    tax year follows the donor's calendar, not the sender's timezone."""
    if not raw:
        return None
    try:
        dt = parsedate_to_datetime(raw)
    except (TypeError, ValueError, IndexError):
        try:
            dt = datetime.fromisoformat(raw.strip())
        except ValueError:
            match = re.search(r"\b(\d{1,2}) (%s) (\d{4})\b" % "|".join(_MONTHS), raw)
            if not match:
                return None
            dt = datetime(int(match.group(3)), _MONTHS.index(match.group(2)) + 1, int(match.group(1)))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(tz)


def html_to_text(markup: str) -> str:
    markup = re.sub(r"(?is)<(style|script|head)[^>]*>.*?</\1>", " ", markup)
    markup = re.sub(r"(?i)<br\s*/?>|</(p|div|tr|li|h\d|table)>", "\n", markup)
    text = html.unescape(re.sub(r"<[^>]+>", " ", markup))
    text = re.sub(r"[ \t\r\f\v\xa0]+", " ", text)
    return re.sub(r"\n\s*\n+", "\n", text).strip()


def _decode_payload(part: Message) -> str | None:
    payload = part.get_payload(decode=True)
    if payload is None:
        return None
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except LookupError:
        return payload.decode("utf-8", errors="replace")


def _bodies(msg: Message) -> tuple[str, str | None]:
    plain = None
    markup = None
    for part in msg.walk():
        if part.is_multipart() or part.get_content_disposition() == "attachment":
            continue
        ctype = part.get_content_type()
        if ctype == "text/plain" and plain is None:
            plain = _decode_payload(part)
        elif ctype == "text/html" and markup is None:
            markup = _decode_payload(part)
    pieces = []
    if markup:
        pieces.append(html_to_text(markup))
    if plain:
        pieces.append(plain)
    return "\n".join(pieces), markup


def _attachments(msg: Message) -> list[Attachment]:
    found = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_type()
        disposition = part.get_content_disposition()
        if disposition != "attachment" and ctype != "application/pdf":
            continue
        data = part.get_payload(decode=True)
        if not data:
            continue
        filename = decode_str(part.get_filename()) or "attachment" + _EXT_BY_TYPE.get(ctype, ".bin")
        found.append(Attachment(filename, ctype, data))
    return found


def _build_email(msg: Message, raw: bytes, account: str, date: datetime) -> Email:
    message_id = (msg.get("Message-ID") or "").strip()
    if not message_id:
        message_id = "<no-id-" + hashlib.sha1(raw).hexdigest()[:16] + ">"
    body_text, body_html = _bodies(msg)
    return Email(
        account=account,
        message_id=message_id,
        date=date,
        sender=decode_str(msg.get("From")),
        subject=decode_str(msg.get("Subject")),
        body_text=body_text,
        body_html=body_html,
        attachments=_attachments(msg),
        raw=raw,
    )


def iter_emails(
    mbox_path: Path,
    account: str,
    years: set[int],
    on_error: Callable[[str], None] | None = None,
) -> Iterator[Email]:
    box = mailbox.mbox(str(mbox_path), create=False)
    try:
        for key in box.iterkeys():
            try:
                raw = box.get_bytes(key)
                headers = _HEADER_PARSER.parsebytes(raw)
                date = parse_date(headers.get("Date", ""))
                if date is None:
                    if on_error:
                        on_error(f"{mbox_path.name} message {key}: no parseable Date header")
                    continue
                if date.year not in years:
                    continue
                parsed = _build_email(email.message_from_bytes(raw), raw, account, date)
            except Exception as exc:  # malformed message: skip and report
                if on_error:
                    on_error(f"{mbox_path.name} message {key}: {exc!r}")
                continue
            yield parsed
    finally:
        box.close()
