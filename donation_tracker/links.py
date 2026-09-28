"""Links back to the original message in Gmail."""
from __future__ import annotations

from urllib.parse import quote


def gmail_url(account: str, message_id: str) -> str:
    if message_id.startswith("<no-id-"):
        return ""
    return f"https://mail.google.com/mail/u/{account}/#search/rfc822msgid:{quote(message_id.strip('<> '), safe='')}"
