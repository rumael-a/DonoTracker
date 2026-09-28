"""Save each donation email (.eml, readable .html) and its attachments.

safe_name and the HTML wrapper are adapted from archived_attempts/attempt2.
"""
from __future__ import annotations

import html as html_mod
import re
from pathlib import Path

from .models import Donation, Email


def safe_name(s: str, maxlen: int = 40) -> str:
    s = re.sub(r'[<>:"/\\|?*#%\x00-\x1f]', "_", s)
    return s[:maxlen].strip("_. ")


def org_slug(name: str) -> str:
    return safe_name(re.sub(r"\s+", "_", name.strip()), 40)


def _safe_filename(filename: str) -> str:
    path = Path(filename)
    # Sanitize extension name without the dot, then add dot back
    ext = path.suffix[1:] if path.suffix else ""
    suffix = ("." + safe_name(ext, 9)) if ext else ""
    return (safe_name(path.stem, 70) or "attachment") + suffix


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def email_to_html(email: Email) -> str:
    if email.body_html:
        body = email.body_html
    else:
        body = (
            "<pre style='font-family:Arial;font-size:14px;white-space:pre-wrap'>"
            f"{html_mod.escape(email.body_text)}</pre>"
        )
    esc = html_mod.escape
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: cid:; style-src 'unsafe-inline'; font-src data:">
<title>{esc(email.subject)}</title>
<style>
  body {{ margin: 0; font-family: Arial, sans-serif; }}
  .meta {{ background:#f5f5f5; padding:12px 20px; border-bottom:1px solid #ddd; font-size:13px; color:#444; }}
  .meta strong {{ color:#111; }}
  .body-wrap {{ padding:20px; }}
</style>
</head><body>
<div class="meta">
  <div><strong>Date:</strong> {email.date:%Y-%m-%d %H:%M %z}</div>
  <div><strong>From:</strong> {esc(email.sender)}</div>
  <div><strong>Mailbox:</strong> {esc(email.account)}</div>
  <div><strong>Subject:</strong> {esc(email.subject)}</div>
  <div><strong>Message-ID:</strong> {esc(email.message_id)}</div>
</div>
<div class="body-wrap">{body}</div>
</body></html>"""


def save_evidence(donation: Donation, evidence_root: Path) -> None:
    email = donation.email
    folder = unique_path(evidence_root / f"{email.date:%Y-%m-%d}_{org_slug(donation.rule.name)}")
    folder.mkdir(parents=True)
    donation.email_eml_path = folder / "email.eml"
    donation.email_eml_path.write_bytes(email.raw)
    donation.email_html_path = folder / "email.html"
    donation.email_html_path.write_text(email_to_html(email), encoding="utf-8")
    for attachment in email.attachments:
        target = unique_path(folder / _safe_filename(attachment.filename))
        target.write_bytes(attachment.data)
        if attachment.is_pdf:
            donation.receipt_paths.append(target)
