# Donation Tracker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A yearly-run offline Python CLI that reads Gmail Takeout `.mbox` exports, finds donation receipts with user-editable YAML rules, and writes a traceable Excel report plus a self-contained accountant folder.

**Architecture:** A `donation_tracker` package of small single-purpose modules (reader → matcher → extractor → evidence → report/pack), orchestrated by `pipeline.run()` and exposed through `donations.py` (`run`, `test-rule`). All configuration is data: `accounts.yaml`, `rules.yaml`, `detector.yaml`.

**Tech Stack:** Python 3.11+, PyYAML, openpyxl, pypdf, pytest; stdlib `mailbox`/`email`; Microsoft Edge headless for HTML→PDF.

**Spec:** `docs/superpowers/specs/2026-09-28-donation-tracker-design.md`

## Global Constraints

- Python 3.11+. Dependencies limited to: `pyyaml`, `openpyxl`, `pypdf`, `pytest` (+ stdlib). No AI/LLM, no network APIs.
- Platform: Windows 11; paths via `pathlib`; all text files written with `encoding="utf-8"`.
- Input is Takeout `.mbox` only. Never commit mail data: `*.mbox`, `abbasirumael/`, `itsrumzy/`, `rumaelabbasi/`, `archived_attempts/`, `output/`, `tests/fixtures/real/` are git-ignored.
- Allowed rule values: `country` ∈ {`CA`,`US`}; `currency` ∈ {`CAD`,`USD`}; `role` ∈ {`receipt`,`confirmation`}; `tax_year` ∈ {`same`,`previous`}; `extract` keys ∈ {`amount`,`receipt_no`,`charity_id`}.
- Status strings exactly `OK` and `NEEDS REVIEW`.
- Totals include only `role == receipt` and `status == OK`. No currency conversion.
- US flag note text exactly: `IRS: written acknowledgment required for $250+`.
- Gmail link format: `https://mail.google.com/mail/u/<account>/#search/rfc822msgid:<url-encoded Message-ID without <>>`.
- A run for year N reads messages dated N and N+1; N+1 messages kept only for rules with `tax_year: previous`. Needs-a-rule candidates come from year N only.
- Output tree: `output/<year>/{donations_<year>.xlsx, suggested_rules.yaml, run_log.txt, evidence/, for_accountant/{donations_<year>_summary.xlsx, CA/, US/}}`. Output year folder is deleted and rebuilt each run. Workbook hyperlinks are relative.
- Accountant pack contains only `role == receipt` and `status == OK` rows.

## File Structure

```
donations.py                    # thin entry: calls donation_tracker.cli.main
requirements.txt
pytest.ini
accounts.yaml                   # account email → mbox path
rules.yaml                      # one entry per organization/email type
detector.yaml                   # keyword lists for "Needs a rule" detection
donation_tracker/
  __init__.py
  models.py                     # Attachment, Email, Rule, Donation, status constants
  mbox_reader.py                # decode headers/bodies/attachments, iter_emails()
  pdf_text.py                   # pdf_to_text(), attachments_text()
  config.py                     # ConfigError, Account, load_accounts(), load_rules()
  matcher.py                    # rule_matches(), match_rule()
  extractor.py                  # extract_donation(), tax_year_for()
  detector.py                   # DetectorConfig, load_detector(), is_candidate(), suggested_rules_yaml()
  links.py                      # gmail_url()
  evidence.py                   # safe_name(), org_slug(), email_to_html(), save_evidence()
  pdf_render.py                 # EdgeRenderer, find_edge()
  report.py                     # ReportRow, Summary, row_from_donation(), summarize(), write_workbook()
  accountant.py                 # build_pack()
  pipeline.py                   # collect(), run(), RunResult
  cli.py                        # argparse: run, test-rule; console summary
tools/
  make_fixture.py               # dump one real email as an anonymized regression fixture
tests/
  helpers.py                    # make_email(), make_mbox(), make_text_pdf(), email_record(), rule_record()
  test_models.py
  test_mbox_reader.py
  test_pdf_text.py
  test_config.py
  test_matcher.py
  test_extractor.py
  test_detector.py
  test_evidence.py
  test_pdf_render.py
  test_report.py
  test_accountant.py
  test_pipeline.py
  test_cli.py
  test_real_rules.py            # parametrized over tests/fixtures/real/*.json
  fixtures/real/                # git-ignored anonymized real-email fixtures
```

---

### Task 0: Repository, environment, models, test helpers

**Files:**
- Create: `requirements.txt`, `pytest.ini`, `donation_tracker/__init__.py`, `donation_tracker/models.py`, `tests/__init__.py`, `tests/helpers.py`, `tests/test_models.py`
- Modify: `.gitignore` (already exists; add `tests/fixtures/real/`)

**Interfaces:**
- Produces: `models.Attachment(filename, content_type, data)` with `.is_pdf`; `models.Email(account, message_id, date, sender, subject, body_text, body_html, attachments, raw)`; `models.Rule(name, country, currency, role, match_from, match_subject=[], extract={}, charity_id=None, tax_year="same")`; `models.Donation(email, rule, amount, receipt_no, charity_id, status, notes=[], gmail_url="", email_html_path=None, email_eml_path=None, receipt_paths=[])` with `.is_countable`; constants `OK = "OK"`, `NEEDS_REVIEW = "NEEDS REVIEW"`.
- Produces (tests): `helpers.make_email(...) -> bytes`, `helpers.make_mbox(path, raws)`, `helpers.make_text_pdf(text) -> bytes`, `helpers.email_record(**kw) -> Email`, `helpers.rule_record(**kw) -> Rule`.

- [ ] **Step 1: Initialize git and commit the spec** (skip any part already done)

The `.gitignore` already exists at the repo root. Add one line to it: `tests/fixtures/real/`.

```powershell
git init
git add .gitignore docs
git status --short   # must list ONLY .gitignore and docs/... — no .mbox, no archived_attempts
git commit -m "Add donation tracker design spec and plan"
```

- [ ] **Step 2: Create environment files**

`requirements.txt`:
```
pyyaml>=6.0
openpyxl>=3.1
pypdf>=4.0
pytest>=8.0
```

`pytest.ini`:
```ini
[pytest]
testpaths = tests
pythonpath = .
```

Run:
```powershell
python --version          # must be 3.11 or newer
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

- [ ] **Step 3: Write the failing test** — `tests/test_models.py`

```python
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
```

- [ ] **Step 4: Run test to verify it fails**

Run: `.venv\Scripts\python -m pytest tests/test_models.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'donation_tracker'`

- [ ] **Step 5: Write implementation**

`donation_tracker/__init__.py`:
```python
"""Donation tracker: find donation receipts in Gmail Takeout exports."""
```

`tests/__init__.py`: empty file.

`donation_tracker/models.py`:
```python
"""Plain data records passed between modules."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal
from pathlib import Path

OK = "OK"
NEEDS_REVIEW = "NEEDS REVIEW"


@dataclass
class Attachment:
    filename: str
    content_type: str
    data: bytes

    @property
    def is_pdf(self) -> bool:
        return self.content_type == "application/pdf" or self.filename.lower().endswith(".pdf")


@dataclass
class Email:
    account: str
    message_id: str
    date: datetime  # always timezone-aware
    sender: str  # full decoded From header, e.g. 'Muslim Hands <mail@muslimhands.ca>'
    subject: str
    body_text: str
    body_html: str | None
    attachments: list[Attachment]
    raw: bytes


@dataclass
class Rule:
    name: str
    country: str
    currency: str
    role: str
    match_from: list[str]
    match_subject: list[str] = field(default_factory=list)
    extract: dict[str, re.Pattern] = field(default_factory=dict)
    charity_id: str | None = None
    tax_year: str = "same"


@dataclass
class Donation:
    email: Email
    rule: Rule
    amount: Decimal | None
    receipt_no: str | None
    charity_id: str | None
    status: str
    notes: list[str] = field(default_factory=list)
    gmail_url: str = ""
    email_html_path: Path | None = None
    email_eml_path: Path | None = None
    receipt_paths: list[Path] = field(default_factory=list)

    @property
    def is_countable(self) -> bool:
        return self.rule.role == "receipt" and self.status == OK
```

`tests/helpers.py`:
```python
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
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_models.py -v`
Expected: 3 passed

- [ ] **Step 7: Commit**

```powershell
git add requirements.txt pytest.ini .gitignore donation_tracker tests
git commit -m "feat: project scaffold, data models and test helpers"
```

---

### Task 1: mbox reader (reusing attempt2 helpers)

**Files:**
- Create: `donation_tracker/mbox_reader.py`
- Test: `tests/test_mbox_reader.py`

**Interfaces:**
- Consumes: `models.Email`, `models.Attachment`.
- Produces: `decode_str(value) -> str`; `parse_date(raw: str) -> datetime | None` (always tz-aware); `html_to_text(markup: str) -> str`; `iter_emails(mbox_path: Path, account: str, years: set[int], on_error: Callable[[str], None] | None = None) -> Iterator[Email]`.

- [ ] **Step 1: Write the failing tests** — `tests/test_mbox_reader.py`

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_mbox_reader.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'donation_tracker.mbox_reader'`

- [ ] **Step 3: Write implementation** — `donation_tracker/mbox_reader.py`

```python
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
from datetime import datetime, timezone
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


def parse_date(raw: str) -> datetime | None:
    if not raw:
        return None
    try:
        dt = parsedate_to_datetime(raw)
    except (TypeError, ValueError, IndexError):
        match = re.search(r"\b(\d{1,2}) (%s) (\d{4})\b" % "|".join(_MONTHS), raw)
        if not match:
            return None
        dt = datetime(int(match.group(3)), _MONTHS.index(match.group(2)) + 1, int(match.group(1)))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_mbox_reader.py -v`
Expected: 7 passed. If `test_decode_str_handles_encoded_words_and_unknown_8bit` fails on the unknown-8bit case, the fallback path is wrong — fix code, not the test.

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/mbox_reader.py tests/test_mbox_reader.py
git commit -m "feat: mbox reader with year filter, bodies and attachments"
```

---

### Task 2: PDF text extraction

**Files:**
- Create: `donation_tracker/pdf_text.py`
- Test: `tests/test_pdf_text.py`

**Interfaces:**
- Consumes: `models.Attachment`.
- Produces: `pdf_to_text(data: bytes) -> str | None` (None when unreadable or no text layer); `attachments_text(attachments: list[Attachment]) -> tuple[str, bool]` returning (joined text of PDFs, True if any PDF had no text).

- [ ] **Step 1: Write the failing tests** — `tests/test_pdf_text.py`

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_pdf_text.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation** — `donation_tracker/pdf_text.py`

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_pdf_text.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/pdf_text.py tests/test_pdf_text.py
git commit -m "feat: PDF attachment text extraction"
```

---

### Task 3: Config loading and validation (accounts.yaml, rules.yaml)

**Files:**
- Create: `donation_tracker/config.py`, `accounts.yaml`
- Test: `tests/test_config.py`

**Interfaces:**
- Consumes: `models.Rule`.
- Produces: `ConfigError(Exception)`; `Account(email: str, mbox: Path)`; `load_accounts(path: Path) -> list[Account]`; `load_rules(path: Path) -> list[Rule]`; `load_yaml(path: Path)` (raises ConfigError for missing file / bad YAML).

- [ ] **Step 1: Write the failing tests** — `tests/test_config.py`

```python
import pytest

from donation_tracker.config import ConfigError, load_accounts, load_rules

GOOD_RULES = """
- name: SickKids Foundation
  country: CA
  currency: CAD
  role: receipt
  match:
    from: ["@sickkids.ca"]
    subject: ["tax receipt"]
  extract:
    amount: 'Amount:\\s*\\$([\\d,.]+)'
  charity_id: "10742 5582 RR0001"
- name: MyTenNights annual
  country: US
  currency: USD
  role: receipt
  tax_year: previous
  match:
    from: mytennights.com
"""


def write(tmp_path, name, text):
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return path


def test_load_rules_parses_all_fields(tmp_path):
    rules = load_rules(write(tmp_path, "rules.yaml", GOOD_RULES))
    first, second = rules
    assert first.name == "SickKids Foundation"
    assert first.match_from == ["@sickkids.ca"]
    assert first.match_subject == ["tax receipt"]
    assert first.extract["amount"].search("Amount: $1,200.00").group(1) == "1,200.00"
    assert first.charity_id == "10742 5582 RR0001"
    assert first.role == "receipt" and first.tax_year == "same"
    assert second.match_from == ["mytennights.com"]  # string normalized to list
    assert second.tax_year == "previous"


def test_empty_rules_file_is_allowed(tmp_path):
    assert load_rules(write(tmp_path, "rules.yaml", "")) == []


@pytest.mark.parametrize(
    "body, message",
    [
        ("- name: X\n  country: FR\n  currency: CAD\n  match: {from: a}\n", "country"),
        ("- name: X\n  country: CA\n  currency: EUR\n  match: {from: a}\n", "currency"),
        ("- name: X\n  country: CA\n  currency: CAD\n  role: gift\n  match: {from: a}\n", "role"),
        ("- name: X\n  country: CA\n  currency: CAD\n  tax_year: next\n  match: {from: a}\n", "tax_year"),
        ("- name: X\n  country: CA\n  currency: CAD\n", "match"),
        ("- name: X\n  country: CA\n  currency: CAD\n  match: {subject: a}\n", "match.from"),
        ("- name: X\n  country: CA\n  currency: CAD\n  match: {from: a}\n  extract: {amount: '([0-9'}\n", "amount"),
        ("- name: X\n  country: CA\n  currency: CAD\n  match: {from: a}\n  extract: {amount: '[0-9]+'}\n", "capture group"),
        ("- name: X\n  country: CA\n  currency: CAD\n  match: {from: a}\n  extract: {total: '(x)'}\n", "total"),
        ("- name: X\n  country: CA\n  currency: CAD\n  match: {from: a}\n"
         "- name: X\n  country: CA\n  currency: CAD\n  match: {from: b}\n", "duplicate"),
        ("name: X\n", "list"),
        ("- [unclosed\n", "invalid YAML"),
    ],
)
def test_invalid_rules_raise_with_rule_name(tmp_path, body, message):
    with pytest.raises(ConfigError) as err:
        load_rules(write(tmp_path, "rules.yaml", body))
    assert message in str(err.value)


def test_load_accounts_resolves_relative_paths(tmp_path):
    (tmp_path / "acct").mkdir()
    mbox = tmp_path / "acct" / "All mail.mbox"
    mbox.write_bytes(b"")
    path = write(tmp_path, "accounts.yaml",
                 "accounts:\n  - email: a@gmail.com\n    mbox: acct/All mail.mbox\n")
    (account,) = load_accounts(path)
    assert account.email == "a@gmail.com"
    assert account.mbox == mbox


def test_load_accounts_missing_mbox(tmp_path):
    path = write(tmp_path, "accounts.yaml", "accounts:\n  - email: a@gmail.com\n    mbox: nope.mbox\n")
    with pytest.raises(ConfigError, match="a@gmail.com"):
        load_accounts(path)


def test_missing_file(tmp_path):
    with pytest.raises(ConfigError, match="not found"):
        load_rules(tmp_path / "missing.yaml")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_config.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation** — `donation_tracker/config.py`

```python
"""Load and validate accounts.yaml and rules.yaml."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import yaml

from .models import Rule

COUNTRIES = {"CA", "US"}
CURRENCIES = {"CAD", "USD"}
ROLES = {"receipt", "confirmation"}
TAX_YEARS = {"same", "previous"}
EXTRACT_KEYS = {"amount", "receipt_no", "charity_id"}


class ConfigError(Exception):
    pass


@dataclass
class Account:
    email: str
    mbox: Path


def load_yaml(path: Path):
    if not path.exists():
        raise ConfigError(f"{path}: file not found")
    try:
        return yaml.safe_load(path.read_text(encoding="utf-8"))
    except yaml.YAMLError as exc:
        raise ConfigError(f"{path.name}: invalid YAML: {exc}") from exc


def _as_list(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    return [str(v) for v in value]


def load_accounts(path: Path) -> list[Account]:
    data = load_yaml(path) or {}
    entries = data.get("accounts") if isinstance(data, dict) else None
    if not entries:
        raise ConfigError(f"{path.name}: needs an 'accounts' list")
    accounts = []
    for entry in entries:
        if not isinstance(entry, dict) or "email" not in entry or "mbox" not in entry:
            raise ConfigError(f"{path.name}: each account needs 'email' and 'mbox'")
        mbox = Path(entry["mbox"])
        if not mbox.is_absolute():
            mbox = path.parent / mbox
        if not mbox.exists():
            raise ConfigError(f"{path.name}: mbox not found for {entry['email']}: {mbox}")
        accounts.append(Account(str(entry["email"]), mbox))
    return accounts


def _compile(where: str, key: str, pattern: str) -> re.Pattern:
    try:
        compiled = re.compile(pattern, re.IGNORECASE)
    except re.error as exc:
        raise ConfigError(f"{where}: extract.{key} is not a valid regex: {exc}") from exc
    if compiled.groups < 1:
        raise ConfigError(f"{where}: extract.{key} needs a capture group (...) around the value")
    return compiled


def _rule(where: str, entry: dict) -> Rule:
    for key in ("name", "country", "currency", "match"):
        if key not in entry:
            raise ConfigError(f"{where}: missing '{key}'")
    role = entry.get("role", "receipt")
    tax_year = entry.get("tax_year", "same")
    checks = [("country", entry["country"], COUNTRIES), ("currency", entry["currency"], CURRENCIES),
              ("role", role, ROLES), ("tax_year", tax_year, TAX_YEARS)]
    for key, value, allowed in checks:
        if value not in allowed:
            raise ConfigError(f"{where}: {key} must be one of {sorted(allowed)}, got {value!r}")
    match = entry["match"]
    if not isinstance(match, dict) or not _as_list(match.get("from")):
        raise ConfigError(f"{where}: match.from is required")
    extract_raw = entry.get("extract") or {}
    unknown = set(extract_raw) - EXTRACT_KEYS
    if unknown:
        raise ConfigError(f"{where}: unknown extract key(s) {sorted(unknown)}; allowed {sorted(EXTRACT_KEYS)}")
    return Rule(
        name=str(entry["name"]),
        country=entry["country"],
        currency=entry["currency"],
        role=role,
        match_from=_as_list(match.get("from")),
        match_subject=_as_list(match.get("subject")),
        extract={key: _compile(where, key, str(p)) for key, p in extract_raw.items()},
        charity_id=str(entry["charity_id"]) if entry.get("charity_id") else None,
        tax_year=tax_year,
    )


def load_rules(path: Path) -> list[Rule]:
    data = load_yaml(path)
    if data is None:
        return []
    if not isinstance(data, list):
        raise ConfigError(f"{path.name}: top level must be a list of rules (each starting with '- name:')")
    rules: list[Rule] = []
    seen: set[str] = set()
    for number, entry in enumerate(data, 1):
        label = entry.get("name") if isinstance(entry, dict) else None
        where = f"{path.name} rule #{number}" + (f" ({label})" if label else "")
        if not isinstance(entry, dict):
            raise ConfigError(f"{where}: must be a mapping")
        rule = _rule(where, entry)
        if rule.name in seen:
            raise ConfigError(f"{where}: duplicate rule name {rule.name!r}")
        seen.add(rule.name)
        rules.append(rule)
    return rules
```

`accounts.yaml` (repo root):
```yaml
accounts:
  - email: abbasirumael@gmail.com
    mbox: abbasirumael/All mail Including Spam and Trash.mbox
  - email: itsrumzy@gmail.com
    mbox: itsrumzy/All mail Including Spam and Trash.mbox
  # - email: rumaelabbasi@gmail.com   # no donation receipts yet; add its mbox here when needed
  #   mbox: rumaelabbasi/All mail Including Spam and Trash.mbox
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_config.py -v`
Expected: all passed (17)

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/config.py tests/test_config.py accounts.yaml
git commit -m "feat: load and validate accounts and rules config"
```

---

### Task 4: Rule matcher

**Files:**
- Create: `donation_tracker/matcher.py`
- Test: `tests/test_matcher.py`

**Interfaces:**
- Consumes: `models.Email`, `models.Rule`.
- Produces: `rule_matches(rule: Rule, email: Email) -> bool`; `match_rule(email: Email, rules: list[Rule]) -> tuple[Rule | None, list[Rule]]` (first match, other matches).

- [ ] **Step 1: Write the failing tests** — `tests/test_matcher.py`

```python
from donation_tracker.matcher import match_rule, rule_matches
from tests.helpers import email_record, rule_record


def test_from_is_case_insensitive_substring_of_whole_header():
    email = email_record(sender="ICNA Relief Canada <receipts+acct_15x@stripe.com>")
    assert rule_matches(rule_record(match_from=["icna relief canada <receipts"]), email)
    assert rule_matches(rule_record(match_from=["STRIPE.COM"]), email)
    assert not rule_matches(rule_record(match_from=["@icnareliefcanada.ca"]), email)


def test_subject_must_also_match_when_given():
    email = email_record(subject="Your Receipt from Muslim Hands")
    assert rule_matches(rule_record(match_from=["sickkids"], match_subject=["receipt from"]), email)
    assert not rule_matches(rule_record(match_from=["sickkids"], match_subject=["water well"]), email)


def test_any_of_lists():
    email = email_record(sender="x <support@themasjidapp.net>", subject="Your Hamilton Downtown Mosque receipt")
    rule = rule_record(match_from=["donorchoice.net", "themasjidapp.net"], match_subject=["nope", "Hamilton Downtown"])
    assert rule_matches(rule, email)


def test_match_rule_first_wins_and_reports_others():
    email = email_record()
    a = rule_record(name="A", match_from=["sickkids"])
    b = rule_record(name="B", match_from=["receipts@"])
    c = rule_record(name="C", match_from=["nomatch"])
    rule, others = match_rule(email, [a, b, c])
    assert rule is a
    assert others == [b]
    assert match_rule(email, [c]) == (None, [])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_matcher.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation** — `donation_tracker/matcher.py`

```python
"""Match emails to rules on the From and Subject headers."""
from __future__ import annotations

from .models import Email, Rule


def rule_matches(rule: Rule, email: Email) -> bool:
    sender = email.sender.lower()
    if not any(fragment.lower() in sender for fragment in rule.match_from):
        return False
    if rule.match_subject:
        subject = email.subject.lower()
        return any(fragment.lower() in subject for fragment in rule.match_subject)
    return True


def match_rule(email: Email, rules: list[Rule]) -> tuple[Rule | None, list[Rule]]:
    matches = [rule for rule in rules if rule_matches(rule, email)]
    if not matches:
        return None, []
    return matches[0], matches[1:]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_matcher.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/matcher.py tests/test_matcher.py
git commit -m "feat: first-match-wins rule matcher"
```

---

### Task 5: Field extractor

**Files:**
- Create: `donation_tracker/extractor.py`
- Test: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `models.Email`, `models.Rule`, `models.Donation`, `OK`, `NEEDS_REVIEW`.
- Produces: `extract_donation(email: Email, rule: Rule, text: str, textless_pdf: bool) -> Donation`; `tax_year_for(email: Email, rule: Rule) -> int`; constant `IRS_NOTE = "IRS: written acknowledgment required for $250+"`.

- [ ] **Step 1: Write the failing tests** — `tests/test_extractor.py`

```python
import re
from datetime import datetime, timezone
from decimal import Decimal

import pytest

from donation_tracker.extractor import IRS_NOTE, extract_donation, tax_year_for
from donation_tracker.models import NEEDS_REVIEW, OK, Attachment
from tests.helpers import email_record, rule_record


@pytest.mark.parametrize(
    "text, expected",
    [
        ("Donation Amount: $50.00", "50.00"),
        ("Total CAD $1,200.00 thank you", "1200.00"),
        ("Your gift of $25 helps", "25"),
        ("Eligible amount for tax purposes: $153.75", "153.75"),
        ("We received $60.00 USD donation", "60.00"),
    ],
)
def test_default_amount_patterns(text, expected):
    d = extract_donation(email_record(), rule_record(), text, False)
    assert d.amount == Decimal(expected)
    assert d.status == OK


def test_custom_amount_regex_overrides_default():
    rule = rule_record(extract={"amount": re.compile(r"Paid\s+([\d.]+)", re.I)})
    d = extract_donation(email_record(), rule, "Total $999.00 Paid 12.34", False)
    assert d.amount == Decimal("12.34")


def test_missing_amount_needs_review_and_mentions_textless_pdf():
    d = extract_donation(email_record(), rule_record(), "Thanks!", True)
    assert d.amount is None
    assert d.status == NEEDS_REVIEW
    assert d.notes == ["Amount not found; PDF has no text layer"]


def test_receipt_number_default_and_custom():
    d = extract_donation(email_record(), rule_record(), "Receipt Number: RD-US1939527 Amount: $60.00", False)
    assert d.receipt_no == "RD-US1939527"
    rule = rule_record(extract={"receipt_no": re.compile(r"Ref (\w+)")})
    assert extract_donation(email_record(), rule, "Ref AB12 Amount: $1.00", False).receipt_no == "AB12"


def test_charity_id_order_regex_then_fixed_then_default():
    text = "Amount: $5.00 Charitable registration 123456789 RR 0001"
    assert extract_donation(email_record(), rule_record(), text, False).charity_id == "123456789 RR 0001"
    fixed = rule_record(charity_id="999999999RR0001")
    assert extract_donation(email_record(), fixed, text, False).charity_id == "999999999RR0001"
    us = extract_donation(email_record(), rule_record(), "Amount: $5.00 EIN 12-3456789", False)
    assert us.charity_id == "12-3456789"


def test_irs_flag_for_us_250_plus_without_receipt():
    us_rule = rule_record(country="US", currency="USD")
    flagged = extract_donation(email_record(), us_rule, "Donation Amount: $250.00", False)
    assert IRS_NOTE in flagged.notes
    with_no = extract_donation(email_record(), us_rule, "Donation Amount: $300.00 Receipt #: A123", False)
    assert IRS_NOTE not in with_no.notes
    pdf = email_record(attachments=[Attachment("r.pdf", "application/pdf", b"")])
    assert IRS_NOTE not in extract_donation(pdf, us_rule, "Donation Amount: $300.00", False).notes
    small = extract_donation(email_record(), us_rule, "Donation Amount: $249.99", False)
    assert IRS_NOTE not in small.notes
    ca = extract_donation(email_record(), rule_record(), "Donation Amount: $500.00", False)
    assert IRS_NOTE not in ca.notes


def test_tax_year_for():
    email = email_record(date=datetime(2026, 2, 24, tzinfo=timezone.utc))
    assert tax_year_for(email, rule_record()) == 2026
    assert tax_year_for(email, rule_record(tax_year="previous")) == 2025
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_extractor.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation** — `donation_tracker/extractor.py`

```python
"""Pull amount, receipt number and charity ID out of a matched email."""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

from .models import NEEDS_REVIEW, OK, Donation, Email, Rule

IRS_NOTE = "IRS: written acknowledgment required for $250+"

DEFAULT_AMOUNT = [
    re.compile(
        r"(?:amount|total|gift|donation|donated|eligible|contribution)[^$\d\n]{0,40}\$\s*([\d,]+(?:\.\d{2})?)",
        re.IGNORECASE,
    ),
    re.compile(r"\$\s*([\d,]+(?:\.\d{2})?)\s*(?:CAD|USD)?\s*(?:donation|gift|contribution)", re.IGNORECASE),
]
DEFAULT_RECEIPT_NO = [
    re.compile(r"receipt\s*(?:no\.?|number|#)\s*[:#]?\s*([A-Z0-9][A-Z0-9-]{2,})", re.IGNORECASE),
]
DEFAULT_CHARITY_ID = [
    re.compile(r"\b(\d{9}\s?RR\s?\d{4})\b", re.IGNORECASE),
    re.compile(r"\b(\d{2}-\d{7})\b"),
]


def _first_group(patterns: list[re.Pattern], text: str) -> str | None:
    for pattern in patterns:
        match = pattern.search(text)
        if match and match.group(1):
            return match.group(1).strip()
    return None


def _parse_amount(value: str) -> Decimal | None:
    try:
        return Decimal(value.replace(",", ""))
    except InvalidOperation:
        return None


def _find_amount(patterns: list[re.Pattern], text: str) -> Decimal | None:
    for pattern in patterns:
        for match in pattern.finditer(text):
            amount = _parse_amount(match.group(1))
            if amount is not None:
                return amount
    return None


def tax_year_for(email: Email, rule: Rule) -> int:
    return email.date.year - (1 if rule.tax_year == "previous" else 0)


def extract_donation(email: Email, rule: Rule, text: str, textless_pdf: bool) -> Donation:
    amount_patterns = [rule.extract["amount"]] if "amount" in rule.extract else DEFAULT_AMOUNT
    receipt_patterns = [rule.extract["receipt_no"]] if "receipt_no" in rule.extract else DEFAULT_RECEIPT_NO
    amount = _find_amount(amount_patterns, text)
    receipt_no = _first_group(receipt_patterns, text)
    charity_id = _first_group([rule.extract["charity_id"]], text) if "charity_id" in rule.extract else None
    charity_id = charity_id or rule.charity_id or _first_group(DEFAULT_CHARITY_ID, text)

    notes = []
    status = OK
    if amount is None:
        status = NEEDS_REVIEW
        notes.append("Amount not found" + ("; PDF has no text layer" if textless_pdf else ""))
    has_pdf = any(attachment.is_pdf for attachment in email.attachments)
    if rule.country == "US" and amount is not None and amount >= 250 and not receipt_no and not has_pdf:
        notes.append(IRS_NOTE)
    return Donation(email, rule, amount, receipt_no, charity_id, status, notes)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_extractor.py -v`
Expected: all passed (11)

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/extractor.py tests/test_extractor.py
git commit -m "feat: extract amount, receipt number and charity id"
```

---

### Task 6: Candidate detector and starter rules

**Files:**
- Create: `donation_tracker/detector.py`, `detector.yaml`
- Test: `tests/test_detector.py`

**Interfaces:**
- Consumes: `config.load_yaml`, `config.ConfigError`, `models.Email`.
- Produces: `DetectorConfig(candidate_subject, candidate_body, exclude_subject: list[re.Pattern], exclude_sender: list[str])`; `load_detector(path: Path) -> DetectorConfig`; `is_candidate(email: Email, cfg: DetectorConfig) -> bool`; `starter_rule(email: Email) -> dict`; `suggested_rules_yaml(emails: list[Email]) -> str`.

- [ ] **Step 1: Write the failing tests** — `tests/test_detector.py`

```python
from pathlib import Path

import pytest
import yaml

from donation_tracker.config import ConfigError
from donation_tracker.detector import is_candidate, load_detector, starter_rule, suggested_rules_yaml
from tests.helpers import email_record

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture
def cfg():
    return load_detector(ROOT / "detector.yaml")


def test_receipt_like_email_is_candidate(cfg):
    assert is_candidate(email_record(subject="Your donation receipt"), cfg)
    assert is_candidate(email_record(subject="Hello", body_text="This is your official receipt for income tax"), cfg)


def test_solicitation_and_commercial_excluded(cfg):
    assert not is_candidate(email_record(subject="This Giving Tuesday, donate to Gaza"), cfg)
    assert not is_candidate(email_record(sender="Airbnb <automated@airbnb.com>", subject="Your receipt"), cfg)


def test_plain_email_not_candidate(cfg):
    assert not is_candidate(email_record(subject="Lunch?", body_text="see you"), cfg)


def test_starter_rule_guesses_country():
    ca = starter_rule(email_record(sender="Food Bank <noreply@foodbank.ca>", subject="Your receipt"))
    assert ca == {
        "name": "Food Bank",
        "country": "CA",
        "currency": "CAD",
        "role": "receipt",
        "match": {"from": ["@foodbank.ca"], "subject": ["Your receipt"]},
    }
    us = starter_rule(email_record(sender="noreply@charity.org", body_text="tax-deductible"))
    assert us["country"] == "US" and us["currency"] == "USD" and us["name"] == "charity.org"
    cra = starter_rule(email_record(sender="x@charity.org", body_text="Registered charity 123456789RR0001"))
    assert cra["country"] == "CA"


def test_suggested_rules_yaml_dedupes_by_domain_and_is_valid_yaml():
    emails = [
        email_record(sender="A <a@one.ca>", subject="Receipt 1"),
        email_record(sender="A <b@one.ca>", subject="Receipt 2"),
        email_record(sender="B <x@two.org>", subject="Thanks"),
    ]
    text = suggested_rules_yaml(emails)
    assert text.count("# REVIEW") == 2
    assert "seen 2 email(s)" in text
    parsed = yaml.safe_load(text)
    assert [r["match"]["from"] for r in parsed] == [["@one.ca"], ["@two.org"]]


def test_bad_detector_regex(tmp_path):
    path = tmp_path / "d.yaml"
    path.write_text("candidate_subject: ['(']\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="candidate_subject"):
        load_detector(path)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_detector.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation**

`detector.yaml` (repo root; lists carried over from attempt2's `parse_donations_by_year.py`):
```yaml
# Decides which UNMATCHED emails appear on the "Needs a rule" sheet.
# All entries are case-insensitive regular expressions except exclude_sender (plain substrings).
candidate_subject:
  - donat
  - tax receipt
  - official receipt
  - charitable
  - contribution receipt
  - 'acknowledg(e)?ment'
  - 'thank you for your (gift|donation|contribution|support)'
  - zakat
  - sadaqah
  - 'night donation'
candidate_body:
  - 'tax.?deductible'
  - '501\(c\)\(3\)'
  - charitable contribution
  - donation receipt
  - tax receipt
  - official receipt
  - no goods or services
  - cra registration
  - registered charity
  - charity reg
  - donation confirmation
exclude_subject:
  - donate to
  - why donate
  - gifts doubled
  - 'last day.*gift'
  - giving tuesday
  - swoop in
  - fundrais
  - matching gift
  - provide hope
  - 'give.*early'
  - "don.*miss out"
  - automate your
  - 'double.*gift'
  - 'gift.*double'
  - last chance
exclude_sender:
  - apple.com
  - shopify
  - etsy.com
  - uniqlo
  - priceline
  - airbnb
  - amazon
  - uber
  - doordash
  - bestbuy
  - walmart
  - costco
  - ebay
  - paypal
  - netflix
  - spotify
  - google.com
  - microsoft
  - nintendo
  - uhaul
  - footlocker
  - autotrader
  - lindt
  - simonsnewsletter
```

`donation_tracker/detector.py`:
```python
"""Spot unmatched emails that look like donations, and draft starter rules for them."""
from __future__ import annotations

import re
from dataclasses import dataclass
from email.utils import parseaddr
from pathlib import Path

import yaml

from .config import ConfigError, load_yaml
from .models import Email

_CA_HINTS = re.compile(r"\bCRA\b|registered charity|charitable registration|\d{9}\s?RR\s?\d{4}", re.IGNORECASE)


@dataclass
class DetectorConfig:
    candidate_subject: list[re.Pattern]
    candidate_body: list[re.Pattern]
    exclude_subject: list[re.Pattern]
    exclude_sender: list[str]


def _patterns(path: Path, data: dict, key: str) -> list[re.Pattern]:
    compiled = []
    for pattern in data.get(key) or []:
        try:
            compiled.append(re.compile(str(pattern), re.IGNORECASE))
        except re.error as exc:
            raise ConfigError(f"{path.name}: {key} entry {pattern!r} is not a valid regex: {exc}") from exc
    return compiled


def load_detector(path: Path) -> DetectorConfig:
    data = load_yaml(path) or {}
    return DetectorConfig(
        candidate_subject=_patterns(path, data, "candidate_subject"),
        candidate_body=_patterns(path, data, "candidate_body"),
        exclude_subject=_patterns(path, data, "exclude_subject"),
        exclude_sender=[str(s).lower() for s in data.get("exclude_sender") or []],
    )


def is_candidate(email: Email, cfg: DetectorConfig) -> bool:
    sender = email.sender.lower()
    if any(fragment in sender for fragment in cfg.exclude_sender):
        return False
    if any(p.search(email.subject) for p in cfg.exclude_subject):
        return False
    if any(p.search(email.subject) for p in cfg.candidate_subject):
        return True
    body = email.body_text[:5000]
    return any(p.search(body) for p in cfg.candidate_body)


def _domain(email: Email) -> str:
    address = parseaddr(email.sender)[1] or email.sender
    return address.rpartition("@")[2].lower() or address.lower()


def starter_rule(email: Email) -> dict:
    display = parseaddr(email.sender)[0].strip()
    domain = _domain(email)
    is_ca = domain.endswith(".ca") or bool(_CA_HINTS.search(email.body_text))
    return {
        "name": display or domain,
        "country": "CA" if is_ca else "US",
        "currency": "CAD" if is_ca else "USD",
        "role": "receipt",
        "match": {"from": [f"@{domain}"], "subject": [email.subject]},
    }


def suggested_rules_yaml(emails: list[Email]) -> str:
    groups: dict[str, list[Email]] = {}
    for email in emails:
        groups.setdefault(_domain(email), []).append(email)
    blocks = []
    for group in groups.values():
        first = min(group, key=lambda e: e.date)
        header = (
            f"# REVIEW: seen {len(group)} email(s); first {first.date:%Y-%m-%d} "
            f"\"{first.subject}\" — check country/currency/role, loosen subject if needed\n"
        )
        body = yaml.safe_dump([starter_rule(first)], sort_keys=False, allow_unicode=True)
        blocks.append(header + body)
    return "\n".join(blocks)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_detector.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/detector.py detector.yaml tests/test_detector.py
git commit -m "feat: needs-a-rule detector and starter rule suggestions"
```

---

### Task 7: Gmail links and evidence files

**Files:**
- Create: `donation_tracker/links.py`, `donation_tracker/evidence.py`
- Test: `tests/test_evidence.py`

**Interfaces:**
- Consumes: `models.Donation`, `models.Email`.
- Produces: `links.gmail_url(account: str, message_id: str) -> str` ("" for synthesized `<no-id-…>`); `evidence.safe_name(s: str, maxlen: int = 40) -> str`; `evidence.org_slug(name: str) -> str`; `evidence.unique_path(path: Path) -> Path`; `evidence.email_to_html(email: Email) -> str`; `evidence.save_evidence(donation: Donation, evidence_root: Path) -> None` (fills `email_eml_path`, `email_html_path`, `receipt_paths`).

- [ ] **Step 1: Write the failing tests** — `tests/test_evidence.py`

```python
from decimal import Decimal

from donation_tracker.evidence import email_to_html, org_slug, safe_name, save_evidence, unique_path
from donation_tracker.links import gmail_url
from donation_tracker.models import OK, Attachment, Donation
from tests.helpers import email_record, rule_record


def test_gmail_url():
    assert gmail_url("abbasirumael@gmail.com", "<abc+1@mail.x.com>") == (
        "https://mail.google.com/mail/u/abbasirumael@gmail.com/#search/rfc822msgid:abc%2B1%40mail.x.com"
    )
    assert gmail_url("a@gmail.com", "<no-id-123>") == ""


def test_safe_name_and_slug():
    assert safe_name('a<b>:c"d/e\\f|g?h*i') == "a_b__c_d_e_f_g_h_i"
    assert org_slug("Muslim Hands Canada") == "Muslim_Hands_Canada"


def test_unique_path(tmp_path):
    (tmp_path / "x").mkdir()
    (tmp_path / "x_2").mkdir()
    assert unique_path(tmp_path / "x") == tmp_path / "x_3"
    (tmp_path / "r.pdf").write_bytes(b"")
    assert unique_path(tmp_path / "r.pdf") == tmp_path / "r_2.pdf"


def test_email_to_html_escapes_plain_text_and_has_header():
    html = email_to_html(email_record(subject="Tax <receipt>", body_text="1 < 2"))
    assert "Tax &lt;receipt&gt;" in html
    assert "1 &lt; 2" in html
    assert "abbasirumael@gmail.com" in html
    assert "<meta charset=\"utf-8\">" in html


def test_email_to_html_uses_html_body_when_present():
    html = email_to_html(email_record(body_html="<p>Gift <b>$5</b></p>", body_text="Gift $5"))
    assert "<p>Gift <b>$5</b></p>" in html


def test_save_evidence_writes_files_and_sets_paths(tmp_path):
    atts = [
        Attachment("RD-US1939527-$60 -2026-04-21.pdf", "application/pdf", b"%PDF-1"),
        Attachment("logo.png", "image/png", b"PNG"),
    ]
    donation = Donation(email_record(attachments=atts, raw=b"RAW"), rule_record(), Decimal("60"), None, None, OK)
    save_evidence(donation, tmp_path / "evidence")
    folder = tmp_path / "evidence" / "2025-03-14_SickKids_Foundation"
    assert donation.email_eml_path == folder / "email.eml"
    assert donation.email_eml_path.read_bytes() == b"RAW"
    assert donation.email_html_path == folder / "email.html"
    assert donation.receipt_paths == [folder / "RD-US1939527-$60 -2026-04-21.pdf"]
    assert (folder / "logo.png").read_bytes() == b"PNG"

    second = Donation(email_record(raw=b"R2"), rule_record(), Decimal("1"), None, None, OK)
    save_evidence(second, tmp_path / "evidence")
    assert second.email_eml_path.parent.name == "2025-03-14_SickKids_Foundation_2"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_evidence.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation**

`donation_tracker/links.py`:
```python
"""Links back to the original message in Gmail."""
from __future__ import annotations

from urllib.parse import quote


def gmail_url(account: str, message_id: str) -> str:
    if message_id.startswith("<no-id-"):
        return ""
    return f"https://mail.google.com/mail/u/{account}/#search/rfc822msgid:{quote(message_id.strip('<> '), safe='')}"
```

`donation_tracker/evidence.py`:
```python
"""Save each donation email (.eml, readable .html) and its attachments.

safe_name and the HTML wrapper are adapted from archived_attempts/attempt2.
"""
from __future__ import annotations

import html as html_mod
import re
from pathlib import Path

from .models import Donation, Email


def safe_name(s: str, maxlen: int = 40) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    return s[:maxlen].strip("_. ")


def org_slug(name: str) -> str:
    return safe_name(re.sub(r"\s+", "_", name.strip()), 40)


def _safe_filename(filename: str) -> str:
    path = Path(filename)
    suffix = safe_name(path.suffix, 10)
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_evidence.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/links.py donation_tracker/evidence.py tests/test_evidence.py
git commit -m "feat: gmail links and evidence folders"
```

---

### Task 8: HTML → PDF with headless Edge

**Files:**
- Create: `donation_tracker/pdf_render.py`
- Test: `tests/test_pdf_render.py`

**Interfaces:**
- Produces: `find_edge() -> str | None`; `EdgeRenderer(exe: str | None = None)` with `.render(html_path: Path, pdf_path: Path) -> bool`. Any object with this `render` signature is a valid renderer (tests use a fake).

- [ ] **Step 1: Write the failing tests** — `tests/test_pdf_render.py`

```python
import pytest

from donation_tracker.pdf_render import EdgeRenderer, find_edge


def test_renderer_without_edge_returns_false(tmp_path):
    renderer = EdgeRenderer(exe="")
    html = tmp_path / "a.html"
    html.write_text("<p>x</p>", encoding="utf-8")
    assert renderer.render(html, tmp_path / "a.pdf") is False


@pytest.mark.skipif(find_edge() is None, reason="Microsoft Edge not installed")
def test_edge_renders_real_pdf(tmp_path):
    html = tmp_path / "email.html"
    html.write_text("<html><body><h1>Receipt $50.00</h1></body></html>", encoding="utf-8")
    pdf = tmp_path / "out" / "email.pdf"
    assert EdgeRenderer().render(html, pdf) is True
    assert pdf.read_bytes().startswith(b"%PDF")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_pdf_render.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation** — `donation_tracker/pdf_render.py`

```python
"""Render saved email HTML to PDF using headless Microsoft Edge (preinstalled on Windows 11)."""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


def find_edge() -> str | None:
    for candidate in EDGE_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return shutil.which("msedge")


class EdgeRenderer:
    def __init__(self, exe: str | None = None):
        self.exe = find_edge() if exe is None else exe
        self._profile = tempfile.TemporaryDirectory(prefix="donations-edge-")

    def render(self, html_path: Path, pdf_path: Path) -> bool:
        if not self.exe:
            return False
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        command = [
            self.exe,
            "--headless",
            "--disable-gpu",
            "--no-first-run",
            f"--user-data-dir={self._profile.name}",
            "--no-pdf-header-footer",
            "--print-to-pdf-no-header",
            f"--print-to-pdf={pdf_path.resolve()}",
            html_path.resolve().as_uri(),
        ]
        try:
            subprocess.run(command, capture_output=True, timeout=90, check=False)
        except (OSError, subprocess.TimeoutExpired):
            return False
        return pdf_path.exists() and pdf_path.stat().st_size > 0
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_pdf_render.py -v`
Expected: 2 passed (second one produces a real PDF on this machine). If the Edge test fails, run the command list by hand in PowerShell to see Edge's stderr; `--headless=new` is the alternative flag for newer Edge builds.

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/pdf_render.py tests/test_pdf_render.py
git commit -m "feat: headless Edge HTML to PDF renderer"
```

---

### Task 9: Excel report writer

**Files:**
- Create: `donation_tracker/report.py`
- Test: `tests/test_report.py`

**Interfaces:**
- Consumes: `models.Donation`, `models.Email`, `links.gmail_url`.
- Produces:
  - `ReportRow(date: date, organization, country, currency, amount: Decimal | None, role, receipt_no, charity_id, account, status, gmail_url, email_file: Path | None, receipt_file: Path | None, notes: str)`
  - `Summary(totals: dict[tuple[str, str], Decimal], by_org: dict[tuple[str, str], Decimal], needs_review: int, needs_rule: int, skipped: int)`
  - `row_from_donation(d: Donation, email_file: Path | None = None, receipt_files: list[Path] | None = None, extra_notes: list[str] | None = None) -> ReportRow` (defaults: `d.email_html_path`, `d.receipt_paths`)
  - `summarize(donations: list[Donation], needs_rule: int, skipped: int) -> Summary`
  - `write_workbook(path: Path, rows: list[ReportRow], summary: Summary, needs_rule: list[Email] | None) -> None` (Needs-a-rule sheet omitted when `None`)

- [ ] **Step 1: Write the failing tests** — `tests/test_report.py`

```python
from decimal import Decimal

from openpyxl import load_workbook

from donation_tracker.models import NEEDS_REVIEW, OK, Donation
from donation_tracker.report import row_from_donation, summarize, write_workbook
from tests.helpers import email_record, rule_record


def donation(amount, *, role="receipt", status=OK, country="CA", currency="CAD", name="SickKids Foundation"):
    rule = rule_record(role=role, country=country, currency=currency, name=name)
    d = Donation(email_record(), rule, Decimal(amount) if amount else None, "R1", None, status, [])
    d.gmail_url = "https://mail.google.com/x"
    return d


def test_summarize_counts_only_ok_receipts():
    ds = [
        donation("100"),
        donation("50", role="confirmation"),
        donation(None, status=NEEDS_REVIEW),
        donation("20", country="US", currency="USD", name="Red Cross"),
        donation("5"),
    ]
    s = summarize(ds, needs_rule=3, skipped=1)
    assert s.totals == {("CA", "CAD"): Decimal("105"), ("US", "USD"): Decimal("20")}
    assert s.by_org == {("SickKids Foundation", "CAD"): Decimal("105"), ("Red Cross", "USD"): Decimal("20")}
    assert (s.needs_review, s.needs_rule, s.skipped) == (1, 3, 1)


def test_write_workbook_sheets_rows_and_relative_links(tmp_path):
    out = tmp_path / "2025"
    folder = out / "evidence" / "2025-03-14_SickKids_Foundation"
    folder.mkdir(parents=True)
    d = donation("100")
    d.email_html_path = folder / "email.html"
    d.receipt_paths = [folder / "a.pdf", folder / "b.pdf"]
    review = donation(None, status=NEEDS_REVIEW)
    review.notes = ["Amount not found"]
    rows = [row_from_donation(d), row_from_donation(review)]
    candidate = email_record(sender="New <x@new.org>", subject="Thanks for your gift")
    write_workbook(out / "donations_2025.xlsx", rows, summarize([d, review], 1, 0), [candidate])

    wb = load_workbook(out / "donations_2025.xlsx")
    assert wb.sheetnames == ["Summary", "Donations", "Needs a rule"]
    ws = wb["Donations"]
    header = [c.value for c in ws[1]]
    assert header == ["Date", "Organization", "Country", "Currency", "Amount", "Role", "Receipt #",
                      "Charity ID", "Account", "Status", "Gmail", "Email file", "Receipt file", "Notes"]
    row = {h: c for h, c in zip(header, ws[2])}
    assert row["Amount"].value == 100
    assert row["Gmail"].hyperlink.target == "https://mail.google.com/x"
    assert row["Email file"].hyperlink.target == "evidence/2025-03-14_SickKids_Foundation/email.html"
    assert row["Receipt file"].hyperlink.target == "evidence/2025-03-14_SickKids_Foundation/a.pdf"
    assert "b.pdf" in row["Notes"].value
    review_row = {h: c for h, c in zip(header, ws[3])}
    assert review_row["Status"].value == "NEEDS REVIEW"
    assert review_row["Status"].fill.fgColor.rgb.endswith("FFFF00")
    summary_values = [c.value for r in wb["Summary"].iter_rows() for c in r if c.value is not None]
    assert "CA / CAD" in summary_values and 100 in summary_values
    rules_sheet = wb["Needs a rule"]
    assert rules_sheet["D2"].value == "Thanks for your gift"


def test_workbook_without_needs_rule_sheet(tmp_path):
    write_workbook(tmp_path / "s.xlsx", [], summarize([], 0, 0), None)
    assert load_workbook(tmp_path / "s.xlsx").sheetnames == ["Summary", "Donations"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_report.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation** — `donation_tracker/report.py`

```python
"""Build the donations Excel workbook."""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from .links import gmail_url
from .models import NEEDS_REVIEW, Donation, Email

HEADERS = ["Date", "Organization", "Country", "Currency", "Amount", "Role", "Receipt #",
           "Charity ID", "Account", "Status", "Gmail", "Email file", "Receipt file", "Notes"]
YELLOW = PatternFill("solid", fgColor="FFFF00")
BOLD = Font(bold=True)
MONEY = "#,##0.00"


@dataclass
class ReportRow:
    date: date
    organization: str
    country: str
    currency: str
    amount: Decimal | None
    role: str
    receipt_no: str | None
    charity_id: str | None
    account: str
    status: str
    gmail_url: str
    email_file: Path | None
    receipt_file: Path | None
    notes: str


@dataclass
class Summary:
    totals: dict[tuple[str, str], Decimal]
    by_org: dict[tuple[str, str], Decimal]
    needs_review: int
    needs_rule: int
    skipped: int


def row_from_donation(
    d: Donation,
    email_file: Path | None = None,
    receipt_files: list[Path] | None = None,
    extra_notes: list[str] | None = None,
) -> ReportRow:
    receipts = d.receipt_paths if receipt_files is None else receipt_files
    notes = list(d.notes) + list(extra_notes or [])
    if len(receipts) > 1:
        notes.append("Other receipt files: " + ", ".join(p.name for p in receipts[1:]))
    return ReportRow(
        date=d.email.date.date(),
        organization=d.rule.name,
        country=d.rule.country,
        currency=d.rule.currency,
        amount=d.amount,
        role=d.rule.role,
        receipt_no=d.receipt_no,
        charity_id=d.charity_id,
        account=d.email.account,
        status=d.status,
        gmail_url=d.gmail_url,
        email_file=email_file or d.email_html_path,
        receipt_file=receipts[0] if receipts else None,
        notes="; ".join(notes),
    )


def summarize(donations: list[Donation], needs_rule: int, skipped: int) -> Summary:
    totals: dict[tuple[str, str], Decimal] = {}
    by_org: dict[tuple[str, str], Decimal] = {}
    for d in donations:
        if not d.is_countable:
            continue
        key = (d.rule.country, d.rule.currency)
        totals[key] = totals.get(key, Decimal(0)) + d.amount
        org_key = (d.rule.name, d.rule.currency)
        by_org[org_key] = by_org.get(org_key, Decimal(0)) + d.amount
    needs_review = sum(1 for d in donations if d.status == NEEDS_REVIEW)
    return Summary(totals, by_org, needs_review, needs_rule, skipped)


def _link(cell, target: str, label: str) -> None:
    cell.value = label
    cell.hyperlink = target
    cell.style = "Hyperlink"


def _relative(target: Path, workbook_path: Path) -> str:
    return os.path.relpath(target, workbook_path.parent).replace("\\", "/")


def _write_summary(ws, summary: Summary) -> None:
    ws.append(["Totals (receipts with status OK)"])
    ws["A1"].font = BOLD
    ws.append(["Country / Currency", "Total"])
    for (country, currency), total in sorted(summary.totals.items()):
        ws.append([f"{country} / {currency}", total])
    ws.append([])
    ws.append(["By organization"])
    ws.cell(ws.max_row, 1).font = BOLD
    ws.append(["Organization", "Currency", "Total"])
    for (org, currency), total in sorted(summary.by_org.items()):
        ws.append([org, currency, total])
    ws.append([])
    ws.append(["Checks"])
    ws.cell(ws.max_row, 1).font = BOLD
    ws.append(["Rows needing review", summary.needs_review])
    ws.append(["Emails needing a rule", summary.needs_rule])
    ws.append(["Malformed messages skipped", summary.skipped])
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, Decimal):
                cell.number_format = MONEY
    ws.column_dimensions["A"].width = 40


def _write_donations(ws, rows: list[ReportRow], workbook_path: Path) -> None:
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = BOLD
    for r in rows:
        ws.append([r.date, r.organization, r.country, r.currency, r.amount, r.role, r.receipt_no,
                   r.charity_id, r.account, r.status, None, None, None, r.notes])
        excel_row = ws[ws.max_row]
        excel_row[0].number_format = "yyyy-mm-dd"
        excel_row[4].number_format = MONEY
        if r.gmail_url:
            _link(excel_row[10], r.gmail_url, "Open in Gmail")
        if r.email_file:
            _link(excel_row[11], _relative(r.email_file, workbook_path), r.email_file.name)
        if r.receipt_file:
            _link(excel_row[12], _relative(r.receipt_file, workbook_path), r.receipt_file.name)
        if r.status == NEEDS_REVIEW:
            for cell in excel_row:
                cell.fill = YELLOW
    widths = [12, 30, 8, 9, 12, 13, 18, 20, 26, 14, 14, 30, 30, 50]
    for index, width in enumerate(widths):
        ws.column_dimensions[chr(ord("A") + index)].width = width
    ws.freeze_panes = "A2"


def _write_needs_rule(ws, emails: list[Email]) -> None:
    ws.append(["Date", "Account", "From", "Subject", "Gmail"])
    for cell in ws[1]:
        cell.font = BOLD
    for email in sorted(emails, key=lambda e: e.date):
        ws.append([email.date.date(), email.account, email.sender, email.subject, None])
        url = gmail_url(email.account, email.message_id)
        if url:
            _link(ws.cell(ws.max_row, 5), url, "Open in Gmail")
    for column, width in zip("ABCDE", [12, 26, 45, 60, 14]):
        ws.column_dimensions[column].width = width


def write_workbook(path: Path, rows: list[ReportRow], summary: Summary, needs_rule: list[Email] | None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    _write_summary(wb.active, summary)
    wb.active.title = "Summary"
    _write_donations(wb.create_sheet("Donations"), rows, path)
    if needs_rule is not None:
        _write_needs_rule(wb.create_sheet("Needs a rule"), needs_rule)
    wb.save(path)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_report.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/report.py tests/test_report.py
git commit -m "feat: Excel report with summary, donations and needs-a-rule sheets"
```

---

### Task 10: Accountant pack

**Files:**
- Create: `donation_tracker/accountant.py`
- Test: `tests/test_accountant.py`

**Interfaces:**
- Consumes: `models.Donation`, `evidence.org_slug`, `evidence.unique_path`, `report.row_from_donation`, `report.write_workbook`, `report.Summary`, a renderer with `.render(html_path, pdf_path) -> bool`.
- Produces: `build_pack(donations: list[Donation], pack_dir: Path, renderer, year: int, summary: Summary, log: list[str]) -> list[Path]` (returns the files written, excluding the workbook).

- [ ] **Step 1: Write the failing tests** — `tests/test_accountant.py`

```python
from decimal import Decimal

from openpyxl import load_workbook

from donation_tracker.accountant import build_pack
from donation_tracker.models import NEEDS_REVIEW, OK, Donation
from donation_tracker.report import summarize
from tests.helpers import email_record, rule_record


class FakeRenderer:
    def __init__(self, ok=True):
        self.ok = ok

    def render(self, html_path, pdf_path):
        if self.ok:
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            pdf_path.write_bytes(b"%PDF-fake")
        return self.ok


def saved_donation(tmp_path, name, amount, *, role="receipt", status=OK, country="CA", currency="CAD", receipts=1):
    folder = tmp_path / "evidence" / name
    folder.mkdir(parents=True)
    html = folder / "email.html"
    html.write_text("<p>x</p>", encoding="utf-8")
    d = Donation(email_record(), rule_record(name=name, role=role, country=country, currency=currency),
                 Decimal(amount) if amount else None, None, None, status, [])
    d.email_html_path = html
    for i in range(receipts):
        pdf = folder / f"r{i}.pdf"
        pdf.write_bytes(b"%PDF")
        d.receipt_paths.append(pdf)
    return d


def test_pack_contains_only_ok_receipts_split_by_country(tmp_path):
    ds = [
        saved_donation(tmp_path, "Muslim Hands", "1200", receipts=2),
        saved_donation(tmp_path, "Red Cross", "250", country="US", currency="USD", receipts=0),
        saved_donation(tmp_path, "MyTenNights nightly", "3.75", role="confirmation"),
        saved_donation(tmp_path, "Broken", None, status=NEEDS_REVIEW),
    ]
    pack = tmp_path / "for_accountant"
    log = []
    files = build_pack(ds, pack, FakeRenderer(), 2025, summarize(ds, 0, 0), log)
    names = sorted(p.relative_to(pack).as_posix() for p in files)
    assert names == [
        "CA/2025-03-14_Muslim_Hands_CAD_1200.00_email.pdf",
        "CA/2025-03-14_Muslim_Hands_CAD_1200.00_receipt.pdf",
        "CA/2025-03-14_Muslim_Hands_CAD_1200.00_receipt_2.pdf",
        "US/2025-03-14_Red_Cross_USD_250.00_email.pdf",
    ]
    wb = load_workbook(pack / "donations_2025_summary.xlsx")
    assert wb.sheetnames == ["Summary", "Donations"]
    ws = wb["Donations"]
    assert ws.max_row == 3
    assert ws["L2"].hyperlink.target == "CA/2025-03-14_Muslim_Hands_CAD_1200.00_email.pdf"
    assert ws["M2"].hyperlink.target == "CA/2025-03-14_Muslim_Hands_CAD_1200.00_receipt.pdf"


def test_render_failure_falls_back_to_html(tmp_path):
    d = saved_donation(tmp_path, "Muslim Hands", "10", receipts=0)
    log = []
    files = build_pack([d], tmp_path / "pack", FakeRenderer(ok=False), 2025, summarize([d], 0, 0), log)
    assert [p.name for p in files] == ["2025-03-14_Muslim_Hands_CAD_10.00_email.html"]
    assert any("PDF render failed" in line for line in log)
    ws = load_workbook(tmp_path / "pack" / "donations_2025_summary.xlsx")["Donations"]
    assert "PDF render failed" in ws["N2"].value
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_accountant.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation** — `donation_tracker/accountant.py`

```python
"""Build the self-contained for_accountant/ folder."""
from __future__ import annotations

import shutil
from pathlib import Path

from .evidence import org_slug, unique_path
from .models import Donation
from .report import Summary, row_from_donation, write_workbook

RENDER_FAILED = "PDF render failed; email included as HTML"


def build_pack(donations: list[Donation], pack_dir: Path, renderer, year: int, summary: Summary,
               log: list[str]) -> list[Path]:
    written: list[Path] = []
    rows = []
    for d in donations:
        if not d.is_countable:
            continue
        folder = pack_dir / d.rule.country
        folder.mkdir(parents=True, exist_ok=True)
        prefix = f"{d.email.date:%Y-%m-%d}_{org_slug(d.rule.name)}_{d.rule.currency}_{d.amount:.2f}"
        notes = []
        email_file = unique_path(folder / f"{prefix}_email.pdf")
        if not renderer.render(d.email_html_path, email_file):
            email_file = unique_path(folder / f"{prefix}_email.html")
            shutil.copy2(d.email_html_path, email_file)
            notes.append(RENDER_FAILED)
            log.append(f"{RENDER_FAILED}: {d.email_html_path}")
        written.append(email_file)
        receipts = []
        for source in d.receipt_paths:
            target = unique_path(folder / f"{prefix}_receipt{source.suffix.lower() or '.pdf'}")
            shutil.copy2(source, target)
            receipts.append(target)
            written.append(target)
        rows.append(row_from_donation(d, email_file=email_file, receipt_files=receipts, extra_notes=notes))
    write_workbook(pack_dir / f"donations_{year}_summary.xlsx", rows, summary, None)
    return written
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest tests/test_accountant.py -v`
Expected: 2 passed

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/accountant.py tests/test_accountant.py
git commit -m "feat: accountant pack with PDF emails and receipts"
```

---

### Task 11: Pipeline, CLI and end-to-end test

**Files:**
- Create: `donation_tracker/pipeline.py`, `donation_tracker/cli.py`, `donations.py`, `rules.yaml` (empty list for now: `[]`)
- Test: `tests/test_pipeline.py`, `tests/test_cli.py`

**Interfaces:**
- Consumes: everything above.
- Produces: `pipeline.collect(accounts, rules, detector, year, log) -> tuple[list[Donation], list[Email], int]`; `pipeline.run(year, accounts, rules, detector, out_root: Path, renderer) -> RunResult`; `RunResult(out_dir, donations, needs_rule, summary, log)`; `cli.main(argv: list[str] | None = None) -> int`; `cli.format_summary(result: RunResult) -> str`.

- [ ] **Step 1: Write the failing tests**

`tests/test_pipeline.py`:
```python
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from donation_tracker.config import Account
from donation_tracker.detector import load_detector
from donation_tracker.pipeline import run
from tests.helpers import make_email, make_mbox, make_text_pdf, rule_record
from tests.test_accountant import FakeRenderer

ROOT = Path(__file__).resolve().parents[1]


def build_accounts(tmp_path):
    shared = make_email(sender="Muslim Hands <mail@muslimhands.ca>", subject="Your Receipt from Muslim Hands",
                        html="<p>Donation Amount: $1,200.00</p><p>Receipt No: MH-778</p>",
                        message_id="<mh1@muslimhands.ca>", date="Fri, 17 Oct 2025 10:00:00 -0400")
    a = [
        shared,
        make_email(sender="MyTenNights <contact@mytennights.com>", subject="Your 21st Night Donation",
                   text="Tonight's donation $7.17", message_id="<n21@mtn>", date="Fri, 21 Mar 2025 22:00:00 +0000"),
        make_email(sender="MyTenNights <contact@mytennights.com>", subject="Your Annual Tax Receipt",
                   text="See attached", message_id="<annual@mtn>", date="Tue, 24 Feb 2026 09:00:00 +0000",
                   attachments=[("MTN2025.pdf", "application/pdf", make_text_pdf("Total donated $54.60"))]),
        make_email(sender="MyTenNights <contact@mytennights.com>", subject="Your Annual Tax Receipt",
                   text="Total donated $99.00", message_id="<annual-prev@mtn>", date="Mon, 24 Feb 2025 09:00:00 +0000"),
        make_email(sender="Human Appeal USA <customercare@humanappeal.org>", subject="Thank you",
                   text="No amount here", message_id="<ha@x>", date="Sat, 21 Jun 2025 09:00:00 +0000"),
        make_email(sender="New Charity <hello@newcharity.org>", subject="Thank you for your donation",
                   text="We received it", message_id="<new@x>", date="Mon, 05 May 2025 09:00:00 +0000"),
        make_email(sender="Friend <f@gmail.com>", subject="Lunch", text="hi", message_id="<f@x>",
                   date="Mon, 05 May 2025 09:00:00 +0000"),
    ]
    b = [shared]  # same Message-ID in second mailbox → de-duplicated
    return [
        Account("abbasirumael@gmail.com", make_mbox(tmp_path / "a.mbox", a)),
        Account("itsrumzy@gmail.com", make_mbox(tmp_path / "b.mbox", b)),
    ]


def build_rules():
    return [
        rule_record(name="Muslim Hands Canada", match_from=["muslimhands.ca"], match_subject=["Your Receipt"]),
        rule_record(name="MyTenNights nightly", country="US", currency="USD", role="confirmation",
                    match_from=["mytennights.com"], match_subject=["Night Donation"]),
        rule_record(name="MyTenNights annual", country="US", currency="USD", tax_year="previous",
                    match_from=["mytennights.com"], match_subject=["Annual Tax Receipt"]),
        rule_record(name="Human Appeal USA", country="US", currency="USD", match_from=["humanappeal"]),
    ]


def test_end_to_end(tmp_path):
    out_root = tmp_path / "output"
    (out_root / "2025").mkdir(parents=True)
    (out_root / "2025" / "stale.txt").write_text("old", encoding="utf-8")
    result = run(2025, build_accounts(tmp_path), build_rules(), load_detector(ROOT / "detector.yaml"),
                 out_root, FakeRenderer())

    assert not (out_root / "2025" / "stale.txt").exists()
    names = [d.rule.name for d in result.donations]
    assert names == ["MyTenNights nightly", "Human Appeal USA", "Muslim Hands Canada", "MyTenNights annual"]
    annual = result.donations[-1]
    assert annual.amount == Decimal("54.60")
    assert annual.email.date.year == 2026
    assert result.summary.totals == {("CA", "CAD"): Decimal("1200.00"), ("US", "USD"): Decimal("54.60")}
    assert result.summary.needs_review == 1
    assert [e.message_id for e in result.needs_rule] == ["<new@x>"]
    assert result.donations[2].gmail_url.endswith("rfc822msgid:mh1%40muslimhands.ca")

    year_dir = out_root / "2025"
    wb = load_workbook(year_dir / "donations_2025.xlsx")
    ws = wb["Donations"]
    assert ws.max_row == 5
    for row in ws.iter_rows(min_row=2):
        for cell in (row[11], row[12]):
            if cell.hyperlink:
                assert (year_dir / cell.hyperlink.target).exists(), cell.hyperlink.target
    assert "@newcharity.org" in (year_dir / "suggested_rules.yaml").read_text(encoding="utf-8")
    assert (year_dir / "run_log.txt").exists()
    pack = year_dir / "for_accountant"
    assert sorted(p.name for p in (pack / "CA").iterdir()) == [
        "2025-10-17_Muslim_Hands_Canada_CAD_1200.00_email.pdf"]
    assert sorted(p.name for p in (pack / "US").iterdir()) == [
        "2026-02-24_MyTenNights_annual_USD_54.60_email.pdf",
        "2026-02-24_MyTenNights_annual_USD_54.60_receipt.pdf"]
    pack_ws = load_workbook(pack / "donations_2025_summary.xlsx")["Donations"]
    for row in pack_ws.iter_rows(min_row=2):
        for cell in (row[11], row[12]):
            if cell.hyperlink:
                assert (pack / cell.hyperlink.target).exists()
```

Note: `<annual-prev@mtn>` (dated Feb 2025) belongs to tax year 2024 and must NOT appear — the `names` assertion covers it.

`tests/test_cli.py`:
```python
from pathlib import Path

from donation_tracker import cli
from tests.helpers import make_email, make_mbox
from tests.test_accountant import FakeRenderer

ROOT = Path(__file__).resolve().parents[1]


def setup_files(tmp_path):
    make_mbox(tmp_path / "a.mbox", [
        make_email(sender="Muslim Hands <mail@muslimhands.ca>", subject="Your Receipt from Muslim Hands",
                   text="Donation Amount: $50.00", date="Fri, 17 Oct 2025 10:00:00 -0400"),
    ])
    (tmp_path / "accounts.yaml").write_text(
        "accounts:\n  - email: abbasirumael@gmail.com\n    mbox: a.mbox\n", encoding="utf-8")
    (tmp_path / "rules.yaml").write_text(
        "- name: Muslim Hands Canada\n  country: CA\n  currency: CAD\n  match:\n    from: muslimhands.ca\n",
        encoding="utf-8")
    return [
        "--accounts", str(tmp_path / "accounts.yaml"),
        "--rules", str(tmp_path / "rules.yaml"),
        "--detector", str(ROOT / "detector.yaml"),
    ]


def test_run_command(tmp_path, capsys, monkeypatch):
    common = setup_files(tmp_path)
    monkeypatch.setattr(cli, "EdgeRenderer", FakeRenderer)
    code = cli.main(["run", "--year", "2025", "--out", str(tmp_path / "output"), *common])
    assert code == 0
    out = capsys.readouterr().out
    assert "CA / CAD: 50.00" in out
    assert (tmp_path / "output" / "2025" / "donations_2025.xlsx").exists()


def test_test_rule_command_prints_matches_without_writing(tmp_path, capsys):
    common = setup_files(tmp_path)
    code = cli.main(["test-rule", "Muslim Hands Canada", "--year", "2025", *common])
    assert code == 0
    out = capsys.readouterr().out
    assert "Your Receipt from Muslim Hands" in out
    assert "50.00" in out
    assert "1 email(s) matched" in out
    assert not (tmp_path / "output").exists()


def test_unknown_rule_and_config_errors(tmp_path, capsys):
    common = setup_files(tmp_path)
    assert cli.main(["test-rule", "Nope", "--year", "2025", *common]) == 2
    (tmp_path / "rules.yaml").write_text("- name: X\n  country: FR\n", encoding="utf-8")
    assert cli.main(["test-rule", "X", "--year", "2025", *common]) == 2
    assert "Config error" in capsys.readouterr().err
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `.venv\Scripts\python -m pytest tests/test_pipeline.py tests/test_cli.py -v`
Expected: FAIL — module not found

- [ ] **Step 3: Write implementation**

`donation_tracker/pipeline.py`:
```python
"""Orchestrate one yearly run: read → match → extract → evidence → report → accountant pack."""
from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from .accountant import build_pack
from .config import Account
from .detector import DetectorConfig, is_candidate, suggested_rules_yaml
from .evidence import save_evidence
from .extractor import extract_donation, tax_year_for
from .links import gmail_url
from .matcher import match_rule
from .mbox_reader import iter_emails
from .models import Donation, Email, Rule
from .pdf_text import attachments_text
from .report import Summary, row_from_donation, summarize, write_workbook


@dataclass
class RunResult:
    out_dir: Path
    donations: list[Donation]
    needs_rule: list[Email]
    summary: Summary
    log: list[str]


def collect(accounts: list[Account], rules: list[Rule], detector: DetectorConfig, year: int,
            log: list[str]) -> tuple[list[Donation], list[Email], int]:
    seen: set[str] = set()
    donations: list[Donation] = []
    needs_rule: list[Email] = []
    skipped = 0

    def on_error(message: str) -> None:
        nonlocal skipped
        skipped += 1
        log.append(f"SKIPPED {message}")

    for account in accounts:
        for email in iter_emails(account.mbox, account.email, {year, year + 1}, on_error):
            if email.message_id in seen:
                continue
            seen.add(email.message_id)
            rule, others = match_rule(email, rules)
            if rule is None:
                if email.date.year == year and is_candidate(email, detector):
                    needs_rule.append(email)
                continue
            if tax_year_for(email, rule) != year:
                continue
            if others:
                log.append(f"CONFLICT {email.message_id}: used '{rule.name}', also matched "
                           + ", ".join(f"'{r.name}'" for r in others))
            pdf_text, textless = attachments_text(email.attachments)
            donation = extract_donation(email, rule, email.body_text + "\n" + pdf_text, textless)
            donation.gmail_url = gmail_url(email.account, email.message_id)
            donations.append(donation)
    donations.sort(key=lambda d: d.email.date)
    return donations, needs_rule, skipped


def run(year: int, accounts: list[Account], rules: list[Rule], detector: DetectorConfig, out_root: Path,
        renderer) -> RunResult:
    out_dir = out_root / str(year)
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)
    log: list[str] = []
    donations, needs_rule, skipped = collect(accounts, rules, detector, year, log)
    for donation in donations:
        save_evidence(donation, out_dir / "evidence")
    summary = summarize(donations, len(needs_rule), skipped)
    write_workbook(out_dir / f"donations_{year}.xlsx", [row_from_donation(d) for d in donations], summary,
                   needs_rule)
    (out_dir / "suggested_rules.yaml").write_text(suggested_rules_yaml(needs_rule), encoding="utf-8")
    build_pack(donations, out_dir / "for_accountant", renderer, year, summary, log)
    (out_dir / "run_log.txt").write_text("\n".join(log) + "\n", encoding="utf-8")
    return RunResult(out_dir, donations, needs_rule, summary, log)
```

`donation_tracker/cli.py`:
```python
"""Command line: `python donations.py run --year 2025` and `python donations.py test-rule NAME --year 2025`."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import ConfigError, load_accounts, load_rules
from .detector import load_detector
from .extractor import extract_donation, tax_year_for
from .matcher import rule_matches
from .mbox_reader import iter_emails
from .pdf_render import EdgeRenderer
from .pdf_text import attachments_text
from .pipeline import RunResult, run


def format_summary(result: RunResult) -> str:
    s = result.summary
    lines = ["Totals (receipts with status OK):"]
    if not s.totals:
        lines.append("  (none)")
    for (country, currency), total in sorted(s.totals.items()):
        lines.append(f"  {country} / {currency}: {total:,.2f}")
    lines += [
        f"Rows needing review:   {s.needs_review}",
        f"Emails needing a rule: {s.needs_rule}  (see suggested_rules.yaml)",
        f"Malformed skipped:     {s.skipped}",
        f"Output: {result.out_dir}",
    ]
    return "\n".join(lines)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="donations.py", description="Donation receipts from Gmail Takeout")
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--year", type=int, required=True)
    common.add_argument("--accounts", type=Path, default=Path("accounts.yaml"))
    common.add_argument("--rules", type=Path, default=Path("rules.yaml"))
    common.add_argument("--detector", type=Path, default=Path("detector.yaml"))
    sub = parser.add_subparsers(dest="command", required=True)
    run_parser = sub.add_parser("run", parents=[common], help="build the report for one year")
    run_parser.add_argument("--out", type=Path, default=Path("output"))
    test_parser = sub.add_parser("test-rule", parents=[common], help="show what one rule matches")
    test_parser.add_argument("name")
    return parser


def _test_rule(args) -> int:
    rules = load_rules(args.rules)
    selected = [r for r in rules if r.name == args.name]
    if not selected:
        print(f"No rule named {args.name!r}. Known: {', '.join(r.name for r in rules)}", file=sys.stderr)
        return 2
    rule = selected[0]
    count = 0
    for account in load_accounts(args.accounts):
        for email in iter_emails(account.mbox, account.email, {args.year, args.year + 1}):
            if not rule_matches(rule, email) or tax_year_for(email, rule) != args.year:
                continue
            pdf_text, textless = attachments_text(email.attachments)
            d = extract_donation(email, rule, email.body_text + "\n" + pdf_text, textless)
            count += 1
            amount = f"{d.amount:,.2f}" if d.amount is not None else "-"
            print(f"{email.date:%Y-%m-%d}  {account.email}  {email.subject}")
            print(f"    amount={amount} {rule.currency}  receipt#={d.receipt_no or '-'}  "
                  f"charity_id={d.charity_id or '-'}  status={d.status}  {'; '.join(d.notes)}")
    print(f"{count} email(s) matched rule {rule.name!r} for tax year {args.year}")
    return 0


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "test-rule":
            return _test_rule(args)
        accounts = load_accounts(args.accounts)
        rules = load_rules(args.rules)
        detector = load_detector(args.detector)
        result = run(args.year, accounts, rules, detector, args.out, EdgeRenderer())
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 2
    except PermissionError as exc:
        print(f"Cannot write output ({exc}). Close any open report files in Excel and re-run.", file=sys.stderr)
        return 1
    print(format_summary(result))
    return 0
```

`donations.py`:
```python
"""Entry point: python donations.py run --year 2025"""
import sys

from donation_tracker.cli import main

if __name__ == "__main__":
    sys.exit(main())
```

`rules.yaml` (temporary; Task 12 fills it):
```yaml
# One entry per organization / email type. See docs/superpowers/specs/2026-09-28-donation-tracker-design.md
[]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `.venv\Scripts\python -m pytest -v`
Expected: all tests pass (Edge test may skip only if Edge is missing).

- [ ] **Step 5: Commit**

```powershell
git add donation_tracker/pipeline.py donation_tracker/cli.py donations.py rules.yaml tests/test_pipeline.py tests/test_cli.py
git commit -m "feat: yearly pipeline, CLI and end-to-end test"
```

---

### Task 12: Initial rules from the real mailboxes + regression fixtures

This task works against the user's real data and **requires the user at the checkpoints**. Scanning a 1.6 GB mbox takes several minutes per run; run commands in the background where possible.

**Files:**
- Create: `tools/make_fixture.py`, `tests/test_real_rules.py`
- Modify: `rules.yaml`
- Create (git-ignored): `tests/fixtures/real/*.json`

**Interfaces:**
- Consumes: `config.load_accounts`, `config.load_rules`, `mbox_reader.iter_emails`, `pdf_text.attachments_text`, `matcher.match_rule`, `extractor.extract_donation`, `models.Email`, `models.Attachment`.
- Produces: fixture JSON schema `{"sender", "subject", "date" (ISO), "has_pdf", "textless_pdf", "text", "expect": {"rule", "amount" (str|null), "receipt_no", "status"}}`.

- [ ] **Step 1: Write the regression test and fixture tool**

`tests/test_real_rules.py`:
```python
"""Regression tests over anonymized real emails (tests/fixtures/real/*.json, git-ignored)."""
import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import pytest

from donation_tracker.config import load_rules
from donation_tracker.extractor import extract_donation
from donation_tracker.matcher import match_rule
from donation_tracker.models import Attachment, Email

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = sorted((Path(__file__).parent / "fixtures" / "real").glob("*.json"))


@pytest.mark.parametrize("path", FIXTURES, ids=lambda p: p.stem)
def test_real_fixture(path):
    fx = json.loads(path.read_text(encoding="utf-8"))
    attachments = [Attachment("receipt.pdf", "application/pdf", b"")] if fx["has_pdf"] else []
    email = Email("fixture@example.com", f"<{path.stem}>", datetime.fromisoformat(fx["date"]),
                  fx["sender"], fx["subject"], fx["text"], None, attachments, b"")
    rule, _ = match_rule(email, load_rules(ROOT / "rules.yaml"))
    expect = fx["expect"]
    assert rule is not None and rule.name == expect["rule"]
    d = extract_donation(email, rule, fx["text"], fx["textless_pdf"])
    assert d.amount == (Decimal(expect["amount"]) if expect["amount"] is not None else None)
    assert d.receipt_no == expect["receipt_no"]
    assert d.status == expect["status"]
```

`tools/make_fixture.py`:
```python
"""Save one real email as an anonymized regression fixture.

Usage:
  .venv\\Scripts\\python tools\\make_fixture.py "<message-id@host>" --year 2025 --name muslim_hands_2025_10 \
      --redact "Rumael" "Abbasi" "Khadeeja" "Farooq" "123 Main St"
Prints the extracted values — CHECK THEM AGAINST THE EMAIL before keeping the fixture.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from donation_tracker.config import load_accounts, load_rules  # noqa: E402
from donation_tracker.extractor import extract_donation  # noqa: E402
from donation_tracker.matcher import match_rule  # noqa: E402
from donation_tracker.mbox_reader import iter_emails  # noqa: E402
from donation_tracker.pdf_text import attachments_text  # noqa: E402


def redact(value: str, words: list[str]) -> str:
    for word in words:
        value = re.sub(re.escape(word), "REDACTED", value, flags=re.IGNORECASE)
    value = re.sub(r"[\w.+-]+@gmail\.com", "person@example.com", value, flags=re.IGNORECASE)
    return value


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("message_id")
    parser.add_argument("--year", type=int, required=True, help="year the email was SENT")
    parser.add_argument("--name", required=True)
    parser.add_argument("--redact", nargs="*", default=[])
    args = parser.parse_args()
    rules = load_rules(ROOT / "rules.yaml")
    for account in load_accounts(ROOT / "accounts.yaml"):
        for email in iter_emails(account.mbox, account.email, {args.year}):
            if email.message_id != args.message_id:
                continue
            pdf_text, textless = attachments_text(email.attachments)
            text = email.body_text + "\n" + pdf_text
            rule, _ = match_rule(email, rules)
            if rule is None:
                print("No rule matches this email; add the rule first.", file=sys.stderr)
                return 1
            d = extract_donation(email, rule, text, textless)
            fixture = {
                "sender": redact(email.sender, args.redact),
                "subject": redact(email.subject, args.redact),
                "date": email.date.isoformat(),
                "has_pdf": any(a.is_pdf for a in email.attachments),
                "textless_pdf": textless,
                "text": redact(text, args.redact),
                "expect": {
                    "rule": rule.name,
                    "amount": str(d.amount) if d.amount is not None else None,
                    "receipt_no": d.receipt_no,
                    "status": d.status,
                },
            }
            out = ROOT / "tests" / "fixtures" / "real" / f"{args.name}.json"
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(json.dumps(fixture, indent=2, ensure_ascii=False), encoding="utf-8")
            print(json.dumps(fixture["expect"], indent=2))
            print(f"Wrote {out} — verify these values against the real email, and skim the file for personal data.")
            return 0
    print("Message-ID not found in that year.", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
```

Run: `.venv\Scripts\python -m pytest tests/test_real_rules.py -v`
Expected: 1 skipped ("got empty parameter set") — no fixtures yet.

Commit:
```powershell
git add tools/make_fixture.py tests/test_real_rules.py
git commit -m "test: real-email regression harness and fixture tool"
```

- [ ] **Step 2: Seed `rules.yaml` with hypotheses from attempt2's findings**

These senders/subjects were observed in `archived_attempts/attempt2/donation_reports/donations_2025.md` / `donations_2026.md`. Each is a hypothesis to verify in Step 3.

```yaml
# One entry per organization / email type. First match wins, so put specific rules above general ones.
# See docs/superpowers/specs/2026-09-28-donation-tracker-design.md for every field.

- name: Muslim Hands Canada
  country: CA
  currency: CAD
  role: receipt
  match:
    from: ["muslimhands.ca"]
    subject: ["Your Receipt from Muslim Hands"]

- name: Muslim Hands USA
  country: US
  currency: USD
  role: receipt
  match:
    from: ["muslimhandsusa.org"]
    subject: ["Your Receipt from Muslim Hands"]

- name: ICNA Relief Canada
  country: CA
  currency: CAD
  role: receipt
  match:
    from: ["no-reply@icnareliefcanada.ca"]
    subject: ["Donation receipt"]

- name: ICNA Relief Canada (card payment)
  country: CA
  currency: CAD
  role: confirmation          # Stripe payment email for the same gift as the official receipt above
  match:
    from: ["ICNA Relief Canada <receipts"]
    subject: ["ICNA Relief Canada receipt"]

- name: Hamilton Downtown Mosque
  country: CA
  currency: CAD
  role: receipt
  match:
    from: ["themasjidapp.net", "donorchoice.net"]
    subject: ["Hamilton Downtown Mosque"]

- name: MyTenNights nightly
  country: CA
  currency: CAD               # user confirmed 2026-09-28: MyTenNights donations are in CAD
  role: confirmation          # nightly confirmations; the annual receipt below is what counts
  match:
    from: ["mytennights.com"]
    subject: ["Night Donation", "Automated Donations"]

- name: MyTenNights annual receipt
  country: CA
  currency: CAD
  role: receipt
  tax_year: previous          # sent in February for the previous year's donations
  match:
    from: ["mytennights.com"]
    subject: ["receipt"]

- name: Muslim Center of Bloomfield Hills
  country: US
  currency: USD
  role: receipt
  match:
    from: ["Muslim Center Of Bloomfield Hills"]
    subject: ["Receipt"]

- name: Human Appeal USA
  country: US
  currency: USD
  role: receipt
  match:
    from: ["Human Appeal USA"]

- name: ICNA Relief USA
  country: US
  currency: USD
  role: receipt
  match:
    from: ["ICNA Relief USA"]
```

Run: `.venv\Scripts\python donations.py test-rule "Muslim Hands Canada" --year 2025` (config loads = rules are valid).

- [ ] **Step 3: Verify each seeded rule against real mail**

For every rule name above, run:
```powershell
.venv\Scripts\python donations.py test-rule "<rule name>" --year 2025
```
For each, check: (a) every listed email really is that org's donation email of that role; (b) amount is right; (c) nothing obviously missing (compare with the attempt2 report). Fix `match` / add `extract` regexes until correct. Known cases to confirm:
- Muslim Hands Canada 2025-10-17 → 1,200.00; ICNA Relief Canada 2025-08-29 → 50.00 (and the Stripe email shows as the confirmation rule, not the receipt rule).
- MyTenNights annual: find the real subject of the Feb 2026 email with attachment `Abbasi Rumael MTN2025_EI7KVALM.pdf`; tighten `subject` to it; confirm amount from the PDF. Currency is CAD (user-confirmed); country assumed CA — confirm the charity registration number in the PDF is a CRA BN (`#########RR####`).
- Hamilton Downtown Mosque receipts had no amount in attempt2's summary — expect to need a custom `extract.amount`.

- [ ] **Step 4: Apply the user's decisions on ambiguous emails** (answered 2026-09-28)

1. "Made in your honour" emails (Islamic Relief Canada "A donation was made in your honour!", Muslim Food Bank "generous gift in honour of …") are **gifts others made in the user's name — NOT the user's donations**. Add no rule; add these patterns to `detector.yaml` → `exclude_subject`:
   ```yaml
     - 'in your honou?r'
     - 'in honou?r of'
   ```
2. Receipts sent by/for Khadeeja Farooq (`Receipt-QF-2026-…pdf`, sender `khadeejafar01@…`) **stay in the user's report**. Identify the issuing organization from the PDF text, then add a rule named after that organization with `match.from: ["khadeejafar01"]` plus a subject/attachment-based discriminator so unrelated emails from her don't match. Add a YAML comment: `# forwarded by Khadeeja; user confirmed these count`.
3. MyTenNights: CAD (already in Step 2 seeds).
4. Muslim Center of Bloomfield Hills weekly "Email Receipt for Order" emails **are donations** — keep the rule as `role: receipt`, US/USD.

- [ ] **Step 5: Full run and triage "Needs a rule"**

```powershell
.venv\Scripts\python donations.py run --year 2025
```
Open `output\2025\donations_2025.xlsx` with the user. For each "Needs a rule" row: if it is a donation, copy the block from `output\2025\suggested_rules.yaml` into `rules.yaml`, fix it, and verify with `test-rule`; if it is not, add a pattern to `detector.yaml` (`exclude_subject` or `exclude_sender`). Repeat the run until "Needs a rule" is empty or contains only items the user has consciously decided to ignore, and every NEEDS REVIEW row is resolved by a rule fix. Then do the same with `--year 2026`.

- [ ] **Step 6: Create one regression fixture per rule**

For each rule, pick one representative email (Message-ID from the `.eml` in `output\<year>\evidence\…` — first `Message-ID:` header), then:
```powershell
.venv\Scripts\python tools\make_fixture.py "<message-id>" --year <year sent> --name <rule_slug>_<yyyy_mm> --redact "Rumael" "Abbasi" "Khadeeja" "Farooq"
```
Check the printed values against the email. Then:
Run: `.venv\Scripts\python -m pytest -v`
Expected: all pass, one `test_real_fixture[...]` per fixture.

- [ ] **Step 7: USER CHECKPOINT — sign-off on 2025 totals**

Show the user the console summary and the Summary sheet. Confirm the CA/CAD and US/USD totals match their own expectation, and open two random rows via the Gmail link and the Email/Receipt file links to prove traceability works in their Excel.

- [ ] **Step 8: Commit**

```powershell
git add rules.yaml detector.yaml
git commit -m "feat: initial donation rules verified against 2025-2026 mail"
```

---

### Task 13: README for yearly use

**Files:**
- Create: `README.md`

- [ ] **Step 1: Write `README.md`**

```markdown
# Donation Tracker

Builds a yearly donation report from Gmail Takeout exports.

## Each year (e.g. for 2026, run in early 2027 after annual receipts arrive)

1. Google Takeout → Mail → export each account; unzip so each account's
   `All mail Including Spam and Trash.mbox` sits in its folder (`abbasirumael/`, `itsrumzy/`).
2. `.venv\Scripts\python donations.py run --year 2026`
3. Open `output\2026\donations_2026.xlsx`:
   - **Summary** — totals per country/currency (only receipts with status OK).
   - **Donations** — one row per email; Gmail / Email file / Receipt file links.
   - **Needs a rule** — emails that look like donations but no rule matched.
4. For each "Needs a rule" email that is a real donation, copy its block from
   `output\2026\suggested_rules.yaml` into `rules.yaml`, adjust, and check it:
   `.venv\Scripts\python donations.py test-rule "Org Name" --year 2026`
   Then re-run step 2.
5. Send `output\2026\for_accountant\` (zip it) to your accountant.

## Adding a rule

    - name: Org Name
      country: CA            # CA or US
      currency: CAD          # CAD or USD
      role: receipt          # receipt (counted) or confirmation (shown, not counted)
      tax_year: same         # or 'previous' for annual receipts sent the next year
      match:
        from: ["@org.ca"]            # any of these appears in the From header
        subject: ["tax receipt"]     # optional; any of these appears in the Subject
      extract:                       # optional; only if the default misses the amount
        amount: 'Total:\s*\$([\d,]+\.\d{2})'

First matching rule wins — put specific rules above general ones.
Not a donation but keeps showing under "Needs a rule"? Add to `detector.yaml` → `exclude_subject` / `exclude_sender`.

## Adding a mailbox

Add an entry to `accounts.yaml` with the account's email and the path to its `.mbox`.

## Tests

    .venv\Scripts\python -m pytest
```

- [ ] **Step 2: Commit**

```powershell
git add README.md
git commit -m "docs: yearly usage and rule-writing guide"
```
