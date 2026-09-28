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
