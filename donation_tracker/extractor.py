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
