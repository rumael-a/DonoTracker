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
