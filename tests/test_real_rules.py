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
