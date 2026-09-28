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
        "match": {"from": ["Food Bank <noreply@foodbank.ca"], "subject": ["Your receipt"]},
    }
    us = starter_rule(email_record(sender="noreply@charity.org", body_text="tax-deductible"))
    assert us["country"] == "US" and us["currency"] == "USD" and us["name"] == "charity.org"
    assert us["match"]["from"] == ["@charity.org"]
    cra = starter_rule(email_record(sender="x@charity.org", body_text="Registered charity 123456789RR0001"))
    assert cra["country"] == "CA"


def test_suggested_rules_yaml_dedupes_by_sender_name_and_domain_and_is_valid_yaml():
    emails = [
        email_record(sender="A <a@one.ca>", subject="Receipt 1"),
        email_record(sender="A <b@one.ca>", subject="Receipt 2"),
        email_record(sender="B <x@two.org>", subject="Thanks"),
    ]
    text = suggested_rules_yaml(emails)
    assert text.count("# REVIEW") == 2
    assert "seen 2 email(s)" in text
    parsed = yaml.safe_load(text)
    assert [r["match"]["from"] for r in parsed] == [["A <a@one.ca"], ["B <x@two.org"]]


def test_suggested_rules_yaml_splits_shared_processor_domain_by_display_name():
    emails = [
        email_record(sender="Muslim Unity Center <contact@transaction.mailer.zeffy.com>", subject="Donation"),
        email_record(sender="Other Mosque <contact@transaction.mailer.zeffy.com>", subject="Donation"),
        email_record(sender="muslim unity center <contact@transaction.mailer.zeffy.com>", subject="Donation"),
    ]
    parsed = yaml.safe_load(suggested_rules_yaml(emails))
    assert [r["name"] for r in parsed] == ["Muslim Unity Center", "Other Mosque"]
    assert parsed[1]["match"]["from"] == ["Other Mosque <contact@transaction.mailer.zeffy.com"]


def test_bad_detector_regex(tmp_path):
    path = tmp_path / "d.yaml"
    path.write_text("candidate_subject: ['(']\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="candidate_subject"):
        load_detector(path)
