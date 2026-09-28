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
