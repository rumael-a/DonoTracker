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
