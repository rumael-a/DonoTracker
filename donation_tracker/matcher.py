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
