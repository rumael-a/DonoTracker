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
