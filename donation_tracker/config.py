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
