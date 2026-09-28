"""Orchestrate one yearly run: read → match → extract → evidence → report → accountant pack."""
from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from .accountant import build_pack
from .config import Account
from .detector import DetectorConfig, is_candidate, suggested_rules_yaml
from .evidence import save_evidence
from .extractor import extract_donation, tax_year_for
from .links import gmail_url
from .matcher import match_rule
from .mbox_reader import iter_emails
from .models import NEEDS_REVIEW, Donation, Email, Rule
from .pdf_text import attachments_text
from .report import Summary, row_from_donation, summarize, write_workbook

SPAM_NOTE = "In Gmail Spam folder — verify"


@dataclass
class RunResult:
    out_dir: Path
    donations: list[Donation]
    needs_rule: list[Email]
    summary: Summary
    log: list[str]


def flag_possible_duplicates(donations: list[Donation], log: list[str]) -> None:
    """Note (don't drop) counted receipts that repeat an earlier one; `donations` must be sorted by date."""
    first_seen: dict[tuple, Donation] = {}
    for d in donations:
        if not d.is_countable:
            continue
        if d.receipt_no is not None:
            key = (d.rule.name, d.receipt_no)
        else:
            key = (d.rule.name, d.email.date.date(), d.amount)
        earlier = first_seen.setdefault(key, d)
        if earlier is d:
            continue
        d.notes.append(f"Possible duplicate of {earlier.email.date:%Y-%m-%d} row")
        log.append(f"WARNING possible duplicate '{d.rule.name}': {d.email.date:%Y-%m-%d} {d.email.message_id} "
                   f"repeats {earlier.email.date:%Y-%m-%d} {earlier.email.message_id}")


def _needs_rule_candidate(email: Email, detector: DetectorConfig, year: int) -> bool:
    """Unmatched donation-like mail from the year, or from Jan-Apr of the next year when it mentions
    the year (an annual receipt that no rule covers yet)."""
    if email.date.year == year:
        return is_candidate(email, detector)
    early_next_year = email.date.year == year + 1 and email.date.month < 5
    mentions_year = str(year) in email.subject or str(year) in email.body_text
    return early_next_year and mentions_year and is_candidate(email, detector)


def collect(accounts: list[Account], rules: list[Rule], detector: DetectorConfig, year: int,
            log: list[str]) -> tuple[list[Donation], list[Email], int]:
    seen: set[str] = set()
    donations: list[Donation] = []
    needs_rule: list[Email] = []
    skipped = 0

    def on_error(message: str) -> None:
        nonlocal skipped
        skipped += 1
        log.append(f"SKIPPED {message}")

    for account in accounts:
        for email in iter_emails(account.mbox, account.email, {year, year + 1}, on_error):
            if email.message_id in seen:
                continue
            seen.add(email.message_id)
            rule, others = match_rule(email, rules)
            if rule is None:
                if _needs_rule_candidate(email, detector, year):
                    needs_rule.append(email)
                continue
            if tax_year_for(email, rule) != year:
                continue
            if others:
                log.append(f"CONFLICT {email.message_id}: used '{rule.name}', also matched "
                           + ", ".join(f"'{r.name}'" for r in others))
            pdf_text, textless = attachments_text(email.attachments)
            donation = extract_donation(email, rule, email.body_text + "\n" + pdf_text, textless)
            if "spam" in email.labels.lower():
                donation.status = NEEDS_REVIEW
                donation.notes.append(SPAM_NOTE)
            donation.gmail_url = gmail_url(email.account, email.message_id)
            donations.append(donation)
    donations.sort(key=lambda d: d.email.date)
    flag_possible_duplicates(donations, log)
    return donations, needs_rule, skipped


def run(year: int, accounts: list[Account], rules: list[Rule], detector: DetectorConfig, out_root: Path,
        renderer) -> RunResult:
    out_dir = out_root / str(year)
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)
    log: list[str] = []
    try:
        donations, needs_rule, skipped = collect(accounts, rules, detector, year, log)
        for donation in donations:
            save_evidence(donation, out_dir / "evidence")
        summary = summarize(donations, len(needs_rule), skipped)
        write_workbook(out_dir / f"donations_{year}.xlsx", [row_from_donation(d) for d in donations], summary,
                       needs_rule, year)
        (out_dir / "suggested_rules.yaml").write_text(suggested_rules_yaml(needs_rule), encoding="utf-8")
        build_pack(donations, out_dir / "for_accountant", renderer, year, summary, log)
    finally:
        (out_dir / "run_log.txt").write_text("\n".join(log) + "\n", encoding="utf-8")
    return RunResult(out_dir, donations, needs_rule, summary, log)
