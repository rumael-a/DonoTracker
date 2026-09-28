"""Command line: `python donations.py run --year 2025` and `python donations.py test-rule NAME --year 2025`."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import ConfigError, load_accounts, load_rules
from .detector import load_detector
from .extractor import extract_donation, tax_year_for
from .matcher import rule_matches
from .mbox_reader import iter_emails
from .pdf_render import EdgeRenderer
from .pdf_text import attachments_text
from .pipeline import RunResult, run


def format_summary(result: RunResult) -> str:
    s = result.summary
    lines = ["Totals (receipts with status OK):"]
    if not s.totals:
        lines.append("  (none)")
    for (country, currency), total in sorted(s.totals.items()):
        lines.append(f"  {country} / {currency}: {total:,.2f}")
    lines += [
        f"Rows needing review:   {s.needs_review}",
        f"Emails needing a rule: {s.needs_rule}  (see suggested_rules.yaml)",
        f"Malformed skipped:     {s.skipped}",
        f"Output: {result.out_dir}",
    ]
    return "\n".join(lines)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="donations.py", description="Donation receipts from Gmail Takeout")
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--year", type=int, required=True)
    common.add_argument("--accounts", type=Path, default=Path("accounts.yaml"))
    common.add_argument("--rules", type=Path, default=Path("rules.yaml"))
    common.add_argument("--detector", type=Path, default=Path("detector.yaml"))
    sub = parser.add_subparsers(dest="command", required=True)
    run_parser = sub.add_parser("run", parents=[common], help="build the report for one year")
    run_parser.add_argument("--out", type=Path, default=Path("output"))
    test_parser = sub.add_parser("test-rule", parents=[common], help="show what one rule matches")
    test_parser.add_argument("name")
    return parser


def _test_rule(args) -> int:
    rules = load_rules(args.rules)
    selected = [r for r in rules if r.name == args.name]
    if not selected:
        print(f"No rule named {args.name!r}. Known: {', '.join(r.name for r in rules)}", file=sys.stderr)
        return 2
    rule = selected[0]
    count = 0
    for account in load_accounts(args.accounts):
        for email in iter_emails(account.mbox, account.email, {args.year, args.year + 1}):
            if not rule_matches(rule, email) or tax_year_for(email, rule) != args.year:
                continue
            pdf_text, textless = attachments_text(email.attachments)
            d = extract_donation(email, rule, email.body_text + "\n" + pdf_text, textless)
            count += 1
            amount = f"{d.amount:,.2f}" if d.amount is not None else "-"
            print(f"{email.date:%Y-%m-%d}  {account.email}  {email.subject}")
            print(f"    amount={amount} {rule.currency}  receipt#={d.receipt_no or '-'}  "
                  f"charity_id={d.charity_id or '-'}  status={d.status}  {'; '.join(d.notes)}")
    print(f"{count} email(s) matched rule {rule.name!r} for tax year {args.year}")
    return 0


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "test-rule":
            return _test_rule(args)
        accounts = load_accounts(args.accounts)
        rules = load_rules(args.rules)
        detector = load_detector(args.detector)
        result = run(args.year, accounts, rules, detector, args.out, EdgeRenderer())
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 2
    except PermissionError as exc:
        print(f"Cannot write output ({exc}). Close any open report files in Excel and re-run.", file=sys.stderr)
        return 1
    print(format_summary(result))
    return 0
