"""Save one real email as an anonymized regression fixture.

Usage:
  .venv\Scripts\python tools\make_fixture.py "<message-id@host>" --year 2025 --name muslim_hands_2025_10 \
      --redact "Rumael" "Abbasi" "Khadeeja" "Farooq" "123 Main St"
Prints the extracted values — CHECK THEM AGAINST THE EMAIL before keeping the fixture.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from donation_tracker.config import load_accounts, load_rules  # noqa: E402
from donation_tracker.extractor import extract_donation  # noqa: E402
from donation_tracker.matcher import match_rule  # noqa: E402
from donation_tracker.mbox_reader import iter_emails  # noqa: E402
from donation_tracker.pdf_text import attachments_text  # noqa: E402


def redact(value: str, words: list[str]) -> str:
    for word in words:
        value = re.sub(re.escape(word), "REDACTED", value, flags=re.IGNORECASE)
    value = re.sub(r"[\w.+-]+@gmail\.com", "person@example.com", value, flags=re.IGNORECASE)
    return value


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("message_id")
    parser.add_argument("--year", type=int, required=True, help="year the email was SENT")
    parser.add_argument("--name", required=True)
    parser.add_argument("--redact", nargs="*", default=[])
    args = parser.parse_args()
    rules = load_rules(ROOT / "rules.yaml")
    for account in load_accounts(ROOT / "accounts.yaml"):
        for email in iter_emails(account.mbox, account.email, {args.year}):
            if email.message_id != args.message_id:
                continue
            pdf_text, textless = attachments_text(email.attachments)
            text = email.body_text + "\n" + pdf_text
            rule, _ = match_rule(email, rules)
            if rule is None:
                print("No rule matches this email; add the rule first.", file=sys.stderr)
                return 1
            d = extract_donation(email, rule, text, textless)
            fixture = {
                "sender": redact(email.sender, args.redact),
                "subject": redact(email.subject, args.redact),
                "date": email.date.isoformat(),
                "has_pdf": any(a.is_pdf for a in email.attachments),
                "textless_pdf": textless,
                "text": redact(text, args.redact),
                "expect": {
                    "rule": rule.name,
                    "amount": str(d.amount) if d.amount is not None else None,
                    "receipt_no": d.receipt_no,
                    "status": d.status,
                },
            }
            out = ROOT / "tests" / "fixtures" / "real" / f"{args.name}.json"
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(json.dumps(fixture, indent=2, ensure_ascii=False), encoding="utf-8")
            print(json.dumps(fixture["expect"], indent=2))
            print(f"Wrote {out} — verify these values against the real email, and skim the file for personal data.")
            return 0
    print("Message-ID not found in that year.", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
