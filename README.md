# Donation Tracker

Builds a yearly donation report from Gmail Takeout exports.

## Setup (once)

1. Install Python 3.11+.
2. `python -m venv .venv`
3. `.venv\Scripts\python -m pip install -r requirements.txt`

Run all commands from the project folder — config paths are relative.

## Each year (e.g. for 2026, run in early 2027 after annual receipts arrive)

1. Google Takeout → Mail → export each account; unzip so each account's
   `All mail Including Spam and Trash.mbox` sits in its folder (`abbasirumael/`, `itsrumzy/`).
2. `.venv\Scripts\python donations.py run --year 2026`
3. Open `output\2026\donations_2026.xlsx`:
   - **Summary** — totals per country/currency (only receipts with status OK).
   - **Donations** — one row per email; Gmail / Email file / Receipt file links.
   - **Needs a rule** — emails that look like donations but no rule matched.
4. For each "Needs a rule" email that is a real donation, copy its block from
   `output\2026\suggested_rules.yaml` into `rules.yaml`, adjust, and check it:
   `.venv\Scripts\python donations.py test-rule "Org Name" --year 2026`
   Then re-run step 2.
5. Send `output\2026\for_accountant\` (zip it) to your accountant.

## Receipts that need manual handling

Neither receipt appears in the report automatically — get them yourself and add them to what you send your accountant:

- **Hamilton Downtown Mosque** — the year-end receipt is not emailed; request it from the mosque.
- **ICNA Relief Canada** — the official CRA receipt is download-only, via a link in their email.

## Adding a rule

    - name: Org Name
      country: CA            # CA or US
      currency: CAD          # CAD or USD
      role: receipt          # receipt (counted) or confirmation (shown, not counted)
      tax_year: same         # or 'previous' for annual receipts sent the next year
      match:
        from: ["@org.ca"]            # any of these appears in the From header
        subject: ["tax receipt"]     # optional; any of these appears in the Subject
      extract:                       # optional; only if the default misses the amount
        amount: 'Total:\s*\$([\d,]+\.\d{2})'

First matching rule wins — put specific rules above general ones.
Not a donation but keeps showing under "Needs a rule"? Add to `detector.yaml` → `exclude_subject` / `exclude_sender`.

## Adding a mailbox

Add an entry to `accounts.yaml` with the account's email and the path to its `.mbox`.

## Tests

    .venv\Scripts\python -m pytest

## Regression fixtures

Real-email fixtures live in `tests/fixtures/real/` (git-ignored, local only); create them with `.venv\Scripts\python tools\make_fixture.py "<message-id>" --year <year sent> --name <name> --redact <your name(s)> <your street> <your city> <your province/state>` and always open the written JSON to check no personal details remain before relying on it.
