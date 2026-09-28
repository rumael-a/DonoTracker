# Donation Tracker — Design

**Date:** 2026-09-28
**Status:** Approved in brainstorming; pending written-spec review

## Goal

A yearly-run, offline Python CLI that reads Google Takeout `.mbox` exports from three Gmail
accounts, identifies charitable donation receipts using a user-editable rules file, and
produces:

1. An Excel report with per-year totals and one row per donation email, each row traceable
   to the original email (Gmail link + local saved copy + receipt attachment).
2. A self-contained `for_accountant/` folder containing PDFs of verified receipt emails and
   their receipt attachments.

Scope: calendar years 2025 onward. Donations to both Canadian (CRA) and US (IRS)
organizations, in CAD and USD.

## Non-goals

- No live Gmail API access; input is Takeout `.mbox` only.
- No AI/LLM extraction; rules only.
- No incremental runs; each run rebuilds one year from scratch.
- No currency conversion; totals are kept per currency.
- No tax advice or eligibility determination beyond flagging missing receipts.

## Usage

```
python donations.py run --year 2025
python donations.py test-rule "SickKids Foundation" --year 2025
```

`accounts.yaml` maps each mbox to its Gmail address:

```yaml
accounts:
  - email: abbasirumael@gmail.com
    mbox: abbasirumael/All mail Including Spam and Trash.mbox
  - email: itsrumzy@gmail.com
    mbox: itsrumzy/All mail Including Spam and Trash.mbox
  # - email: rumaelabbasi@gmail.com   # no donation receipts yet; add its mbox here when needed
  #   mbox: <path>
```

Accounts can be added or removed in `accounts.yaml` without code changes.

## Architecture

Single Python package, one module per responsibility:

| Module | Responsibility | Input → Output |
|---|---|---|
| `mbox_reader` | Stream messages from each mbox, filter to target year, decode headers/bodies | mbox path, year → `Email` records |
| `pdf_text` | Extract text from PDF attachments | bytes → text (or "no text layer") |
| `rules` | Load and validate `rules.yaml` | file → list of `Rule` |
| `rule_matcher` | First-match-wins rule matching on sender/subject | `Email`, rules → `Rule` or None |
| `extractor` | Pull amount, receipt #, charity ID, date using rule regexes or built-in defaults | `Email` + `Rule` → `Donation` |
| `candidate_detector` | Keyword test for "looks like a donation" on unmatched mail; produce starter rule | `Email` → bool, YAML block |
| `evidence` | Save `.eml`, `.html`, attachments; render PDF via headless Edge | `Donation` → file paths |
| `report_writer` | Build `donations_<year>.xlsx` | donations, candidates → xlsx |
| `accountant_pack` | Build `for_accountant/` folder | verified receipts → folder |
| `donations.py` | CLI entry point (`run`, `test-rule`) | args → orchestrates the above |

### `Email` record

`account, message_id, date, sender, subject, body_text, body_html, attachments[(filename, content_type, bytes)], raw_bytes`

### `Donation` record

`date, organization, country, currency, amount (Decimal|None), role, receipt_no, charity_id,
account, status (OK|NEEDS REVIEW), notes, gmail_url, email_html_path, email_eml_path,
receipt_paths[]`

## Data flow

1. Load and validate `accounts.yaml` and `rules.yaml`. Any error → stop with a message naming
   the file/rule; no mail is processed.
2. For each account, stream its mbox; keep messages whose `Date` falls in the target year or
   the following year (see `tax_year` under Rules file).
3. De-duplicate across accounts by `Message-ID` (first account seen wins).
4. Match each message against rules on headers only (first match wins; additional matches
   logged as conflicts).
5. For matched messages whose tax year equals the target year: build searchable text = body
   text + text of every PDF attachment, extract fields → `Donation`.
   - Unmatched, dated in the target year, and `candidate_detector` says yes (subject + body
     text; PDFs not read, for speed) → "Needs a rule" list + starter YAML block.
   - Otherwise ignored.
6. Save evidence for every `Donation`.
7. Write the Excel report, `suggested_rules.yaml`, `run_log.txt`, and the accountant pack.
8. Print console summary: totals per country/currency, NEEDS REVIEW count, Needs-a-rule
   count, output path.

## Rules file (`rules.yaml`)

```yaml
- name: SickKids Foundation
  country: CA            # CA | US
  currency: CAD          # CAD | USD
  role: receipt          # receipt | confirmation
  match:
    from: ["@sickkidsfoundation.com"]   # any of; substring of From address, case-insensitive
    subject: ["tax receipt"]            # optional; any of; substring, case-insensitive
  extract:                              # optional; each key overrides the built-in default
    amount: 'Amount:\s*\$?([\d,]+\.\d{2})'
    receipt_no: 'Receipt\s*#?\s*:?\s*(\w+)'
    charity_id: '...'
  charity_id: "10742 5582 RR0001"       # optional fixed value; used when extract.charity_id absent or fails
  tax_year: same                        # same (default) | previous — for annual receipts sent the next year
```

- `match.from` is a case-insensitive substring of the whole decoded `From` header (display
  name + address), so payment processors can be matched by display name, e.g.
  `"ICNA Relief Canada <receipts"` for Stripe-sent receipts.
- `tax_year: previous` attributes the email to the year before it was sent (e.g. MyTenNights'
  annual receipt, sent each February). To support this, a run for year N reads messages dated
  N and N+1; messages from N+1 are kept only when their rule has `tax_year: previous`.
  "Needs a rule" candidates come from year N only.

- `match.from` is required; `match.subject` is optional. Both conditions must hold when both
  are given.
- `extract` regexes use capture group 1. Searched over body text + PDF text.
- Built-in defaults:
  - amount: first dollar amount adjacent to `amount|total|gift|donation|donated|eligible`.
  - receipt_no: value following `receipt (no|number|#)`.
  - charity_id: CRA BN pattern `\d{9}\s?RR\s?\d{4}` or EIN pattern `\d{2}-\d{7}`.
  - donation date: the email `Date` header.
- Validation at load: required keys present, `country`/`currency`/`role`/`tax_year` in
  allowed sets, every regex compiles and has a capture group, names unique.

### Initial rule set

Built during implementation by scanning the provided mboxes for 2025+ donation emails.
Known starting organizations from earlier attempts: Islamic Relief Canada, Muslim Hands
Canada, ICNA Relief Canada, ICNA Relief USA, Human Appeal USA, MyTenNights, Hamilton
Downtown Mosque. Each rule ships with a regression test fixture (see Testing).

## Status and totals logic

- `status = NEEDS REVIEW` when: amount not extracted, or a PDF attachment had no text layer
  and the amount came from nowhere else.
- Totals (Summary sheet, console) include only rows with `role = receipt` and `status = OK`.
- `role = confirmation` rows appear in the Donations sheet for reference; never totalled.
- US flag: rows with `country = US`, `amount >= 250` and no receipt number and no receipt
  attachment get a note "IRS: written acknowledgment required for $250+".

## Gmail link

`https://mail.google.com/mail/u/<account email>/#search/rfc822msgid:<url-encoded Message-ID>`

Opens the exact message in the correct account (user must be signed in to it).

## Outputs

```
output/2025/
  donations_2025.xlsx
  suggested_rules.yaml
  run_log.txt
  evidence/
    2025-03-14_SickKids_Foundation/
      email.eml
      email.html
      receipt.pdf            # original attachment name, sanitized
  for_accountant/
    donations_2025_summary.xlsx
    CA/
      2025-03-14_SickKids_Foundation_CAD_100.00_email.pdf
      2025-03-14_SickKids_Foundation_CAD_100.00_receipt.pdf
    US/
      ...
```

Evidence folder name collisions get a `_2`, `_3` suffix. All workbook links are relative,
so each folder can be zipped and shared.

### `donations_<year>.xlsx`

- **Summary:** totals per country/currency; subtotals per organization; counts of NEEDS
  REVIEW rows, Needs-a-rule emails, skipped malformed messages.
- **Donations:** columns `Date, Organization, Country, Currency, Amount, Role, Receipt #,
  Charity ID, Account, Status, Gmail, Email file, Receipt file, Notes`. Sorted by date.
  Hyperlinks for Gmail / Email file / Receipt file. NEEDS REVIEW rows highlighted yellow.
  Multiple receipt attachments → first in the Receipt file column, the rest listed in Notes.
  The Notes column is regenerated each run (user notes are not preserved).
- **Needs a rule:** `Date, Account, From, Subject, Gmail`.

### `for_accountant/`

- Contains only `role = receipt` AND `status = OK` rows.
- Each email rendered to PDF with headless Microsoft Edge
  (`msedge --headless --print-to-pdf`), from the same HTML used for `email.html` (with a
  Date/From/Subject header block).
- Receipt attachments copied alongside with the same filename prefix.
- Split into `CA/` and `US/`.
- `donations_<year>_summary.xlsx` has the Summary and Donations sheets with links pointing
  to the files inside `for_accountant/` (no Gmail-link dependency required to use it,
  though the Gmail column is kept).
- If Edge PDF rendering fails for an email, the `.html` is copied instead and the row's
  Notes say so.

## Candidate detection ("Needs a rule")

Unmatched messages are candidates when the subject or body matches donation keywords
(`donat`, `tax receipt`, `official receipt`, `charitable`, `501(c)(3)`, `registered
charity`, `thank you for your (gift|donation)`, `zakat`, `sadaqah`, …) and are not excluded
by solicitation-subject patterns (`donate to`, `giving tuesday`, `matching gift`, …) or
known commercial senders. Keyword lists live in a `detector.yaml` so they can be tuned
without code changes.

For each candidate, a starter rule is appended to `suggested_rules.yaml`, with `name` from
the sender display name, `match.from` from the sender domain, `match.subject` from the
subject, and `country`/`currency` guessed (`.ca` domain or "CRA"/"registered charity" →
CA/CAD; else US/USD), marked `# REVIEW`.

## Error handling

| Situation | Behaviour |
|---|---|
| Invalid `rules.yaml` / `accounts.yaml` | Stop before processing; message names file and rule |
| Missing mbox file | Stop with the path |
| Malformed message | Skip, log to `run_log.txt`, count on Summary |
| PDF without text layer | Save attachment; if amount not otherwise found → NEEDS REVIEW with note |
| Multiple rules match | First wins; conflict logged |
| Edge PDF render fails | Copy `.html` into accountant pack; note on row |
| Output folder for year exists | Deleted and rebuilt (it is fully regenerated output) |

## Testing

- pytest unit tests per module using small hand-made `.eml` / `.mbox` fixtures:
  header decoding (incl. `unknown-8bit`), year filtering, Message-ID de-dup, rule
  validation errors, default and custom extraction, receipt vs confirmation totals, US $250
  flag, candidate detection, starter-rule generation, relative hyperlinks.
- Regression fixtures: one anonymized real email per rule (amounts, org, receipt # kept;
  personal name/address removed). Each fixture asserts the extracted fields.
- End-to-end test: fixture mbox → full output tree; assert totals, row count, and that every
  hyperlink target exists.
- Edge rendering is behind an interface so tests can substitute a fake renderer.

## Dependencies

Python 3.11+, `pyyaml`, `openpyxl`, `pypdf`, `pytest`. Microsoft Edge (preinstalled on
Windows 11). Standard library `mailbox`, `email`.

## Reuse from earlier attempts

Earlier attempts live in `archived_attempts/attempt1` and `archived_attempts/attempt2`.

| Old code | Reused as | Notes |
|---|---|---|
| `decode_str` (attempt2 version) | `mbox_reader` | Keep the `unknown-8bit` fallback |
| `get_text_body` | `mbox_reader` | Keep; split into `body_text` and `body_html` |
| `parse_date` | `mbox_reader` | Keep `parsedate_to_datetime` + regex fallback; return a `datetime` |
| `safe_name` | `evidence` | Keep as-is |
| Attachment walk loop (`extract_attachments.py`) | `mbox_reader` attachments | Keep ext map for unnamed parts; drop hard-coded paths |
| `get_best_body` + HTML wrapper (`build_2025_folder.py`) | `evidence` html + PDF source | Keep header block; add HTML escaping of plain-text fallback (already present) |
| `SUBJECT_KEYWORDS`, `BODY_KEYWORDS`, `RECEIPT_*`, `SOLICITATION_SUBJECTS`, `COMMERCIAL_SENDERS` | `detector.yaml` | Move from code to data |
| Organization names seen in attempt2 outputs | initial `rules.yaml` | Verified against 2025+ mail during implementation |

Not reused: hard-coded paths, `TARGETS` / `PDFS_TO_COPY` one-off lists, A/B/C/D
classification, Markdown and CSV writers, `AMOUNT_RE` "first 5 amounts" approach (replaced by
labelled-amount extraction).

## Accounts in scope

- `abbasirumael@gmail.com` and `itsrumzy@gmail.com`: exports present in the project folder.
- `rumaelabbasi@gmail.com`: out of scope for now (no donation receipts); left commented out
  in `accounts.yaml`.
