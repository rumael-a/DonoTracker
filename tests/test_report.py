from decimal import Decimal

from openpyxl import load_workbook

from donation_tracker.models import NEEDS_REVIEW, OK, Donation
from donation_tracker.report import row_from_donation, summarize, write_workbook
from tests.helpers import email_record, rule_record


def donation(amount, *, role="receipt", status=OK, country="CA", currency="CAD", name="SickKids Foundation"):
    rule = rule_record(role=role, country=country, currency=currency, name=name)
    d = Donation(email_record(), rule, Decimal(amount) if amount else None, "R1", None, status, [])
    d.gmail_url = "https://mail.google.com/x"
    return d


def test_summarize_counts_only_ok_receipts():
    ds = [
        donation("100"),
        donation("50", role="confirmation"),
        donation(None, status=NEEDS_REVIEW),
        donation("20", country="US", currency="USD", name="Red Cross"),
        donation("5"),
    ]
    s = summarize(ds, needs_rule=3, skipped=1)
    assert s.totals == {("CA", "CAD"): Decimal("105"), ("US", "USD"): Decimal("20")}
    assert s.by_org == {("SickKids Foundation", "CAD"): Decimal("105"), ("Red Cross", "USD"): Decimal("20")}
    assert (s.needs_review, s.needs_rule, s.skipped) == (1, 3, 1)


def test_write_workbook_sheets_rows_and_relative_links(tmp_path):
    out = tmp_path / "2025"
    folder = out / "evidence" / "2025-03-14_SickKids_Foundation"
    folder.mkdir(parents=True)
    d = donation("100")
    d.email_html_path = folder / "email.html"
    d.receipt_paths = [folder / "a.pdf", folder / "b.pdf"]
    review = donation(None, status=NEEDS_REVIEW)
    review.notes = ["Amount not found"]
    rows = [row_from_donation(d), row_from_donation(review)]
    candidate = email_record(sender="New <x@new.org>", subject="Thanks for your gift")
    write_workbook(out / "donations_2025.xlsx", rows, summarize([d, review], 1, 0), [candidate])

    wb = load_workbook(out / "donations_2025.xlsx")
    assert wb.sheetnames == ["Summary", "Donations", "Needs a rule"]
    ws = wb["Donations"]
    header = [c.value for c in ws[1]]
    assert header == ["Date", "Organization", "Country", "Currency", "Amount", "Role", "Receipt #",
                      "Charity ID", "Account", "Status", "Gmail", "Email file", "Receipt file", "Notes"]
    row = {h: c for h, c in zip(header, ws[2])}
    assert row["Amount"].value == 100
    assert row["Gmail"].hyperlink.target == "https://mail.google.com/x"
    assert row["Email file"].hyperlink.target == "evidence/2025-03-14_SickKids_Foundation/email.html"
    assert row["Receipt file"].hyperlink.target == "evidence/2025-03-14_SickKids_Foundation/a.pdf"
    assert "b.pdf" in row["Notes"].value
    review_row = {h: c for h, c in zip(header, ws[3])}
    assert review_row["Status"].value == "NEEDS REVIEW"
    assert review_row["Status"].fill.fgColor.rgb.endswith("FFFF00")
    summary_values = [c.value for r in wb["Summary"].iter_rows() for c in r if c.value is not None]
    assert "CA / CAD" in summary_values and 100 in summary_values
    rules_sheet = wb["Needs a rule"]
    assert rules_sheet["D2"].value == "Thanks for your gift"


def test_workbook_without_needs_rule_sheet(tmp_path):
    write_workbook(tmp_path / "s.xlsx", [], summarize([], 0, 0), None)
    assert load_workbook(tmp_path / "s.xlsx").sheetnames == ["Summary", "Donations"]
