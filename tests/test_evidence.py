from decimal import Decimal

from donation_tracker.evidence import email_to_html, org_slug, safe_name, save_evidence, unique_path
from donation_tracker.links import gmail_url
from donation_tracker.models import OK, Attachment, Donation
from tests.helpers import email_record, rule_record


def test_gmail_url():
    assert gmail_url("abbasirumael@gmail.com", "<abc+1@mail.x.com>") == (
        "https://mail.google.com/mail/u/abbasirumael@gmail.com/#search/rfc822msgid:abc%2B1%40mail.x.com"
    )
    assert gmail_url("a@gmail.com", "<no-id-123>") == ""


def test_safe_name_and_slug():
    assert safe_name('a<b>:c"d/e\\f|g?h*i') == "a_b__c_d_e_f_g_h_i"
    assert org_slug("Muslim Hands Canada") == "Muslim_Hands_Canada"


def test_unique_path(tmp_path):
    (tmp_path / "x").mkdir()
    (tmp_path / "x_2").mkdir()
    assert unique_path(tmp_path / "x") == tmp_path / "x_3"
    (tmp_path / "r.pdf").write_bytes(b"")
    assert unique_path(tmp_path / "r.pdf") == tmp_path / "r_2.pdf"


def test_email_to_html_escapes_plain_text_and_has_header():
    html = email_to_html(email_record(subject="Tax <receipt>", body_text="1 < 2"))
    assert "Tax &lt;receipt&gt;" in html
    assert "1 &lt; 2" in html
    assert "abbasirumael@gmail.com" in html
    assert "<meta charset=\"utf-8\">" in html


def test_email_to_html_uses_html_body_when_present():
    html = email_to_html(email_record(body_html="<p>Gift <b>$5</b></p>", body_text="Gift $5"))
    assert "<p>Gift <b>$5</b></p>" in html


def test_save_evidence_writes_files_and_sets_paths(tmp_path):
    atts = [
        Attachment("RD-US1939527-$60 -2026-04-21.pdf", "application/pdf", b"%PDF-1"),
        Attachment("logo.png", "image/png", b"PNG"),
    ]
    donation = Donation(email_record(attachments=atts, raw=b"RAW"), rule_record(), Decimal("60"), None, None, OK)
    save_evidence(donation, tmp_path / "evidence")
    folder = tmp_path / "evidence" / "2025-03-14_SickKids_Foundation"
    assert donation.email_eml_path == folder / "email.eml"
    assert donation.email_eml_path.read_bytes() == b"RAW"
    assert donation.email_html_path == folder / "email.html"
    assert donation.receipt_paths == [folder / "RD-US1939527-$60 -2026-04-21.pdf"]
    assert (folder / "logo.png").read_bytes() == b"PNG"

    second = Donation(email_record(raw=b"R2"), rule_record(), Decimal("1"), None, None, OK)
    save_evidence(second, tmp_path / "evidence")
    assert second.email_eml_path.parent.name == "2025-03-14_SickKids_Foundation_2"
