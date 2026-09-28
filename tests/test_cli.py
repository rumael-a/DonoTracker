from pathlib import Path

from donation_tracker import cli
from tests.helpers import make_email, make_mbox
from tests.test_accountant import FakeRenderer

ROOT = Path(__file__).resolve().parents[1]


def setup_files(tmp_path):
    make_mbox(tmp_path / "a.mbox", [
        make_email(sender="Muslim Hands <mail@muslimhands.ca>", subject="Your Receipt from Muslim Hands",
                   text="Donation Amount: $50.00", date="Fri, 17 Oct 2025 10:00:00 -0400"),
    ])
    (tmp_path / "accounts.yaml").write_text(
        "accounts:\n  - email: abbasirumael@gmail.com\n    mbox: a.mbox\n", encoding="utf-8")
    (tmp_path / "rules.yaml").write_text(
        "- name: Muslim Hands Canada\n  country: CA\n  currency: CAD\n  match:\n    from: muslimhands.ca\n",
        encoding="utf-8")
    return [
        "--accounts", str(tmp_path / "accounts.yaml"),
        "--rules", str(tmp_path / "rules.yaml"),
        "--detector", str(ROOT / "detector.yaml"),
    ]


def test_run_command(tmp_path, capsys, monkeypatch):
    common = setup_files(tmp_path)
    monkeypatch.setattr(cli, "EdgeRenderer", FakeRenderer)
    code = cli.main(["run", "--year", "2025", "--out", str(tmp_path / "output"), *common])
    assert code == 0
    out = capsys.readouterr().out
    assert "CA / CAD: 50.00" in out
    assert "Orgs with confirmations but no counted receipt: (none)" in out
    assert (tmp_path / "output" / "2025" / "donations_2025.xlsx").exists()


def test_test_rule_command_prints_matches_without_writing(tmp_path, capsys):
    common = setup_files(tmp_path)
    code = cli.main(["test-rule", "Muslim Hands Canada", "--year", "2025", *common])
    assert code == 0
    out = capsys.readouterr().out
    assert "Your Receipt from Muslim Hands" in out
    assert "50.00" in out
    assert "1 email(s) matched" in out
    assert not (tmp_path / "output").exists()


def test_unknown_rule_and_config_errors(tmp_path, capsys):
    common = setup_files(tmp_path)
    assert cli.main(["test-rule", "Nope", "--year", "2025", *common]) == 2
    (tmp_path / "rules.yaml").write_text("- name: X\n  country: FR\n", encoding="utf-8")
    assert cli.main(["test-rule", "X", "--year", "2025", *common]) == 2
    assert "Config error" in capsys.readouterr().err
