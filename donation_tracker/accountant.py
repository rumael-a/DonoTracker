"""Build the self-contained for_accountant/ folder."""
from __future__ import annotations

import shutil
from pathlib import Path

from .evidence import org_slug, unique_path
from .models import Donation
from .report import Summary, row_from_donation, write_workbook

RENDER_FAILED = "PDF render failed; email included as HTML"


def build_pack(donations: list[Donation], pack_dir: Path, renderer, year: int, summary: Summary,
               log: list[str]) -> list[Path]:
    written: list[Path] = []
    rows = []
    for d in donations:
        if not d.is_countable:
            continue
        folder = pack_dir / d.rule.country
        folder.mkdir(parents=True, exist_ok=True)
        prefix = f"{d.email.date:%Y-%m-%d}_{org_slug(d.rule.name)}_{d.rule.currency}_{d.amount:.2f}"
        notes = []
        email_file = unique_path(folder / f"{prefix}_email.pdf")
        if not renderer.render(d.email_html_path, email_file):
            email_file = unique_path(folder / f"{prefix}_email.html")
            shutil.copy2(d.email_html_path, email_file)
            notes.append(RENDER_FAILED)
            log.append(f"{RENDER_FAILED}: {d.email_html_path}")
        written.append(email_file)
        receipts = []
        for source in d.receipt_paths:
            target = unique_path(folder / f"{prefix}_receipt{source.suffix.lower() or '.pdf'}")
            shutil.copy2(source, target)
            receipts.append(target)
            written.append(target)
        rows.append(row_from_donation(d, email_file=email_file, receipt_files=receipts, extra_notes=notes))
    write_workbook(pack_dir / f"donations_{year}_summary.xlsx", rows, summary, None)
    return written
