"""Render saved email HTML to PDF using headless Microsoft Edge (preinstalled on Windows 11)."""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]


def find_edge() -> str | None:
    for candidate in EDGE_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return shutil.which("msedge")


class EdgeRenderer:
    def __init__(self, exe: str | None = None):
        self.exe = find_edge() if exe is None else exe
        self._profile = tempfile.TemporaryDirectory(prefix="donations-edge-")

    def render(self, html_path: Path, pdf_path: Path) -> bool:
        if not self.exe:
            return False
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        command = [
            self.exe,
            "--headless",
            "--disable-gpu",
            "--no-first-run",
            f"--user-data-dir={self._profile.name}",
            "--no-pdf-header-footer",
            "--print-to-pdf-no-header",
            f"--print-to-pdf={pdf_path.resolve()}",
            html_path.resolve().as_uri(),
        ]
        try:
            subprocess.run(command, capture_output=True, timeout=90, check=False)
        except (OSError, subprocess.TimeoutExpired):
            return False
        return pdf_path.exists() and pdf_path.stat().st_size > 0
