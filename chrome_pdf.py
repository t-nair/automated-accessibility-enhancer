"""Export semantic HTML through Chrome's tagged-PDF print path."""

from __future__ import annotations

import subprocess
import tempfile
import time
import os
import shutil
from pathlib import Path


CHROME_CANDIDATES = (
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
    Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
    Path("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
)

CHROME_COMMANDS = ("google-chrome", "chromium", "chromium-browser", "microsoft-edge", "msedge")


def find_chrome() -> Path:
    configured = os.environ.get("PDF_PIPELINE_CHROME")
    if configured:
        candidate = Path(configured).expanduser()
        if candidate.is_file():
            return candidate
        raise RuntimeError("PDF_PIPELINE_CHROME does not point to a browser executable.")
    for candidate in CHROME_CANDIDATES:
        if candidate.exists():
            return candidate
    for command in CHROME_COMMANDS:
        executable = shutil.which(command)
        if executable:
            return Path(executable)
    raise RuntimeError(
        "Chrome, Chromium, or Edge is required to export the tagged companion PDF. "
        "Set PDF_PIPELINE_CHROME when it is not on PATH."
    )


def export_tagged_pdf(html_path: str | Path, pdf_path: str | Path) -> Path:
    source = Path(html_path).resolve()
    output = Path(pdf_path).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    command = [
        str(find_chrome()),
        "--headless=new",
        # The exporter opens only locally generated, escaped HTML. These flags
        # avoid Windows GPU/AppContainer failures in localhost development.
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--disable-background-networking",
        "--no-first-run",
        "--no-default-browser-check",
        "--no-pdf-header-footer",
        "--generate-pdf-document-outline",
        source.as_uri(),
    ]
    # Isolate headless printing from an already-open user browser profile.
    with tempfile.TemporaryDirectory(prefix="uw-pdf-browser-", ignore_cleanup_errors=True) as profile:
        staged = Path(profile) / "printed.pdf"
        command.insert(1, f"--user-data-dir={profile}")
        command.insert(2, f"--print-to-pdf={staged}")
        result = subprocess.run(command, capture_output=True, text=True, timeout=120, check=False,
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        # Edge can return before its print subprocess finishes writing the file.
        deadline = time.monotonic() + 15
        while result.returncode == 0 and time.monotonic() < deadline:
            if staged.exists():
                data = staged.read_bytes()
                if data.rstrip().endswith(b"%%EOF"):
                    output.write_bytes(data)
                    return output
            time.sleep(0.2)
        detail = (result.stderr or result.stdout or "No complete PDF was produced.").strip()
        raise RuntimeError(f"Tagged PDF export failed: {detail}")
