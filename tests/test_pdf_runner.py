import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pdf_runner import pdf_error_details


def test_an_encrypted_pdf_is_told_to_upload_an_unlocked_copy(tmp_path):
    # inspect_pdf stops a password-protected file with a ValueError, before any manifest exists
    error = ValueError("Encrypted PDFs require an explicit password-handling route.")

    ok, message, steps = pdf_error_details(Path(tmp_path) / "missing.remediation.json", error)

    assert ok is False
    assert "Encrypted" in message
    assert any("unlocked copy" in step for step in steps)
