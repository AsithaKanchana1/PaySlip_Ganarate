"""
conftest.py — shared pytest fixtures for PaySlip_Ganarate tests.

Uses the real-world Excel template from the Excel/ folder (password: 1111).
The Excel/ folder is git-ignored to protect salary data from being committed.
"""

import os
import pytest

# ── Path to the real-world template Excel (git-ignored, never committed) ───────
PROJECT_ROOT = os.path.join(os.path.dirname(__file__), "..")
EXCEL_DIR    = os.path.join(PROJECT_ROOT, "Excel")
EXCEL_PASSWORD = "1111"


def _find_excel_file():
    """Return the path of the first .xlsx/.xls file found in Excel/, or None."""
    if not os.path.isdir(EXCEL_DIR):
        return None
    for fname in sorted(os.listdir(EXCEL_DIR)):
        if fname.lower().endswith((".xlsx", ".xls")):
            return os.path.join(EXCEL_DIR, fname)
    return None


# ── Session-level fixture: path to real Excel (with password encoded) ──────────

@pytest.fixture(scope="session")
def real_excel_path():
    """
    Returns the 'path::password' string for the real Excel template.
    Skips the test if no Excel file is found in Excel/ (e.g. fresh CI clone).
    """
    path = _find_excel_file()
    if path is None:
        pytest.skip(
            "No Excel file found in Excel/ folder. "
            "Add a salary .xlsx file to run integration tests."
        )
    return f"{path}::{EXCEL_PASSWORD}"


@pytest.fixture(scope="session")
def real_excel_plain_path():
    """Returns just the file path (without password) for metadata checks."""
    path = _find_excel_file()
    if path is None:
        pytest.skip("No Excel file found in Excel/ folder.")
    return path
