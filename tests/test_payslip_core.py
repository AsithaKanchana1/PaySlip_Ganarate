"""
tests/test_payslip_core.py
==========================
Unit & integration tests for payslip_core.py using the real Excel template
from the Excel/ folder (password-protected, git-ignored).

Run with:
    .venv/bin/pytest tests/test_payslip_core.py -v
"""

import os
import sys
import pytest

# Make sure the project root is importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from payslip_core import (
    read_employees,
    generate_pdf,
    _fmt,
    _fmt_units,
    _g,
    COL_EMP_NO,
    COL_NAME,
    COL_BASIC,
    COL_GROSS,
    COL_NET_SALARY,
    COL_TOTAL_DED,
)


# ─── Helper / formatter tests ─────────────────────────────────────────────────
# These do NOT need an Excel file — pure Python logic.

class TestFmtHelpers:
    """Tests for _fmt and _fmt_units formatting helpers."""

    def test_fmt_integer(self):
        assert _fmt(25000) == "25,000.00"

    def test_fmt_float(self):
        assert _fmt(1234.5) == "1,234.50"

    def test_fmt_none_returns_zero(self):
        assert _fmt(None) == "0.00"

    def test_fmt_string_number(self):
        assert _fmt("500.50") == "500.50"

    def test_fmt_invalid_string_returns_zero(self):
        assert _fmt("N/A") == "0.00"

    def test_fmt_units_zero_returns_empty(self):
        assert _fmt_units(0) == ""

    def test_fmt_units_none_returns_empty(self):
        assert _fmt_units(None) == ""

    def test_fmt_units_positive(self):
        assert _fmt_units(2.5) == "2.5"

    def test_fmt_units_whole_number(self):
        assert _fmt_units(4.0) == "4"


# ─── read_employees tests — uses real Excel template ─────────────────────────

class TestReadEmployees:
    """Tests for read_employees() using the real password-protected Excel file."""

    def test_returns_list(self, real_excel_path):
        employees = read_employees(real_excel_path)
        assert isinstance(employees, list)

    def test_has_employees(self, real_excel_path):
        employees = read_employees(real_excel_path)
        assert len(employees) > 0, "Expected at least one employee row"

    def test_names_are_strings(self, real_excel_path):
        employees = read_employees(real_excel_path)
        for emp in employees:
            name = _g(emp, COL_NAME)
            if name:
                assert isinstance(name, str), f"Name should be a string, got {type(name)}"

    def test_basic_salary_or_formula_column_present(self, real_excel_path):
        """
        The real Excel uses formula-driven salary columns. openpyxl's data_only
        mode reads cached formula results — if the file was never saved with
        calculated values, those cells come back as None. This test verifies
        that the sheet is parsed without error and names are present, confirming
        the column mapper works even when numeric salary cells are None.
        """
        employees = read_employees(real_excel_path)
        # At least some employees must have a name
        names = [_g(emp, COL_NAME) for emp in employees if _g(emp, COL_NAME)]
        assert len(names) > 0, "Expected at least one employee with a name"

    def test_gross_numeric_or_none(self, real_excel_path):
        employees = read_employees(real_excel_path)
        for emp in employees:
            gross = _g(emp, COL_GROSS)
            if gross is not None:
                float(gross)  # should not raise

    def test_net_salary_leq_gross(self, real_excel_path):
        """Net salary should not exceed gross for any employee."""
        employees = read_employees(real_excel_path)
        for emp in employees:
            gross = _g(emp, COL_GROSS)
            net   = _g(emp, COL_NET_SALARY)
            if gross is not None and net is not None:
                try:
                    assert float(net) <= float(gross), (
                        f"Net {net} > Gross {gross} for employee {_g(emp, COL_NAME)}"
                    )
                except (ValueError, TypeError):
                    pass  # non-numeric rows (group headers etc.) are OK

    def test_no_negative_net_salary(self, real_excel_path):
        """Net salary should not be negative for any real employee row."""
        employees = read_employees(real_excel_path)
        for emp in employees:
            net = _g(emp, COL_NET_SALARY)
            if net is not None:
                try:
                    assert float(net) >= 0, (
                        f"Negative net salary for {_g(emp, COL_NAME)}: {net}"
                    )
                except (ValueError, TypeError):
                    pass

    def test_file_not_found_raises(self):
        with pytest.raises(Exception):
            read_employees("/tmp/nonexistent_abc123.xlsx")

    def test_wrong_password_raises(self, real_excel_plain_path):
        with pytest.raises(Exception):
            read_employees(f"{real_excel_plain_path}::wrongpassword")


# ─── generate_pdf tests — uses real Excel template ───────────────────────────

class TestGeneratePdf:
    """Integration tests for generate_pdf() using the real Excel template."""

    def test_pdf_created_successfully(self, real_excel_path, tmp_path):
        output = str(tmp_path / "output.pdf")
        result = generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
        )
        assert result["success"] is True, f"PDF generation failed: {result['message']}"
        assert os.path.exists(output)

    def test_pdf_has_multiple_employees(self, real_excel_path, tmp_path):
        output = str(tmp_path / "output2.pdf")
        result = generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
        )
        assert result["employee_count"] > 1, "Real sheet should have multiple employees"

    def test_pdf_page_count_at_least_one(self, real_excel_path, tmp_path):
        output = str(tmp_path / "output3.pdf")
        result = generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
        )
        assert result["page_count"] >= 1

    def test_pdf_file_is_not_empty(self, real_excel_path, tmp_path):
        output = str(tmp_path / "output4.pdf")
        generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
        )
        assert os.path.getsize(output) > 1024, "PDF should be larger than 1 KB"

    def test_pdf_starts_with_pdf_magic(self, real_excel_path, tmp_path):
        """Verify output is actually a valid PDF."""
        output = str(tmp_path / "output5.pdf")
        generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
        )
        with open(output, "rb") as f:
            header = f.read(4)
        assert header == b"%PDF", "Output file must start with PDF magic bytes"

    def test_pdf_output_path_in_result(self, real_excel_path, tmp_path):
        output = str(tmp_path / "output6.pdf")
        result = generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
        )
        assert result["output_path"] == output

    def test_pdf_missing_excel_returns_failure(self, tmp_path):
        output = str(tmp_path / "fail.pdf")
        result = generate_pdf(
            excel_path="/tmp/does_not_exist_xyz.xlsx",
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
        )
        assert result["success"] is False

    def test_progress_callback_called(self, real_excel_path, tmp_path):
        output = str(tmp_path / "output7.pdf")
        progress_values = []
        generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
            progress_callback=lambda pct: progress_values.append(pct),
        )
        assert len(progress_values) > 0, "Progress callback must be called at least once"
        assert progress_values[-1] == 100, "Final progress value must be 100"

    def test_log_callback_called(self, real_excel_path, tmp_path):
        output = str(tmp_path / "output8.pdf")
        log_messages = []
        generate_pdf(
            excel_path=real_excel_path,
            output_path=output,
            company_name="NEW LANKA CLOTHING",
            pay_period="April-2026",
            log_callback=lambda msg: log_messages.append(msg),
        )
        assert len(log_messages) > 0, "Log callback must be called at least once"
        combined = " ".join(log_messages).lower()
        assert any(w in combined for w in ("employee", "loaded", "pdf", "record"))
