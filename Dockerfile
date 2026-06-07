# ─────────────────────────────────────────────────────────────────────────────
# Dockerfile — PaySlip_Ganarate (CLI mode only, no GUI)
# ─────────────────────────────────────────────────────────────────────────────
# Build:
#   docker build -t payslip-generator .
#
# Run (generate pay slips from a local Excel file):
#   docker run --rm \
#     -v /path/to/your/Excel:/app/Excel \
#     -v /path/to/output:/app/output \
#     payslip-generator \
#     python generate_payslips.py
#
# Run tests inside the container:
#   docker run --rm payslip-generator pytest tests/ -v
# ─────────────────────────────────────────────────────────────────────────────

FROM python:3.11-slim

LABEL maintainer="Asitha Kanchana <https://github.com/AsithaKanchana1>"
LABEL description="Pay Slip Generator — New Lanka Clothing (CLI/headless)"

# ── System packages ────────────────────────────────────────────────────────────
# Fonts needed by reportlab for PDF text rendering
RUN apt-get update && apt-get install -y --no-install-recommends \
        fonts-liberation \
        libfreetype6 \
    && rm -rf /var/lib/apt/lists/*

# ── Working directory ──────────────────────────────────────────────────────────
WORKDIR /app

# ── Python dependencies ────────────────────────────────────────────────────────
# Copy only requirements first so Docker layer caching is efficient
COPY requirements.txt .
RUN pip install --no-cache-dir \
        openpyxl \
        reportlab \
        pillow \
        msoffcrypto-tool \
        xlrd \
        pytest \
        pytest-cov

# NOTE: customtkinter and pyinstaller are NOT installed — they require a display
#       and are not needed for headless/CLI operation.

# ── Copy source code ────────────────────────────────────────────────────────────
COPY payslip_core.py        .
COPY generate_payslips.py   .
COPY tests/                 ./tests/

# ── Excel input directory (bind-mount your file here at runtime) ───────────────
RUN mkdir -p /app/Excel /app/output

# ── Smoke test: verify imports work ───────────────────────────────────────────
RUN python -c "from payslip_core import generate_pdf, read_employees; print('✅ payslip_core imported OK')"

# ── Default command: run pytest ────────────────────────────────────────────────
CMD ["pytest", "tests/", "-v", "--tb=short"]
