"""Webview-based desktop launcher for the pay slip generator."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import uuid
from copy import deepcopy
from datetime import datetime
from typing import Any

try:
    import webview
except ImportError:  # pragma: no cover - handled at runtime
    webview = None

from payslip_core import generate_pdf

APP_NAME = "Pay Slip Generator"
APP_COMPANY = "New Lanka Clothing"
DEVELOPER_NAME = "Asitha Kanchana"
DEVELOPER_GITHUB = "https://github.com/AsithaKanchana1"
DEVELOPER_SOCIAL_LINKS = [
    {"label": "Website", "url": "https://asitha.site"},
    {"label": "YouTube", "url": "https://youtube.com/@asi_solution"},
    {"label": "LinkedIn", "url": "https://linkedin.com/in/asitha-kanchana-35aa531a8/"},
    {"label": "Facebook", "url": "https://facebook.com/asithakanchana01"},
    {"label": "Instagram", "url": "https://instagram.com/asithakanchana01"},
    {"label": "Stack Overflow", "url": "https://stackoverflow.com/users/18616349/asitha-kanchana"},
    {"label": "WhatsApp", "url": "https://wa.me/94701336364"},
    {"label": "Buy Me a Coffee", "url": "https://buymeacoffee.com/asitha"},
    {"label": "GitHub", "url": "https://github.com/AsithaKanchana1"},
    {"label": "GitLab", "url": "https://gitlab.com/"},
    {"label": "ORCID", "url": "https://orcid.org/0009-0003-2200-4467"},
]
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".payslip_settings.json")


def load_settings() -> dict:
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as handle:
            return json.load(handle)
    except Exception:
        return {}


def save_settings(data: dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)
    except Exception:
        pass


def open_file(path: str):
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


class PaySlipBridge:
    def __init__(self, base_dir: str):
        self.base_dir = base_dir
        self.settings = load_settings()
        self.window = None
        self._jobs: dict[str, dict[str, Any]] = {}
        self._jobs_lock = threading.Lock()

    def set_window(self, window):
        self.window = window

    def get_bootstrap(self) -> dict:
        return {
            "appName": APP_NAME,
            "company": APP_COMPANY,
            "developer": DEVELOPER_NAME,
            "developerGitHub": DEVELOPER_GITHUB,
            "socialLinks": DEVELOPER_SOCIAL_LINKS,
            "months": MONTHS,
            "settings": deepcopy(self.settings),
        }

    def save_user_settings(self, settings: dict) -> dict:
        self.settings.update(settings or {})
        save_settings(self.settings)
        return deepcopy(self.settings)

    def select_excel_file(self, current_path: str = "") -> dict:
        if not self.window:
            return {"ok": False, "message": "Window is not ready."}

        initial_dir = os.path.dirname(current_path) if current_path else os.path.expanduser("~")
        result = self.window.create_file_dialog(
            webview.FileDialog.OPEN,
            directory=initial_dir,
            allow_multiple=False,
            file_types=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")),
        )
        path = result[0] if result else ""
        return self._file_info(path)

    def select_output_file(self, current_path: str = "") -> dict:
        if not self.window:
            return {"ok": False, "message": "Window is not ready."}

        initial_dir = os.path.dirname(current_path) if current_path else os.path.expanduser("~")
        result = self.window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory=initial_dir,
            save_filename="PaySlips_Output.pdf",
            file_types=(("PDF files", "*.pdf"),),
        )
        path = result[0] if result else ""
        return self._file_info(path)

    def open_pdf(self, path: str) -> dict:
        if not path or not os.path.exists(path):
            return {"ok": False, "message": "PDF file not found."}
        open_file(path)
        return {"ok": True, "message": "Opened PDF."}

    def start_generation(self, payload: dict) -> dict:
        data = payload or {}
        excel_path = str(data.get("excelPath", "")).strip()
        output_path = str(data.get("outputPath", "")).strip()
        company_name = str(data.get("companyName", "")).strip()
        month = str(data.get("month", "")).strip()
        year = str(data.get("year", "")).strip()

        if not excel_path:
            return {"ok": False, "message": "Please select the Excel file first."}
        if not os.path.exists(excel_path):
            return {"ok": False, "message": f"Excel file not found: {excel_path}"}
        if not output_path:
            return {"ok": False, "message": "Please choose an output PDF path."}
        if not company_name:
            return {"ok": False, "message": "Please enter the company name."}
        if month not in MONTHS:
            return {"ok": False, "message": "Please select a valid month."}
        if not year.isdigit() or len(year) != 4:
            return {"ok": False, "message": "Please enter a valid 4-digit year."}

        pay_period = f"{month}-{year}"
        self.save_user_settings({
            "excel_path": excel_path,
            "output_path": output_path,
            "company": company_name,
            "month": month,
            "year": year,
        })

        job_id = uuid.uuid4().hex
        self._store_job(job_id, {
            "jobId": job_id,
            "status": "running",
            "progress": 0,
            "logs": [],
            "result": None,
            "message": "",
            "createdAt": datetime.utcnow().isoformat() + "Z",
            "updatedAt": datetime.utcnow().isoformat() + "Z",
        })

        def push_log(message: str):
            self._append_log(job_id, message)

        def push_progress(percent: int):
            self._update_job(job_id, progress=int(percent))

        def worker():
            try:
                result = generate_pdf(
                    excel_path=excel_path,
                    output_path=output_path,
                    company_name=company_name,
                    pay_period=pay_period,
                    progress_callback=push_progress,
                    log_callback=push_log,
                )
                if result.get("success"):
                    self._update_job(
                        job_id,
                        status="success",
                        progress=100,
                        message=result.get("message", "Completed."),
                        result=result,
                    )
                else:
                    push_log(result.get("message", "Generation failed."))
                    self._update_job(
                        job_id,
                        status="error",
                        message=result.get("message", "Generation failed."),
                        result=result,
                    )
            except Exception as exc:  # pragma: no cover - safety net
                self._append_log(job_id, f"❌  Error: {exc}")
                self._update_job(
                    job_id,
                    status="error",
                    message=str(exc),
                    result={"success": False, "message": str(exc)},
                )

        threading.Thread(target=worker, daemon=True).start()
        return {"ok": True, "jobId": job_id}

    def get_job_status(self, job_id: str) -> dict:
        with self._jobs_lock:
            job = deepcopy(self._jobs.get(job_id))
        if not job:
            return {"ok": False, "message": "Job not found."}
        job["ok"] = True
        return job

    def _file_info(self, path: str) -> dict:
        if not path:
            return {"ok": False, "path": "", "message": "No file selected."}
        try:
            size_kb = os.path.getsize(path) / 1024
            return {
                "ok": True,
                "path": path,
                "name": os.path.basename(path),
                "sizeKb": round(size_kb, 1),
                "message": "Selected.",
            }
        except Exception as exc:
            return {"ok": False, "path": path, "message": str(exc)}

    def _store_job(self, job_id: str, payload: dict):
        with self._jobs_lock:
            self._jobs[job_id] = payload

    def _update_job(self, job_id: str, **updates):
        with self._jobs_lock:
            job = self._jobs.get(job_id)
            if not job:
                return
            job.update(updates)
            job["updatedAt"] = datetime.utcnow().isoformat() + "Z"
            self._jobs[job_id] = job

    def _append_log(self, job_id: str, message: str):
        with self._jobs_lock:
            job = self._jobs.get(job_id)
            if not job:
                return
            job.setdefault("logs", []).append(message)
            job["updatedAt"] = datetime.utcnow().isoformat() + "Z"
            self._jobs[job_id] = job


def main():
    if webview is None:
        raise SystemExit(
            "pywebview is not installed. Install dependencies with: pip install -r requirements.txt"
        )

    base_dir = os.path.dirname(os.path.abspath(__file__))
    index_path = os.path.join(base_dir, "webview", "index.html")

    if not os.path.exists(index_path):
        raise SystemExit(f"Web UI not found: {index_path}")

    bridge = PaySlipBridge(base_dir)
    window = webview.create_window(
        APP_NAME,
        index_path,
        js_api=bridge,
        maximized=True,
        min_size=(1180, 760),
        resizable=True,
    )
    bridge.set_window(window)
    webview.start(debug=False)


if __name__ == "__main__":
    main()
