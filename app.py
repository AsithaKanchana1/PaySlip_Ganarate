"""
app.py  —  New Lanka Clothing · Pay Slip Generator
====================================================
Desktop GUI application — runs on Windows, Linux, and macOS.
Non-technical users can browse for the Excel file, pick the month,
and click Generate. The PDF opens automatically when done.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import subprocess
import sys
import os
import json
import webbrowser
from datetime import datetime

# Import our core generation module
try:
    from payslip_core import generate_pdf
except ImportError:
    # When running as a frozen exe, both files are in the same folder
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from payslip_core import generate_pdf

# ─── App settings file (remembers last-used values) ──────────────────────────
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".payslip_settings.json")

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# ─── Theme ────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def load_settings() -> dict:
    try:
        with open(SETTINGS_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data: dict):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def open_file(path: str):
    """Open a file with the default OS application."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ─── Main Application Window ─────────────────────────────────────────────────

class PaySlipApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.settings = load_settings()
        now = datetime.now()

        self.title("Pay Slip Generator — New Lanka Clothing")
        self.geometry("680x800")
        self.minsize(600, 720)
        self.resizable(True, True)

        # App icon (if available)
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # State
        self._generating = False
        self._last_output = ""

        self._build_ui(now)
        self._build_developer_footer()
        self._load_saved_settings()

    # ── UI Builder ────────────────────────────────────────────────────────────

    def _build_ui(self, now):
        # ── Top header ────────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color=("#1a2744", "#0d1b3e"), corner_radius=0)
        header.pack(fill="x", padx=0, pady=0)

        ctk.CTkLabel(
            header,
            text="🧾  Pay Slip Generator",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="white",
        ).pack(side="left", padx=24, pady=18)

        ctk.CTkLabel(
            header,
            text="New Lanka Clothing",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8",
        ).pack(side="right", padx=24, pady=18)

        # ── Scrollable main body ──────────────────────────────────────────────
        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(16, 8))

        # ── Card 1 — Company Settings ─────────────────────────────────────────
        self._card(body, "⚙️  Company Settings")

        # Company name
        ctk.CTkLabel(body, text="Company Name", font=ctk.CTkFont(size=12),
                     text_color="#94a3b8").pack(anchor="w", padx=4, pady=(4, 2))
        self.company_var = ctk.StringVar(value="NEW LANKA CLOTHING")
        ctk.CTkEntry(body, textvariable=self.company_var,
                     font=ctk.CTkFont(size=13), height=40,
                     placeholder_text="Enter company name").pack(
            fill="x", padx=4, pady=(0, 12))

        # Pay period row (month + year side by side)
        ctk.CTkLabel(body, text="Pay Period", font=ctk.CTkFont(size=12),
                     text_color="#94a3b8").pack(anchor="w", padx=4, pady=(0, 2))

        period_row = ctk.CTkFrame(body, fg_color="transparent")
        period_row.pack(fill="x", padx=4, pady=(0, 16))

        self.month_var = ctk.StringVar(value=MONTHS[now.month - 1])
        self.month_menu = ctk.CTkOptionMenu(
            period_row, variable=self.month_var,
            values=MONTHS,
            font=ctk.CTkFont(size=13),
            width=180, height=40,
            fg_color=("#2563eb", "#1d4ed8"),
            button_color=("#1d4ed8", "#1e40af"),
        )
        self.month_menu.pack(side="left", padx=(0, 8))

        self.year_var = ctk.StringVar(value=str(now.year))
        self.year_entry = ctk.CTkEntry(
            period_row, textvariable=self.year_var,
            font=ctk.CTkFont(size=13), width=100, height=40,
            placeholder_text="Year")
        self.year_entry.pack(side="left")

        # Preview label
        self.period_preview = ctk.CTkLabel(
            period_row, text="", font=ctk.CTkFont(size=12),
            text_color="#60a5fa")
        self.period_preview.pack(side="left", padx=12)

        self.month_var.trace_add("write", self._update_period_preview)
        self.year_var.trace_add("write", self._update_period_preview)
        self._update_period_preview()

        # ── Card 2 — Excel Data File ──────────────────────────────────────────
        self._card(body, "📊  Excel Data File")

        ctk.CTkLabel(body,
                     text="Select the monthly salary Excel file (*.xlsx)",
                     font=ctk.CTkFont(size=12), text_color="#94a3b8").pack(
            anchor="w", padx=4, pady=(4, 6))

        file_row = ctk.CTkFrame(body, fg_color="transparent")
        file_row.pack(fill="x", padx=4, pady=(0, 4))

        self.excel_path_var = ctk.StringVar(value="")
        self.excel_entry = ctk.CTkEntry(
            file_row, textvariable=self.excel_path_var,
            font=ctk.CTkFont(size=11), height=40,
            placeholder_text="No file selected — click Browse to choose…",
            state="readonly")
        self.excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            file_row, text="📁  Browse",
            font=ctk.CTkFont(size=13, weight="bold"),
            width=100, height=40,
            command=self._browse_excel,
        ).pack(side="right")

        self.excel_status = ctk.CTkLabel(body, text="", font=ctk.CTkFont(size=11))
        self.excel_status.pack(anchor="w", padx=4, pady=(2, 12))

        # ── Card 3 — Output Location ──────────────────────────────────────────
        self._card(body, "📁  Output PDF Location")

        ctk.CTkLabel(body,
                     text="Where to save the generated PDF file",
                     font=ctk.CTkFont(size=12), text_color="#94a3b8").pack(
            anchor="w", padx=4, pady=(4, 6))

        out_row = ctk.CTkFrame(body, fg_color="transparent")
        out_row.pack(fill="x", padx=4, pady=(0, 16))

        self.output_path_var = ctk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Desktop", "PaySlips_Output.pdf"))
        self.output_entry = ctk.CTkEntry(
            out_row, textvariable=self.output_path_var,
            font=ctk.CTkFont(size=11), height=40)
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            out_row, text="📂  Choose",
            font=ctk.CTkFont(size=13, weight="bold"),
            width=100, height=40,
            command=self._browse_output,
        ).pack(side="right")

        # ── Generate button ────────────────────────────────────────────────────
        self.generate_btn = ctk.CTkButton(
            body,
            text="🖨️   Generate Pay Slips PDF",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=56,
            fg_color=("#16a34a", "#15803d"),
            hover_color=("#15803d", "#166534"),
            command=self._start_generation,
        )
        self.generate_btn.pack(fill="x", padx=4, pady=(8, 4))

        # ── Progress bar ───────────────────────────────────────────────────────
        self.progress_var = ctk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(body, variable=self.progress_var,
                                               height=10)
        self.progress_bar.pack(fill="x", padx=4, pady=(4, 2))
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            body, text="Ready — fill in the details above and click Generate",
            font=ctk.CTkFont(size=12), text_color="#94a3b8")
        self.status_label.pack(anchor="w", padx=4, pady=(2, 8))

        # ── Log box ────────────────────────────────────────────────────────────
        self._card(body, "📋  Activity Log")
        self.log_box = ctk.CTkTextbox(
            body, height=160,
            font=ctk.CTkFont(family="Courier New", size=11),
            fg_color=("#0f172a", "#020817"),
            text_color="#94a3b8",
            state="disabled",
        )
        self.log_box.pack(fill="x", padx=4, pady=(0, 8))

        # ── Open PDF button (hidden until generated) ──────────────────────────
        self.open_pdf_btn = ctk.CTkButton(
            body,
            text="📄  Open PDF",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=46,
            fg_color=("#2563eb", "#1d4ed8"),
            hover_color=("#1d4ed8", "#1e40af"),
            command=self._open_pdf,
            state="disabled",
        )
        self.open_pdf_btn.pack(fill="x", padx=4, pady=(0, 20))

    # ── Developer footer ──────────────────────────────────────────────────────

    def _build_developer_footer(self):
        """Pinned footer bar — always visible at the very bottom of the window."""
        footer = ctk.CTkFrame(
            self,
            fg_color=("#0d1220", "#080e1a"),
            corner_radius=0,
            height=32,
        )
        footer.pack(fill="x", side="bottom", padx=0, pady=0)
        footer.pack_propagate(False)   # keep fixed height

        # Thin separator line at the top of the footer
        ctk.CTkFrame(
            footer, height=1,
            fg_color=("#1e3a8a", "#1e2d5a"),
            corner_radius=0,
        ).pack(fill="x", padx=0, pady=0)

        # Clickable developer credit label
        dev_label = ctk.CTkLabel(
            footer,
            text="Asitha  ❤️  Kanchana",
            font=ctk.CTkFont(family="Helvetica", size=11),
            text_color=("#4a6fa5", "#5b82c0"),
            cursor="hand2",
        )
        dev_label.pack(expand=True)

        # Hover effect — brighten on mouse-over
        def _on_enter(e):
            dev_label.configure(text_color=("#93c5fd", "#93c5fd"))

        def _on_leave(e):
            dev_label.configure(text_color=("#4a6fa5", "#5b82c0"))

        def _open_github(e):
            self._show_developer_details()

        dev_label.bind("<Enter>",    _on_enter)
        dev_label.bind("<Leave>",    _on_leave)
        dev_label.bind("<Button-1>", _open_github)

    def _show_developer_details(self):
        """Show a small modal with developer social links."""
        links = [
            ("Website", "https://asitha.site"),
            ("YouTube", "https://youtube.com/@asi_solution"),
            ("LinkedIn", "https://linkedin.com/in/asitha-kanchana-35aa531a8/"),
            ("Facebook", "https://facebook.com/asithakanchana01"),
            ("Instagram", "https://instagram.com/asithakanchana01"),
            ("Stack Overflow", "https://stackoverflow.com/users/18616349/asitha-kanchana"),
            ("WhatsApp", "https://wa.me/94701336364"),
            ("Buy Me a Coffee", "https://buymeacoffee.com/asitha"),
            ("GitHub", "https://github.com/AsithaKanchana1"),
            ("GitLab", "https://gitlab.com/"),
            ("ORCID", "https://orcid.org/0009-0003-2200-4467"),
        ]

        try:
            dlg = ctk.CTkToplevel(self)
        except Exception:
            # Fallback if CTkToplevel is not available
            from tkinter import Toplevel
            dlg = Toplevel(self)

        dlg.title("Developer — Asitha Kanchana")
        dlg.geometry("520x420")
        dlg.resizable(False, False)
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass

        body = ctk.CTkFrame(dlg, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=12)

        ctk.CTkLabel(body, text="Developer & Social Links",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="#93c5fd").pack(anchor="w", pady=(2, 12))

        for label, url in links:
            lbl = ctk.CTkLabel(body, text=f"{label}: {url}", anchor="w",
                               cursor="hand2", wraplength=480,
                               text_color="#cbd5e1")
            lbl.pack(fill="x", pady=6)

            def _open(u=url):
                webbrowser.open(u)

            lbl.bind("<Button-1>", lambda e, u=url: _open(u))

        ctk.CTkButton(body, text="Close", width=120,
                      command=lambda: dlg.destroy()).pack(pady=14)

    def _card(self, parent, title: str):
        """Draw a section header."""
        ctk.CTkLabel(
            parent, text=title,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#1e3a8a", "#93c5fd"),
        ).pack(anchor="w", padx=4, pady=(16, 4))
        ctk.CTkFrame(parent, height=1,
                     fg_color=("#cbd5e1", "#1e3a8a")).pack(fill="x", padx=4, pady=(0, 8))

    # ── Settings persistence ──────────────────────────────────────────────────

    def _load_saved_settings(self):
        s = self.settings
        if "company" in s:
            self.company_var.set(s["company"])
        if "month" in s and s["month"] in MONTHS:
            self.month_var.set(s["month"])
        if "year" in s:
            self.year_var.set(s["year"])
        if "excel_path" in s and os.path.exists(s["excel_path"]):
            self._set_excel_path(s["excel_path"])
        if "output_path" in s:
            self.output_path_var.set(s["output_path"])

    def _save_current_settings(self):
        save_settings({
            "company":     self.company_var.get(),
            "month":       self.month_var.get(),
            "year":        self.year_var.get(),
            "excel_path":  self.excel_path_var.get(),
            "output_path": self.output_path_var.get(),
        })

    # ── Period preview ────────────────────────────────────────────────────────

    def _update_period_preview(self, *_):
        month = self.month_var.get()
        year  = self.year_var.get()
        self.period_preview.configure(text=f"→  {month}-{year}")

    def _get_pay_period(self) -> str:
        return f"{self.month_var.get()}-{self.year_var.get()}"

    # ── File pickers ──────────────────────────────────────────────────────────

    def _browse_excel(self):
        initial = os.path.dirname(self.excel_path_var.get()) or os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Select salary Excel file",
            initialdir=initial,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self._set_excel_path(path)

    def _set_excel_path(self, path: str):
        self.excel_path_var.set(path)
        size_kb = os.path.getsize(path) / 1024
        self.excel_status.configure(
            text=f"✅  {os.path.basename(path)}  ({size_kb:.1f} KB)",
            text_color=("#16a34a", "#4ade80"))

        # Auto-set output path next to the Excel file
        out_dir  = os.path.dirname(path)
        out_file = os.path.join(out_dir, "PaySlips_Output.pdf")
        self.output_path_var.set(out_file)

    def _browse_output(self):
        initial = os.path.dirname(self.output_path_var.get()) or os.path.expanduser("~")
        path = filedialog.asksaveasfilename(
            title="Save PDF as …",
            initialdir=initial,
            initialfile="PaySlips_Output.pdf",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self.output_path_var.set(path)

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _log(self, message: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ── Generation ────────────────────────────────────────────────────────────

    def _start_generation(self):
        if self._generating:
            return

        # Validate inputs
        excel = self.excel_path_var.get().strip()
        output = self.output_path_var.get().strip()
        company = self.company_var.get().strip()
        year = self.year_var.get().strip()

        if not excel:
            messagebox.showerror("Missing File",
                                 "Please browse and select the Excel salary file first.")
            return
        if not os.path.exists(excel):
            messagebox.showerror("File Not Found",
                                 f"Excel file not found:\n{excel}")
            return
        if not output:
            messagebox.showerror("Missing Output",
                                 "Please choose where to save the PDF.")
            return
        if not company:
            messagebox.showerror("Missing Company",
                                 "Please enter the company name.")
            return
        if not year.isdigit() or len(year) != 4:
            messagebox.showerror("Invalid Year",
                                 "Please enter a valid 4-digit year (e.g. 2026).")
            return

        pay_period = self._get_pay_period()
        self._save_current_settings()
        self._run_in_thread(excel, output, company, pay_period)

    def _run_in_thread(self, excel, output, company, pay_period):
        self._generating = True
        self.generate_btn.configure(state="disabled", text="⏳  Generating …")
        self.open_pdf_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Generating pay slips …", text_color="#fbbf24")
        self._clear_log()
        self._log(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        self._log(f"  Company   : {company}")
        self._log(f"  Period    : {pay_period}")
        self._log(f"  Excel     : {os.path.basename(excel)}")
        self._log(f"  Output    : {output}")
        self._log(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        def worker():
            result = generate_pdf(
                excel_path=excel,
                output_path=output,
                company_name=company,
                pay_period=pay_period,
                progress_callback=lambda pct: self.after(0, self._set_progress, pct),
                log_callback=lambda msg: self.after(0, self._log, msg),
            )
            self.after(0, self._on_done, result)

        threading.Thread(target=worker, daemon=True).start()

    def _set_progress(self, pct: int):
        self.progress_bar.set(pct / 100)

    def _on_done(self, result: dict):
        self._generating = False
        self.generate_btn.configure(state="normal",
                                    text="🖨️   Generate Pay Slips PDF")

        if result["success"]:
            self._last_output = result["output_path"]
            self.status_label.configure(
                text=f"✅  Done!  {result['employee_count']} slips  ·  "
                     f"{result['page_count']} page(s)",
                text_color=("#16a34a", "#4ade80"))
            self.open_pdf_btn.configure(state="normal")
            self._log("")
            self._log(f"🎉  {result['message']}")
            messagebox.showinfo(
                "Success",
                f"Pay slips generated successfully!\n\n"
                f"👥  Employees : {result['employee_count']}\n"
                f"📄  Pages     : {result['page_count']}\n\n"
                f"📁  Saved to:\n{result['output_path']}",
            )
        else:
            self.status_label.configure(
                text=f"❌  Error — see log below",
                text_color=("#dc2626", "#f87171"))
            self._log(f"\n❌  FAILED: {result['message']}")
            messagebox.showerror("Generation Failed", result["message"])

    def _open_pdf(self):
        if self._last_output and os.path.exists(self._last_output):
            open_file(self._last_output)
        else:
            messagebox.showwarning("File Not Found",
                                   "Could not find the generated PDF file.")


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = PaySlipApp()
    app.mainloop()
