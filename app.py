"""
Automated Provincial Email Dispatcher (Email Blast BPS)
Main Desktop Application GUI built with Tkinter & TTK.
"""

import json
import os
import sys
import threading
import queue
import time
import tempfile
import webbrowser
import re
from typing import List, Dict, Any, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.excel_parser import BPSExcelParser
from core.attachment_engine import AttachmentEngine, get_file_size_formatted
from core.history_manager import HistoryManager, get_default_db_path
from core.mailer import (
    replace_template,
    test_smtp_connection,
    EmailDispatcher,
    check_smtp_sender_compatibility,
)


def get_config_path() -> str:
    """Get absolute path to config.json, compatible with both script and PyInstaller exe."""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "config.json")


class BPSDispatcherApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Automated Provincial Email Dispatcher - BPS RI")
        self.root.geometry("1100x750")
        self.root.minsize(960, 650)

        # Apply modern Windows look & style
        self._setup_styles()

        # Application State
        self.excel_path = tk.StringVar(value="")
        self.folder_path = tk.StringVar(value="")
        self.smtp_host = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port = tk.StringVar(value="587")
        self.smtp_username = tk.StringVar(value="")
        self.sender_email = tk.StringVar(value="")
        self.smtp_pass = tk.StringVar(value="")
        self.smtp_show_pass = tk.BooleanVar(value=False)

        # Backward compatibility alias
        self.smtp_user = self.smtp_username

        self.email_subject = tk.StringVar(
            value="[BPS RI] Penyampaian Dokumen Naskah Publikasi Statistik - BPS Provinsi {provinsi} ({kode})"
        )
        self.email_delay = tk.DoubleVar(value=1.0)
        self.smtp_timeout = tk.IntVar(value=45)
        self.html_file_path = tk.StringVar(value="")

        # Load saved config from config.json if present
        self._load_config()

        # Save config when application window closes
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)

        # Loaded data structures
        self.records: List[Dict[str, Any]] = []
        self.email_columns: List[Dict[str, Any]] = []
        self.dynamic_tags: List[str] = [
            "kode", "provinsi", "kepala_prov", "nama_ketua", "nama_koordinator", "nama_staf", "no"
        ]
        self.target_records: List[Dict[str, Any]] = []
        self.target_mode_desc: str = "0 Provinsi Terpilih [✓]"
        self.row_checked_state: Dict[str, bool] = {}
        self.custom_attachment_files: List[str] = []
        self.attachment_engine: Optional[AttachmentEngine] = None
        self.dispatcher: Optional[EmailDispatcher] = None
        self.is_running = False

        # History Manager & Tab 5 state
        self.history_manager = HistoryManager()
        self.hist_sender_filter = tk.StringVar(value="[Semua Akun Pengirim]")
        self.hist_status_filter = tk.StringVar(value="[Semua Status]")
        self.hist_search_var = tk.StringVar(value="")
        self.selected_history_id: Optional[int] = None

        # Thread communication queue
        self.log_queue = queue.Queue()

        # Build UI
        self._build_header()
        self._build_notebook()
        self._build_statusbar()

        # Populate initial tag helper UI
        self._update_tag_ui()

        # Auto-detect official BPS Excel file if present in workspace
        self._auto_detect_official_excel()

        # Start log consumer loop
        self._after_id = self.root.after(100, self._process_log_queue)

    def _auto_detect_official_excel(self):
        """Auto-load official BPS Excel file if present in workspace and excel_path is empty."""
        cur_path = self.excel_path.get().strip()
        if cur_path and os.path.exists(cur_path):
            return
        base_dir = os.path.dirname(os.path.abspath(__file__))
        official_name = "Daftar E-mail Kepala Provinsi dan PJK Statistik Industri Provinsi.xlsx"
        candidates = [
            os.path.join(base_dir, official_name),
            os.path.join(os.getcwd(), official_name),
        ]
        found = next((c for c in candidates if os.path.exists(c)), None)
        if found:
            self.excel_path.set(found)
            self._load_excel_data()

    def cleanup(self):
        if hasattr(self, "_after_id") and self._after_id:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass

    def _setup_styles(self):
        style = ttk.Style()
        # Use clam or default vista
        available_themes = style.theme_names()
        if "vista" in available_themes:
            style.theme_use("vista")
        elif "clam" in available_themes:
            style.theme_use("clam")

        # Fonts
        font_family = "Segoe UI"
        style.configure(".", font=(font_family, 9))
        style.configure("Header.TLabel", font=(font_family, 14, "bold"), foreground="#0D47A1")
        style.configure("SubHeader.TLabel", font=(font_family, 9), foreground="#555555")
        style.configure("Section.TLabelframe.Label", font=(font_family, 10, "bold"), foreground="#1565C0")
        style.configure("Badge.TLabel", font=(font_family, 9, "bold"), foreground="#2E7D32")
        style.configure("BadgeWarn.TLabel", font=(font_family, 9, "bold"), foreground="#D84315")

        style.configure("Action.TButton", font=(font_family, 9, "bold"))

    def _load_config(self):
        """Load persistent SMTP and application settings from config.json."""
        config_file = get_config_path()
        if not os.path.isfile(config_file):
            return

        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            if isinstance(data, dict):
                if data.get("smtp_host"):
                    self.smtp_host.set(str(data["smtp_host"]))
                if data.get("smtp_port"):
                    self.smtp_port.set(str(data["smtp_port"]))
                if "smtp_username" in data:
                    self.smtp_username.set(str(data["smtp_username"]))
                elif "smtp_user" in data:
                    self.smtp_username.set(str(data["smtp_user"]))
                if "sender_email" in data:
                    self.sender_email.set(str(data["sender_email"]))
                if "smtp_pass" in data:
                    self.smtp_pass.set(str(data["smtp_pass"]))
                if "email_delay" in data:
                    try:
                        self.email_delay.set(float(data["email_delay"]))
                    except (ValueError, TypeError):
                        pass
                if "smtp_timeout" in data:
                    try:
                        self.smtp_timeout.set(int(data["smtp_timeout"]))
                    except (ValueError, TypeError):
                        pass
                if data.get("email_subject"):
                    self.email_subject.set(str(data["email_subject"]))
                if data.get("html_template_path"):
                    h_path = str(data["html_template_path"])
                    if os.path.isfile(h_path):
                        self.html_file_path.set(h_path)
        except Exception as e:
            print(f"[WARN] Gagal memuat config.json: {e}")

    def _save_config(self, show_popup: bool = False):
        """Save current SMTP and application settings to config.json."""
        config_file = get_config_path()
        data = {
            "smtp_host": self.smtp_host.get().strip(),
            "smtp_port": self.smtp_port.get().strip(),
            "smtp_username": self.smtp_username.get().strip(),
            "sender_email": self.sender_email.get().strip(),
            "smtp_pass": self.smtp_pass.get(),
            "email_delay": self.email_delay.get(),
            "smtp_timeout": self.smtp_timeout.get(),
            "email_subject": self.email_subject.get().strip(),
            "html_template_path": self.html_file_path.get().strip(),
        }
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            if show_popup:
                messagebox.showinfo("Pengaturan Disimpan", f"Konfigurasi berhasil disimpan ke:\n{config_file}")
        except Exception as e:
            if show_popup:
                messagebox.showerror("Gagal Menyimpan", f"Gagal menyimpan konfigurasi ke config.json:\n{str(e)}")
            else:
                print(f"[WARN] Gagal menyimpan config.json: {e}")

    def _on_window_close(self):
        try:
            self._save_config(show_popup=False)
        except Exception:
            pass
        self.root.destroy()

    def _create_primary_button(self, parent, text: str, command: Any) -> tk.Button:
        """Create a prominent primary action button with guaranteed high-contrast white text on solid blue."""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg="#1565C0",
            fg="#FFFFFF",
            activebackground="#0D47A1",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.RAISED,
            bd=1,
            padx=14,
            pady=5,
            cursor="hand2",
        )
        btn.bind("<Enter>", lambda e: btn.config(bg="#0D47A1") if btn["state"] == tk.NORMAL else None)
        btn.bind("<Leave>", lambda e: btn.config(bg="#1565C0") if btn["state"] == tk.NORMAL else None)
        return btn

    def _build_header(self):
        header_frame = ttk.Frame(self.root, padding="15 10 15 10")
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(
            header_frame,
            text="📧 Automated Provincial Email Dispatcher (BPS RI)",
            style="Header.TLabel",
        )
        title_lbl.pack(anchor=tk.W)

        sub_lbl = ttk.Label(
            header_frame,
            text="Aplikasi distribusi dokumen dan naskah statistik ke 38 BPS Provinsi dengan pencocokan lampiran dinamis.",
            style="SubHeader.TLabel",
        )
        sub_lbl.pack(anchor=tk.W)

        sep = ttk.Separator(self.root, orient=tk.HORIZONTAL)
        sep.pack(fill=tk.X, padx=10, pady=(0, 5))

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=5)

        # Tab 1: Konfigurasi Sistem
        self.tab_config = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_config, text="  1. Konfigurasi Sistem  ")

        # Tab 2: Template Pesan & Editor HTML
        self.tab_template = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_template, text="  2. Template Pesan & Editor HTML  ")

        # Tab 3: Preview & Verifikasi
        self.tab_preview = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_preview, text="  3. Preview & Verifikasi  ")

        # Tab 4: Eksekusi & Log Real-Time
        self.tab_exec = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_exec, text="  4. Eksekusi & Log Pengiriman  ")

        # Tab 5: Riwayat Pengiriman
        self.tab_history = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_history, text="  5. 📜 Riwayat Pengiriman  ")

        self._setup_tab_config()
        self._setup_tab_template()
        self._setup_tab_preview()
        self._setup_tab_exec()
        self._setup_tab_history()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _setup_tab_config(self):
        # Two-column layout: Left column for Files & Attachments, Right column for SMTP Configuration
        container = ttk.Frame(self.tab_config)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(0, weight=1)

        left_col = ttk.Frame(container)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        right_col = ttk.Frame(container)
        right_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        # --- LEFT: Files & Data Source ---
        file_frame = ttk.LabelFrame(left_col, text=" 📂 Sumber Data & Berkas Lampiran ", style="Section.TLabelframe", padding=10)
        file_frame.pack(fill=tk.BOTH, expand=True)

        # Excel Picker
        ttk.Label(file_frame, text="File Excel Database Kontak BPS:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        excel_row = ttk.Frame(file_frame)
        excel_row.pack(fill=tk.X, pady=(2, 4))
        self.excel_entry = ttk.Entry(excel_row, textvariable=self.excel_path)
        self.excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(excel_row, text="Pilih Excel...", command=self._browse_excel).pack(side=tk.RIGHT)

        self.excel_status_lbl = ttk.Label(file_frame, text="Belum ada file Excel yang dipilih.", style="SubHeader.TLabel")
        self.excel_status_lbl.pack(anchor=tk.W, pady=(0, 8))

        # File Picker
        ttk.Label(file_frame, text="Sumber Lampiran Berkas (Folder atau Multi-File):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        folder_row = ttk.Frame(file_frame)
        folder_row.pack(fill=tk.X, pady=(2, 4))
        self.folder_entry = ttk.Entry(folder_row, textvariable=self.folder_path)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(folder_row, text="📄 + Pilih File...", command=self._browse_files).pack(side=tk.RIGHT, padx=(2, 0))
        ttk.Button(folder_row, text="📁 Pilih Folder...", command=self._browse_folder).pack(side=tk.RIGHT, padx=(2, 0))
        ttk.Button(folder_row, text="✕ Reset", command=self._reset_attachments).pack(side=tk.RIGHT, padx=(2, 0))

        # Folder / File status
        info_subrow = ttk.Frame(file_frame)
        info_subrow.pack(fill=tk.X, pady=(2, 4))
        self.folder_status_lbl = ttk.Label(info_subrow, text="Belum ada folder atau berkas lampiran yang dipilih.", style="SubHeader.TLabel")
        self.folder_status_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Helpful Naming Rules Box
        naming_box = ttk.LabelFrame(file_frame, text=" 💡 Format Penamaan Berkas / Folder (Otomatis Terdeteksi) ", padding=8)
        naming_box.pack(fill=tk.X, pady=(8, 4))
        ttk.Label(
            naming_box,
            text="• Berkas Daerah: Kode provinsi bebas di awal atau akhir (misal: 'Tabel Profil IMK 2025 - Provinsi 16.xlsx', '16_Tabel.xlsx', atau folder 'Provinsi 16/')",
            font=("Segoe UI", 8),
            foreground="#1565C0",
            wraplength=480,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=1)
        ttk.Label(
            naming_box,
            text="• Berkas Nasional: Awalan/akhiran '00' atau kata 'Nasional' bebas format (.xlsx, .pdf, .docx, .zip, dll., contoh: '00_Tabel_IMK.xlsx', '00_Naskah.pdf' -> otomatis ke semua 38 provinsi)",
            font=("Segoe UI", 8),
            foreground="#2E7D32",
            wraplength=480,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=1)
        ttk.Label(
            naming_box,
            text="• Deteksi Otomatis: Baris provinsi dengan lampiran langsung 'Siap Kirim' & tercentang [✓] otomatis tanpa perlu klik manual.",
            font=("Segoe UI", 8, "italic"),
            foreground="#666666",
            wraplength=480,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=1)

        # --- RIGHT: SMTP Configuration ---
        smtp_frame = ttk.LabelFrame(right_col, text=" 🔒 Konfigurasi Server Pengirim (SMTP) & Jeda ", style="Section.TLabelframe", padding=10)
        smtp_frame.pack(fill=tk.BOTH, expand=True)

        grid_smtp = ttk.Frame(smtp_frame)
        grid_smtp.pack(fill=tk.X, pady=4)
        grid_smtp.columnconfigure(1, weight=1)

        ttk.Label(grid_smtp, text="SMTP Host:").grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Entry(grid_smtp, textvariable=self.smtp_host, width=22).grid(row=0, column=1, sticky=tk.W, padx=5, pady=3)

        ttk.Label(grid_smtp, text="Port:").grid(row=0, column=2, sticky=tk.W, pady=3)
        ttk.Entry(grid_smtp, textvariable=self.smtp_port, width=8).grid(row=0, column=3, sticky=tk.W, padx=5, pady=3)

        ttk.Label(grid_smtp, text="Username SMTP / Login:").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(grid_smtp, textvariable=self.smtp_username).grid(row=1, column=1, columnspan=3, sticky=tk.EW, padx=5, pady=3)

        ttk.Label(grid_smtp, text="Email Pengirim (From):").grid(row=2, column=0, sticky=tk.W, pady=3)
        sender_frame = ttk.Frame(grid_smtp)
        sender_frame.grid(row=2, column=1, columnspan=3, sticky=tk.EW, padx=5, pady=3)
        self.sender_entry = ttk.Entry(sender_frame, textvariable=self.sender_email)
        self.sender_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(
            sender_frame,
            text="📋 Samakan dg Username",
            font=("Segoe UI", 8),
            relief=tk.GROOVE,
            bd=1,
            padx=5,
            cursor="hand2",
            command=self._sync_sender_with_username,
        ).pack(side=tk.RIGHT, padx=(4, 0))

        ttk.Label(grid_smtp, text="Password / App Password:").grid(row=3, column=0, sticky=tk.W, pady=3)
        self.pass_entry = ttk.Entry(grid_smtp, textvariable=self.smtp_pass, show="*")
        self.pass_entry.grid(row=3, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=3)

        self.show_pass_btn = ttk.Checkbutton(
            grid_smtp, text="Lihat", variable=self.smtp_show_pass, command=self._toggle_pass_visibility
        )
        self.show_pass_btn.grid(row=3, column=3, sticky=tk.W, padx=5, pady=3)

        # Delay Antar-Email (Free text decimal + spinbox)
        ttk.Label(grid_smtp, text="Jeda Antar-Email (detik):").grid(row=4, column=0, sticky=tk.W, pady=3)
        delay_spin = ttk.Spinbox(grid_smtp, from_=0.0, to=60.0, increment=0.5, textvariable=self.email_delay, width=8)
        delay_spin.grid(row=4, column=1, sticky=tk.W, padx=5, pady=3)

        # Timeout setting
        ttk.Label(grid_smtp, text="SMTP Timeout (detik):").grid(row=4, column=2, sticky=tk.W, pady=3)
        timeout_spin = ttk.Spinbox(grid_smtp, from_=5, to=300, increment=5, textvariable=self.smtp_timeout, width=8)
        timeout_spin.grid(row=4, column=3, sticky=tk.W, padx=5, pady=3)

        # Delay Presets Row
        preset_frame = ttk.Frame(smtp_frame)
        preset_frame.pack(fill=tk.X, pady=(2, 4))
        ttk.Label(preset_frame, text="Preset Jeda:", font=("Segoe UI", 8, "bold"), foreground="#555555").pack(side=tk.LEFT, padx=(0, 5))
        
        presets = [
            ("⚡ 0s (Instant)", 0.0),
            ("🚀 0.5s", 0.5),
            ("🛡️ 1.5s (Aman)", 1.5),
            ("🐢 3.0s", 3.0),
        ]
        for p_label, p_val in presets:
            btn = tk.Button(
                preset_frame,
                text=p_label,
                font=("Segoe UI", 8),
                relief=tk.GROOVE,
                bd=1,
                padx=5,
                pady=1,
                cursor="hand2",
                command=lambda v=p_val: self.email_delay.set(v),
            )
            btn.pack(side=tk.LEFT, padx=2)

        # Quick Server Profile Presets
        srv_preset_frame = ttk.Frame(smtp_frame)
        srv_preset_frame.pack(fill=tk.X, pady=(2, 6))
        ttk.Label(srv_preset_frame, text="Preset Server:", font=("Segoe UI", 8, "bold"), foreground="#1565C0").pack(side=tk.LEFT, padx=(0, 5))

        def set_server_preset(host_val, port_val, tip_text):
            if host_val:
                self.smtp_host.set(host_val)
            self.smtp_port.set(str(port_val))
            self.smtp_result_lbl.config(text=tip_text, foreground="#1565C0")

        tk.Button(
            srv_preset_frame,
            text="🏢 BPS STARTTLS (587)",
            font=("Segoe UI", 8),
            relief=tk.GROOVE,
            bd=1,
            padx=5,
            pady=1,
            cursor="hand2",
            command=lambda: set_server_preset("mail.bps.go.id", "587", "Port 587 (STARTTLS BPS)"),
        ).pack(side=tk.LEFT, padx=2)

        tk.Button(
            srv_preset_frame,
            text="🔒 BPS SSL (465)",
            font=("Segoe UI", 8),
            relief=tk.GROOVE,
            bd=1,
            padx=5,
            pady=1,
            cursor="hand2",
            command=lambda: set_server_preset("mail.bps.go.id", "465", "Port 465 (SSL BPS)"),
        ).pack(side=tk.LEFT, padx=2)

        tk.Button(
            srv_preset_frame,
            text="⚡ Intranet BPS (25)",
            font=("Segoe UI", 8),
            relief=tk.GROOVE,
            bd=1,
            padx=5,
            pady=1,
            cursor="hand2",
            command=lambda: set_server_preset("", "25", "Port 25 (Relay Intranet/VPN)"),
        ).pack(side=tk.LEFT, padx=2)

        tk.Button(
            srv_preset_frame,
            text="🔴 Gmail (587)",
            font=("Segoe UI", 8),
            relief=tk.GROOVE,
            bd=1,
            padx=5,
            pady=1,
            cursor="hand2",
            command=lambda: set_server_preset("smtp.gmail.com", "587", "Gmail Port 587"),
        ).pack(side=tk.LEFT, padx=2)

        # SMTP Buttons Row
        smtp_btn_row = ttk.Frame(smtp_frame)
        smtp_btn_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(smtp_btn_row, text="⚡ Uji Koneksi SMTP", command=self._test_smtp).pack(side=tk.LEFT)
        ttk.Button(smtp_btn_row, text="💾 Simpan Pengaturan", command=lambda: self._save_config(show_popup=True)).pack(side=tk.LEFT, padx=(6, 0))
        self.smtp_result_lbl = ttk.Label(smtp_btn_row, text="", font=("Segoe UI", 8))
        self.smtp_result_lbl.pack(side=tk.LEFT, padx=10)

        smtp_tip = ttk.Label(
            smtp_frame,
            text="💡 Petunjuk Konfigurasi Server SMTP:\n"
                 "• Server BPS RI (Rekomendasi Resmi): Masukkan Host SMTP BPS (misal mail.bps.go.id atau smtp.bps.go.id). "
                 "Gunakan Port 587 (STARTTLS), 465 (SSL), atau 25 (Relay Intranet/VPN). Masukkan username & password akun dinas BPS Anda "
                 "(atau kosongkan password jika menggunakan server relay internal yang tidak mewajibkan autentikasi).\n"
                 "• Gmail: Gunakan Host smtp.gmail.com, Port 587, dan App Password 16-digit dari Akun Google.\n"
                 "• Relay Eksternal (Brevo/SendGrid/dsb.): Masukkan username relay dan email pengirim yang sudah terverifikasi.",
            font=("Segoe UI", 8, "italic"),
            foreground="#555555",
            wraplength=460,
            justify=tk.LEFT,
        )
        smtp_tip.pack(anchor=tk.W, pady=(8, 4))

        # Bottom navigation button to Tab 2
        nav_config = ttk.Frame(right_col)
        nav_config.pack(fill=tk.X, pady=(10, 0))
        self._create_primary_button(
            nav_config,
            text="➡️ Lanjut ke Template Pesan & Editor HTML",
            command=lambda: self.notebook.select(self.tab_template),
        ).pack(side=tk.RIGHT)

    def _setup_tab_template(self):
        # Tab 2: Full-screen Template and HTML WYSIWYG Editor
        tmpl_frame = ttk.LabelFrame(
            self.tab_template,
            text=" 📝 Subjek & Template Pesan Email (Dukungan HTML & WYSIWYG Editor) ",
            style="Section.TLabelframe",
            padding=12,
        )
        tmpl_frame.pack(fill=tk.BOTH, expand=True)
        tmpl_frame.columnconfigure(0, weight=1)
        tmpl_frame.rowconfigure(3, weight=1)

        # Row 0: Subject Label
        ttk.Label(tmpl_frame, text="Subjek Email:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 2))

        # Row 1: Enlarged Subject Entry
        self.subject_entry = ttk.Entry(tmpl_frame, textvariable=self.email_subject, font=("Segoe UI", 10))
        self.subject_entry.grid(row=1, column=0, sticky="ew", pady=(2, 6), ipady=4)

        # Row 2: HTML File Loader Box (Simpel & Praktis)
        html_box = ttk.LabelFrame(
            tmpl_frame,
            text=" 📂 Taruh / Muat Berkas Template HTML (.html / .htm) ",
            padding=6,
        )
        html_box.grid(row=2, column=0, sticky="ew", pady=(2, 6))
        html_box.columnconfigure(1, weight=1)

        ttk.Label(html_box, text="File HTML:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(4, 6))
        self.html_file_entry = ttk.Entry(html_box, textvariable=self.html_file_path, font=("Segoe UI", 9))
        self.html_file_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))

        ttk.Button(html_box, text="📁 Pilih File HTML...", command=self._browse_html_file).grid(row=0, column=2, padx=2)
        ttk.Button(html_box, text="🔄 Muat Ulang", command=self._reload_html_file).grid(row=0, column=3, padx=2)
        ttk.Button(html_box, text="✕ Reset", command=self._reset_html_file).grid(row=0, column=4, padx=2)

        self.html_status_lbl = ttk.Label(
            html_box,
            text="💡 Anda dapat memilih file template .html dari komputer agar langsung termuat rapi ke area pesan di bawah.",
            font=("Segoe UI", 8, "italic"),
            foreground="#555555",
        )
        self.html_status_lbl.grid(row=1, column=0, columnspan=5, sticky="w", padx=(4, 0), pady=(3, 1))

        # Row 3: WYSIWYG Toolbar & Mode Info
        bar_header = ttk.Frame(tmpl_frame)
        bar_header.grid(row=3, column=0, sticky="ew", pady=(4, 2))
        bar_header.columnconfigure(1, weight=1)

        ttk.Label(bar_header, text="Isi Pesan (Body Email):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(
            bar_header,
            text="🏷️ Format Teks Biasa & Kode HTML Didukung Penuh",
            font=("Segoe UI", 8, "bold"),
            foreground="#0D47A1",
        ).grid(row=0, column=1, sticky="e")

        # Row 4: Toolbar Frame
        toolbar = ttk.Frame(tmpl_frame)
        toolbar.grid(row=4, column=0, sticky="ew", pady=(0, 4))

        # Formatting buttons
        btn_b = tk.Button(toolbar, text="B", font=("Segoe UI", 9, "bold"), width=3, relief=tk.GROOVE, bd=1, cursor="hand2", command=self._format_bold)
        btn_b.pack(side=tk.LEFT, padx=1)

        btn_i = tk.Button(toolbar, text="I", font=("Segoe UI", 9, "italic"), width=3, relief=tk.GROOVE, bd=1, cursor="hand2", command=self._format_italic)
        btn_i.pack(side=tk.LEFT, padx=1)

        btn_u = tk.Button(toolbar, text="U", font=("Segoe UI", 9, "underline"), width=3, relief=tk.GROOVE, bd=1, cursor="hand2", command=self._format_underline)
        btn_u.pack(side=tk.LEFT, padx=1)

        btn_h2 = tk.Button(toolbar, text="H2 Judul", font=("Segoe UI", 8, "bold"), relief=tk.GROOVE, bd=1, padx=6, cursor="hand2", command=self._format_heading)
        btn_h2.pack(side=tk.LEFT, padx=(4, 1))

        btn_link = tk.Button(toolbar, text="🔗 Link", font=("Segoe UI", 8), relief=tk.GROOVE, bd=1, padx=6, cursor="hand2", command=self._insert_link_dialog)
        btn_link.pack(side=tk.LEFT, padx=1)

        btn_tbl = tk.Button(toolbar, text="📊 Tabel HTML", font=("Segoe UI", 8), relief=tk.GROOVE, bd=1, padx=6, cursor="hand2", command=self._insert_table_dialog)
        btn_tbl.pack(side=tk.LEFT, padx=1)

        sep_t = ttk.Separator(toolbar, orient=tk.VERTICAL)
        sep_t.pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)

        btn_open_html = tk.Button(
            toolbar,
            text="📂 Buka File HTML...",
            font=("Segoe UI", 8, "bold"),
            bg="#FFF9C4",
            fg="#E65100",
            relief=tk.GROOVE,
            bd=1,
            padx=8,
            cursor="hand2",
            command=self._browse_html_file,
        )
        btn_open_html.pack(side=tk.LEFT, padx=2)

        btn_bps = tk.Button(
            toolbar,
            text="📋 Template Resmi BPS",
            font=("Segoe UI", 8, "bold"),
            bg="#E8F5E9",
            fg="#2E7D32",
            relief=tk.GROOVE,
            bd=1,
            padx=8,
            cursor="hand2",
            command=self._load_formal_bps_template,
        )
        btn_bps.pack(side=tk.LEFT, padx=2)

        btn_prev = tk.Button(
            toolbar,
            text="👁️ Pratinjau HTML",
            font=("Segoe UI", 8, "bold"),
            bg="#E3F2FD",
            fg="#0D47A1",
            relief=tk.GROOVE,
            bd=1,
            padx=8,
            cursor="hand2",
            command=self._preview_html_dialog,
        )
        btn_prev.pack(side=tk.LEFT, padx=2)

        btn_clean = tk.Button(
            toolbar,
            text="🧹 Bersihkan HTML",
            font=("Segoe UI", 8),
            relief=tk.GROOVE,
            bd=1,
            padx=6,
            cursor="hand2",
            command=self._clean_html_tags,
        )
        btn_clean.pack(side=tk.RIGHT, padx=2)

        # Row 5: Body Template Text Area
        body_frame = ttk.Frame(tmpl_frame)
        body_frame.grid(row=5, column=0, sticky="nsew", pady=(2, 4))
        body_frame.columnconfigure(0, weight=1)
        body_frame.rowconfigure(0, weight=1)
        tmpl_frame.rowconfigure(5, weight=1)

        self.body_text = tk.Text(body_frame, wrap=tk.WORD, font=("Segoe UI", 10), height=14, undo=True)
        body_scroll = ttk.Scrollbar(body_frame, orient=tk.VERTICAL, command=self.body_text.yview)
        self.body_text.configure(yscrollcommand=body_scroll.set)
        self.body_text.grid(row=0, column=0, sticky="nsew")
        body_scroll.grid(row=0, column=1, sticky="ns")

        default_body = (
            "Yth. Kepala BPS Provinsi {{provinsi}} ({{nama_kepala}}),\n"
            "Cc: Ketua Tim / PJK ({{nama_ketua}})\n\n"
            "Bersama ini kami sampaikan dokumen naskah statistik dan berkas lampiran pendukung "
            "untuk wilayah Provinsi {{provinsi}} (Kode: {{kode}}).\n\n"
            "Seluruh dokumen terkait telah dilampirkan pada email ini.\n"
            "Mohon perkenan untuk melakukan konfirmasi penerimaan berkas ini.\n\n"
            "Demikian kami sampaikan. Atas perhatian dan kerja sama yang baik, diucapkan terima kasih.\n\n"
            "Hormat kami,\n"
            "Tim Pengelola & Distribusi Naskah Statistik BPS RI"
        )
        self.body_text.insert(tk.END, default_body)

        # Focus tracking so tag buttons insert into whichever field had focus
        self.last_focused_text_widget = self.body_text
        self.subject_entry.bind("<FocusIn>", lambda e: setattr(self, "last_focused_text_widget", self.subject_entry))
        self.body_text.bind("<FocusIn>", lambda e: setattr(self, "last_focused_text_widget", self.body_text))

        # Row 6: Dynamic Tags Helper Box (Below Body Email)
        tag_box = ttk.LabelFrame(tmpl_frame, text=" 🏷️ Tag Dinamis Tersedia (Klik Tag untuk Menyisipkan ke Posisi Kursor) ", padding=8)
        tag_box.grid(row=6, column=0, sticky="ew", pady=(4, 6))
        tag_box.columnconfigure(0, weight=1)

        self.tag_guide_lbl = ttk.Label(
            tag_box,
            text="Tag dinamis tersedia: {{kode}}, {{provinsi}}, {{kepala_prov}}, {{nama_ketua}}, {{nama_koordinator}}, {{nama_staf}}, {{no}}",
            font=("Segoe UI", 8, "italic"),
            foreground="#0D47A1",
            wraplength=950,
            justify=tk.LEFT,
        )
        self.tag_guide_lbl.pack(anchor=tk.W, pady=(0, 4))

        self.tag_buttons_frame = ttk.Frame(tag_box)
        self.tag_buttons_frame.pack(fill=tk.X, expand=True)

        # Row 7: Bottom Navigation to Tab 3
        nav_frame = ttk.Frame(tmpl_frame)
        nav_frame.grid(row=7, column=0, sticky="ew", pady=(4, 0))
        nav_frame.columnconfigure(0, weight=1)

        ttk.Label(
            nav_frame,
            text="💡 Tip: Anda dapat memilih file .html langsung atau klik 'Pratinjau HTML' untuk melihat hasil tampilan visual.",
            font=("Segoe UI", 8, "italic"),
            foreground="#555555",
        ).grid(row=0, column=0, sticky="w")

        self._create_primary_button(
            nav_frame, text="➡️ Lanjut ke Preview & Verifikasi", command=self._go_to_preview
        ).grid(row=0, column=1, sticky="e")

    def _set_delay_preset(self, seconds: float):
        """Set delay directly from preset button."""
        self.email_delay.set(seconds)

    # --- HTML FILE LOADER METHODS ---
    def _browse_html_file(self):
        file_path = filedialog.askopenfilename(
            title="Pilih File Template Pesan HTML",
            filetypes=[
                ("Berkas HTML", "*.html *.htm"),
                ("Berkas Teks", "*.txt"),
                ("Semua Berkas", "*.*"),
            ],
            parent=self.root,
        )
        if file_path:
            self.html_file_path.set(file_path)
            self._load_html_file_content(file_path)

    def _load_html_file_content(self, file_path: str):
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("File Tidak Ditemukan", f"Berkas HTML tidak ditemukan:\n{file_path}", parent=self.root)
            return

        try:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(file_path, "r", encoding="latin-1") as f:
                    content = f.read()

            self.body_text.delete("1.0", tk.END)
            self.body_text.insert(tk.END, content)

            title_match = re.search(r"<title>(.*?)</title>", content, re.IGNORECASE | re.DOTALL)
            if title_match:
                extracted_title = title_match.group(1).strip()
                if extracted_title:
                    self.email_subject.set(extracted_title)

            file_sz = get_file_size_formatted(file_path)
            file_name = os.path.basename(file_path)
            if hasattr(self, "html_status_lbl"):
                self.html_status_lbl.config(
                    text=f"✓ File HTML '{file_name}' ({file_sz}) berhasil dimuat ke editor.",
                    foreground="#2E7D32",
                )
            messagebox.showinfo(
                "File HTML Berhasil Dimuat",
                f"Isi berkas '{file_name}' ({file_sz}) berhasil dimuat ke area Isi Pesan.\n\n"
                f"Anda dapat menyuntingnya atau klik '👁️ Pratinjau HTML' untuk melihat tampilannya.",
                parent=self.root,
            )
        except Exception as e:
            messagebox.showerror("Gagal Membaca File", f"Gagal membaca berkas HTML:\n{str(e)}", parent=self.root)

    def _reload_html_file(self):
        f = self.html_file_path.get().strip()
        if not f:
            self._browse_html_file()
        else:
            self._load_html_file_content(f)

    def _reset_html_file(self):
        self.html_file_path.set("")
        if hasattr(self, "html_status_lbl"):
            self.html_status_lbl.config(
                text="Belum ada file HTML eksternal yang dipilih.",
                foreground="#555555",
            )

    # --- WYSIWYG & HTML HELPER METHODS ---
    def _format_wrap_body(self, tag_open: str, tag_close: str, default_text: str = ""):
        """Wrap selected text or insert tag with placeholder into body_text."""
        try:
            sel_start = self.body_text.index(tk.SEL_FIRST)
            sel_end = self.body_text.index(tk.SEL_LAST)
            selected_text = self.body_text.get(sel_start, sel_end)
            self.body_text.delete(sel_start, sel_end)
            self.body_text.insert(sel_start, f"{tag_open}{selected_text}{tag_close}")
            self.body_text.focus_set()
        except tk.TclError:
            insert_pos = self.body_text.index(tk.INSERT)
            placeholder = default_text or "teks"
            self.body_text.insert(insert_pos, f"{tag_open}{placeholder}{tag_close}")
            self.body_text.focus_set()

    def _format_bold(self):
        self._format_wrap_body("<b>", "</b>", "Teks Tebal")

    def _format_italic(self):
        self._format_wrap_body("<i>", "</i>", "Teks Miring")

    def _format_underline(self):
        self._format_wrap_body("<u>", "</u>", "Teks Garis Bawah")

    def _format_heading(self):
        self._format_wrap_body('<h2 style="color: #004b87; margin-top: 15px; margin-bottom: 5px;">', "</h2>", "Judul Bagian")

    def _insert_link_dialog(self):
        url = simpledialog.askstring("Sisipkan Tautan (Link)", "Masukkan URL / Alamat Website (contoh: https://bps.go.id):", parent=self.root)
        if not url:
            return
        url = url.strip()
        selected_text = ""
        try:
            selected_text = self.body_text.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass

        link_text = simpledialog.askstring("Teks Tautan", "Masukkan teks yang ingin ditampilkan untuk tautan:", initialvalue=selected_text or url, parent=self.root)
        if not link_text:
            link_text = url

        html_link = f'<a href="{url}" style="color: #004b87; font-weight: bold; text-decoration: underline;">{link_text}</a>'
        try:
            if selected_text:
                self.body_text.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.body_text.insert(tk.INSERT, html_link)
        except Exception:
            self.body_text.insert(tk.END, html_link)
        self.body_text.focus_set()

    def _insert_table_dialog(self):
        sample_table = (
            '\n<table style="border-collapse: collapse; width: 100%; margin: 15px 0; font-size: 13px; font-family: Segoe UI, sans-serif;">\n'
            '  <thead>\n'
            '    <tr style="background-color: #004b87; color: #ffffff;">\n'
            '      <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">No</th>\n'
            '      <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">Nama Dokumen Publikasi</th>\n'
            '      <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">Status Berkas</th>\n'
            '    </tr>\n'
            '  </thead>\n'
            '  <tbody>\n'
            '    <tr>\n'
            '      <td style="border: 1px solid #dddddd; padding: 8px;">1</td>\n'
            '      <td style="border: 1px solid #dddddd; padding: 8px;">Naskah Publikasi Statistik Provinsi {{provinsi}}</td>\n'
            '      <td style="border: 1px solid #dddddd; padding: 8px; color: #2E7D32; font-weight: bold;">Terlampir</td>\n'
            '    </tr>\n'
            '    <tr style="background-color: #f9f9f9;">\n'
            '      <td style="border: 1px solid #dddddd; padding: 8px;">2</td>\n'
            '      <td style="border: 1px solid #dddddd; padding: 8px;">Tabel Profil IMK 2025 (Kode: {{kode}})</td>\n'
            '      <td style="border: 1px solid #dddddd; padding: 8px; color: #2E7D32; font-weight: bold;">Terlampir</td>\n'
            '    </tr>\n'
            '  </tbody>\n'
            '</table>\n'
        )
        self.body_text.insert(tk.INSERT, sample_table)
        self.body_text.focus_set()

    def _load_formal_bps_template(self):
        confirm = messagebox.askyesno(
            "Muat Template Resmi BPS",
            "Apakah Anda ingin mengganti isi pesan saat ini dengan Template Resmi HTML BPS RI?\n\n"
            "Template ini dilengkapi kop surat resmi, kotak catatan penting, dan format tata letak profesional.",
            icon=messagebox.QUESTION,
        )
        if not confirm:
            return

        formal_bps_html = (
            '<div style="font-family: \'Segoe UI\', Arial, sans-serif; font-size: 14px; color: #222222; line-height: 1.6; max-width: 700px; margin: auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff;">\n'
            '  <!-- Header Kop Resmi BPS -->\n'
            '  <div style="border-bottom: 3px solid #004b87; padding-bottom: 12px; margin-bottom: 20px;">\n'
            '    <h2 style="color: #004b87; margin: 0; font-size: 20px; text-transform: uppercase; letter-spacing: 0.5px;">BADAN PUSAT STATISTIK REPUBLIK INDONESIA</h2>\n'
            '    <p style="margin: 3px 0 0 0; font-size: 12px; color: #666666;">Jl. Dr. Sutomo No. 6-8 Jakarta 10710 | Telp (021) 3841195 | www.bps.go.id</p>\n'
            '  </div>\n\n'
            '  <p>Yth. <b>Kepala BPS Provinsi {{provinsi}}</b> ({{nama_kepala}}),<br>\n'
            '  Cc: Ketua Tim / PJK ({{nama_ketua}})</p>\n\n'
            '  <p>Bersama ini kami sampaikan dokumen naskah publikasi statistik dan berkas lampiran pendukung untuk wilayah <b>Provinsi {{provinsi}} (Kode Wilayah: {{kode}})</b>.</p>\n\n'
            '  <div style="background-color: #f4f8fb; border-left: 4px solid #004b87; padding: 12px 15px; margin: 15px 0; border-radius: 0 4px 4px 0;">\n'
            '    <p style="margin: 0; font-weight: bold; color: #004b87;">📌 Catatan Pengiriman:</p>\n'
            '    <ul style="margin: 6px 0 0 18px; padding: 0;">\n'
            '      <li>Seluruh berkas naskah daerah dan pedoman nasional telah terlampir pada email ini.</li>\n'
            '      <li>Mohon perkenan Bapak/Ibu untuk memeriksa kelengkapan data dan memberikan konfirmasi penerimaan.</li>\n'
            '    </ul>\n'
            '  </div>\n\n'
            '  <p>Demikian kami sampaikan. Atas perhatian, kerja sama, dan dedikasi Bapak/Ibu sekalian, kami ucapkan terima kasih.</p>\n\n'
            '  <br>\n'
            '  <div style="border-top: 1px solid #eeeeee; padding-top: 12px; margin-top: 20px;">\n'
            '    <p style="margin: 0; font-weight: bold; color: #004b87;">Tim Pengelola &amp; Distribusi Naskah Statistik</p>\n'
            '    <p style="margin: 2px 0 0 0; color: #555555; font-size: 13px;">Badan Pusat Statistik Republik Indonesia</p>\n'
            '  </div>\n'
            '</div>'
        )
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert(tk.END, formal_bps_html)

    def _preview_html_dialog(self):
        """Display an HTML preview popup with option to open visual rendering in default browser."""
        subject_raw = self.email_subject.get()
        body_raw = self.body_text.get("1.0", tk.END).strip()

        sample_rec = {
            "kode": "11",
            "provinsi": "Aceh",
            "nama_kepala": "Dr. Ahmad, M.Si",
            "nama_ketua": "Budi Santoso, S.Si",
            "nama_koordinator": "Ir. Siti Rahma",
            "nama_staf": "Andi Pratama",
            "no": "1",
        }
        if self.records:
            sample_rec = self.records[0]

        subject_rendered = replace_template(subject_raw, sample_rec)
        body_rendered = replace_template(body_raw, sample_rec)

        dialog = tk.Toplevel(self.root)
        dialog.title("👁️ Pratinjau Tampilan Email HTML")
        dialog.geometry("780x600")
        dialog.transient(self.root)

        top_info = ttk.Frame(dialog, padding=10)
        top_info.pack(fill=tk.X)

        ttk.Label(top_info, text="Subjek:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Label(top_info, text=subject_rendered, font=("Segoe UI", 9)).grid(row=0, column=1, sticky=tk.W, padx=6, pady=2)

        ttk.Label(top_info, text="Simulasi Target:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Label(top_info, text=f"BPS Provinsi {sample_rec.get('provinsi', '')} (Kode: {sample_rec.get('kode', '')})", font=("Segoe UI", 9)).grid(row=1, column=1, sticky=tk.W, padx=6, pady=2)

        btn_bar = ttk.Frame(dialog, padding=(10, 0, 10, 10))
        btn_bar.pack(fill=tk.X)

        def open_in_browser():
            full_html = (
                "<!DOCTYPE html>\n"
                "<html>\n"
                "<head>\n"
                "<meta charset=\"utf-8\">\n"
                "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
                f"<title>{subject_rendered}</title>\n"
                "<style>\n"
                "body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f8; margin: 0; padding: 25px; }\n"
                ".email-card { background: #ffffff; max-width: 720px; margin: 0 auto; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); overflow: hidden; }\n"
                ".email-meta { background: #004b87; color: #ffffff; padding: 15px 20px; font-size: 13px; line-height: 1.6; }\n"
                ".email-meta b { color: #bbdefb; }\n"
                ".email-body { padding: 25px; }\n"
                "</style>\n"
                "</head>\n"
                "<body>\n"
                "<div class=\"email-card\">\n"
                "  <div class=\"email-meta\">\n"
                f"    <div><b>Subjek:</b> {subject_rendered}</div>\n"
                f"    <div><b>Dari:</b> {self.sender_email.get() or 'admin@bps.go.id'}</div>\n"
                f"    <div><b>Kepada:</b> {sample_rec.get('provinsi', '')} (Simulasi)</div>\n"
                "  </div>\n"
                "  <div class=\"email-body\">\n"
                f"{body_rendered}\n"
                "  </div>\n"
                "</div>\n"
                "</body>\n"
                "</html>\n"
            )
            try:
                with tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8") as tf:
                    tf.write(full_html)
                    tmp_file = tf.name
                webbrowser.open("file://" + os.path.abspath(tmp_file))
            except Exception as e:
                messagebox.showerror("Gagal Membuka Browser", f"Gagal membuka browser: {e}")

        self._create_primary_button(
            btn_bar,
            text="🌐 Buka Tampilan Visual Lengkap di Web Browser (Chrome / Edge)",
            command=open_in_browser,
        ).pack(side=tk.LEFT)

        ttk.Button(btn_bar, text="Tutup", command=dialog.destroy).pack(side=tk.RIGHT)

        text_preview_frame = ttk.LabelFrame(dialog, text=" Pratinjau Teks / Kode Sumber Template ", padding=8)
        text_preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        preview_box = scrolledtext.ScrolledText(text_preview_frame, wrap=tk.WORD, font=("Consolas", 9))
        preview_box.pack(fill=tk.BOTH, expand=True)
        preview_box.insert(tk.END, body_rendered)
        preview_box.config(state=tk.DISABLED)

    def _clean_html_tags(self):
        """Strip HTML tags and leave clean plain text."""
        raw_text = self.body_text.get("1.0", tk.END).strip()
        if not raw_text:
            return
        confirm = messagebox.askyesno(
            "Bersihkan Tag HTML",
            "Apakah Anda yakin ingin menghapus semua tag HTML dan menyisakan teks biasa?",
            icon=messagebox.QUESTION,
        )
        if not confirm:
            return
        clean_text = re.sub(r"<br\s*/?>", "\n", raw_text, flags=re.IGNORECASE)
        clean_text = re.sub(r"</p>", "\n\n", clean_text, flags=re.IGNORECASE)
        clean_text = re.sub(r"<[^>]+>", "", clean_text)
        clean_text = re.sub(r"\n{3,}", "\n\n", clean_text).strip()
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert(tk.END, clean_text)

    def _setup_tab_preview(self):
        # Preview Header Controls
        ctrl_frame = ttk.Frame(self.tab_preview)
        ctrl_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(ctrl_frame, text="🔄 Segarkan Data", command=self._refresh_preview).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(ctrl_frame, text="☑ Pilih Semua", command=self._select_all_rows).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(ctrl_frame, text="☐ Batal Pilih Semua", command=self._deselect_all_rows).pack(side=tk.LEFT, padx=(0, 8))
        self.preview_stats_lbl = ttk.Label(ctrl_frame, text="Klik 'Segarkan Data' untuk memuat tabel verifikasi.", font=("Segoe UI", 9, "bold"))
        self.preview_stats_lbl.pack(side=tk.LEFT, padx=5)

        self._create_primary_button(
            ctrl_frame,
            text="🚀 Lanjut ke Eksekusi Pengiriman",
            command=self._proceed_to_execution,
        ).pack(side=tk.RIGHT)

        # Split pane: Treeview on top/middle, detail inspection at bottom
        preview_paned = ttk.PanedWindow(self.tab_preview, orient=tk.VERTICAL)
        preview_paned.pack(fill=tk.BOTH, expand=True)

        table_frame = ttk.Frame(preview_paned)
        preview_paned.add(table_frame, weight=3)

        # Treeview for records with dynamic email columns
        self.tree = ttk.Treeview(table_frame, show="headings", selectmode="extended")
        self._configure_tree_columns()

        # Scrollbars
        y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        # Bindings for selection and checkbox toggle
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<space>", self._toggle_selected_checkboxes)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Shift-MouseWheel>", lambda e: self.tree.xview_scroll(int(-1 * (e.delta / 120)), "units"))

        # Tag configurations for colors
        self.tree.tag_configure("valid", foreground="#1B5E20")
        self.tree.tag_configure("warning", foreground="#E65100", background="#FFF8E1")
        self.tree.tag_configure("error", foreground="#B71C1C", background="#FFEBEE")

        # Details Panel
        detail_frame = ttk.LabelFrame(preview_paned, text=" 🔍 Detail Lampiran & Pratinjau Email Baris Terpilih ", padding=8)
        preview_paned.add(detail_frame, weight=2)

        self.detail_text = scrolledtext.ScrolledText(detail_frame, height=6, wrap=tk.WORD, font=("Consolas", 9))
        self.detail_text.pack(fill=tk.BOTH, expand=True)
        self.detail_text.insert(tk.END, "Pilih salah satu baris provinsi di tabel atas untuk melihat pratinjau lengkap.")
        self.detail_text.config(state=tk.DISABLED)

    def _setup_tab_exec(self):
        # Target Indicator Banner
        target_banner = ttk.LabelFrame(self.tab_exec, text=" 🎯 Sasaran Pengiriman Tercentang ", padding="10 8 10 8")
        target_banner.pack(fill=tk.X, pady=(0, 10))

        self.target_banner_lbl = ttk.Label(
            target_banner,
            text="Target Pengiriman: 0 Provinsi Terpilih [✓] (Menyiapkan pengiriman untuk 0 provinsi terpilih)",
            font=("Segoe UI", 10, "bold"),
            foreground="#0D47A1",
        )
        self.target_banner_lbl.pack(side=tk.LEFT)

        ttk.Button(
            target_banner,
            text="✏️ Ubah Pilihan di Tabel Preview",
            command=lambda: self.notebook.select(self.tab_preview),
        ).pack(side=tk.RIGHT)

        # Top summary cards
        summary_frame = ttk.Frame(self.tab_exec)
        summary_frame.pack(fill=tk.X, pady=(0, 10))

        self.card_total = self._create_metric_card(summary_frame, "Total Sasaran", "0", "#1565C0")
        self.card_total.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        self.card_success = self._create_metric_card(summary_frame, "Berhasil Terkirim", "0", "#2E7D32")
        self.card_success.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        self.card_skipped = self._create_metric_card(summary_frame, "Dilewati / Peringatan", "0", "#EF6C00")
        self.card_skipped.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        self.card_error = self._create_metric_card(summary_frame, "Gagal Terkirim", "0", "#C62828")
        self.card_error.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        # Progress bar
        prog_frame = ttk.Frame(self.tab_exec)
        prog_frame.pack(fill=tk.X, pady=(5, 10))

        self.prog_lbl = ttk.Label(prog_frame, text="Menunggu eksekusi...", font=("Segoe UI", 9, "bold"))
        self.prog_lbl.pack(anchor=tk.W, pady=(0, 2))

        self.progressbar = ttk.Progressbar(prog_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.progressbar.pack(fill=tk.X)

        # Action Buttons
        btn_frame = ttk.Frame(self.tab_exec)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_dry_run = tk.Button(
            btn_frame,
            text="🧪 Simulasi Pengiriman (Dry Run)",
            command=self._start_dry_run,
            bg="#37474F",
            fg="#FFFFFF",
            activebackground="#263238",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.RAISED,
            bd=1,
            padx=12,
            pady=5,
            cursor="hand2",
        )
        self.btn_dry_run.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_send = tk.Button(
            btn_frame,
            text="🚀 Mulai Pengiriman Email Riil",
            command=self._confirm_and_start_real_send,
            bg="#1565C0",
            fg="#FFFFFF",
            activebackground="#0D47A1",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.RAISED,
            bd=1,
            padx=14,
            pady=5,
            cursor="hand2",
        )
        self.btn_send.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_stop = tk.Button(
            btn_frame,
            text="🛑 Hentikan Proses",
            command=self._stop_dispatch,
            bg="#E0E0E0",
            fg="#9E9E9E",
            activebackground="#C62828",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            bd=1,
            padx=12,
            pady=5,
            state=tk.DISABLED,
        )
        self.btn_stop.pack(side=tk.LEFT)

        ttk.Button(btn_frame, text="Bersihkan Log", command=self._clear_log).pack(side=tk.RIGHT)

        # Real-time Colored Console
        log_frame = ttk.LabelFrame(self.tab_exec, text=" 📜 Log Aktivitas Pengiriman Real-Time ", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_console = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 9), bg="#1E1E1E", fg="#D4D4D4")
        self.log_console.pack(fill=tk.BOTH, expand=True)

        # Console color tags
        self.log_console.tag_configure("SUCCESS", foreground="#4CAF50")
        self.log_console.tag_configure("SKIPPED", foreground="#FF9800")
        self.log_console.tag_configure("ERROR", foreground="#F44336")
        self.log_console.tag_configure("INFO", foreground="#64B5F6")
        self.log_console.tag_configure("TIMESTAMP", foreground="#888888")

        self._log_to_console("INFO", "Sistem siap. Silakan verifikasi data di Tab 2 sebelum mengirim email.")

    def _on_tab_changed(self, event=None):
        try:
            selected_tab = self.notebook.select()
            if hasattr(self, "tab_history") and selected_tab == str(self.tab_history):
                self._refresh_history_ui()
        except Exception:
            pass

    def _setup_tab_history(self):
        """Construct Tab 5: Riwayat Pengiriman with multi-sender filter, search, master table, and detail viewer."""
        container = ttk.Frame(self.tab_history)
        container.pack(fill=tk.BOTH, expand=True)

        # 1. Top Filter & Action Toolbar
        toolbar = ttk.LabelFrame(container, text=" 🔍 Filter & Pencarian Riwayat Pengiriman ", style="Section.TLabelframe", padding=8)
        toolbar.pack(fill=tk.X, pady=(0, 8))

        f_row = ttk.Frame(toolbar)
        f_row.pack(fill=tk.X, pady=2)

        # Sender Account Filter (Dropdown)
        ttk.Label(f_row, text="Akun Pengirim:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.combo_hist_sender = ttk.Combobox(
            f_row,
            textvariable=self.hist_sender_filter,
            values=["[Semua Akun Pengirim]"],
            state="readonly",
            width=28,
            font=("Segoe UI", 9),
        )
        self.combo_hist_sender.pack(side=tk.LEFT, padx=(0, 12))
        self.combo_hist_sender.bind("<<ComboboxSelected>>", self._on_history_filter_change)

        # Status Filter (Dropdown)
        ttk.Label(f_row, text="Status:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.combo_hist_status = ttk.Combobox(
            f_row,
            textvariable=self.hist_status_filter,
            values=["[Semua Status]", "✓ Berhasil (SUCCESS)", "✗ Gagal (FAILED)"],
            state="readonly",
            width=20,
            font=("Segoe UI", 9),
        )
        self.combo_hist_status.pack(side=tk.LEFT, padx=(0, 12))
        self.combo_hist_status.bind("<<ComboboxSelected>>", self._on_history_filter_change)

        # Search Query
        ttk.Label(f_row, text="Pencarian:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.entry_hist_search = ttk.Entry(f_row, textvariable=self.hist_search_var, width=22, font=("Segoe UI", 9))
        self.entry_hist_search.pack(side=tk.LEFT, padx=(0, 4))
        self.entry_hist_search.bind("<Return>", self._on_history_filter_change)

        ttk.Button(f_row, text="Cari", command=self._on_history_filter_change).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(f_row, text="✕ Reset", command=self._reset_history_filter).pack(side=tk.LEFT, padx=(0, 10))

        # Right Action Buttons: Refresh, Export, Clear
        ttk.Button(f_row, text="🔄 Refresh", command=self._refresh_history_ui).pack(side=tk.RIGHT, padx=(2, 0))
        ttk.Button(f_row, text="📊 Ekspor Excel", command=self._export_history_excel).pack(side=tk.RIGHT, padx=(2, 0))
        ttk.Button(f_row, text="🗑️ Hapus Riwayat", command=self._confirm_clear_history).pack(side=tk.RIGHT, padx=(2, 0))

        # 2. Main Vertical PanedWindow (Split: Master Table on Top, Detail Viewer on Bottom)
        paned = ttk.PanedWindow(container, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # --- Top Pane: History Table ---
        top_pane = ttk.Frame(paned)
        paned.add(top_pane, weight=4)

        table_lf = ttk.LabelFrame(top_pane, text=" 📋 Daftar Riwayat Email Terkirim ", padding=6)
        table_lf.pack(fill=tk.BOTH, expand=True)

        tree_frame = ttk.Frame(table_lf)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("id", "timestamp", "sender", "provinsi", "recipient", "subject", "status", "type")
        self.history_tree = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="headings",
            selectmode="browse",
            height=8,
        )

        self.history_tree.heading("id", text="ID", anchor=tk.CENTER)
        self.history_tree.heading("timestamp", text="Waktu Kirim", anchor=tk.CENTER)
        self.history_tree.heading("sender", text="Akun Pengirim", anchor=tk.W)
        self.history_tree.heading("provinsi", text="Provinsi", anchor=tk.W)
        self.history_tree.heading("recipient", text="Penerima (To)", anchor=tk.W)
        self.history_tree.heading("subject", text="Subjek Email", anchor=tk.W)
        self.history_tree.heading("status", text="Status", anchor=tk.CENTER)
        self.history_tree.heading("type", text="Tipe", anchor=tk.CENTER)

        self.history_tree.column("id", width=45, minwidth=35, anchor=tk.CENTER, stretch=False)
        self.history_tree.column("timestamp", width=135, minwidth=110, anchor=tk.CENTER, stretch=False)
        self.history_tree.column("sender", width=180, minwidth=140, anchor=tk.W, stretch=False)
        self.history_tree.column("provinsi", width=150, minwidth=120, anchor=tk.W, stretch=False)
        self.history_tree.column("recipient", width=180, minwidth=140, anchor=tk.W, stretch=False)
        self.history_tree.column("subject", width=250, minwidth=180, anchor=tk.W, stretch=True)
        self.history_tree.column("status", width=95, minwidth=80, anchor=tk.CENTER, stretch=False)
        self.history_tree.column("type", width=85, minwidth=70, anchor=tk.CENTER, stretch=False)

        hist_yscroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        hist_xscroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=hist_yscroll.set, xscrollcommand=hist_xscroll.set)

        hist_yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        hist_xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.history_tree.tag_configure("SUCCESS", foreground="#2E7D32")
        self.history_tree.tag_configure("FAILED", foreground="#C62828")
        self.history_tree.tag_configure("ERROR", foreground="#C62828")
        self.history_tree.tag_configure("SKIPPED", foreground="#EF6C00")

        self.history_tree.bind("<<TreeviewSelect>>", self._on_history_row_selected)
        self.history_tree.bind("<Double-1>", lambda e: self._open_history_in_browser())

        hist_bar = ttk.Frame(table_lf)
        hist_bar.pack(fill=tk.X, pady=(4, 0))
        self.hist_count_lbl = ttk.Label(hist_bar, text="Belum ada data riwayat.", font=("Segoe UI", 8, "italic"), foreground="#555555")
        self.hist_count_lbl.pack(side=tk.LEFT)
        ttk.Label(hist_bar, text="💡 Klik satu baris untuk melihat isi email, atau klik dua kali untuk buka di browser.", font=("Segoe UI", 8), foreground="#1565C0").pack(side=tk.RIGHT)

        # --- Bottom Pane: Detail & Body Viewer ---
        bottom_pane = ttk.Frame(paned)
        paned.add(bottom_pane, weight=5)

        detail_lf = ttk.LabelFrame(bottom_pane, text=" ✉️ Pratinjau & Detail Isi Email Terkirim ", padding=8)
        detail_lf.pack(fill=tk.BOTH, expand=True)

        # Top Metadata Box (grid)
        meta_card = ttk.Frame(detail_lf, relief=tk.GROOVE, borderwidth=1, padding=6)
        meta_card.pack(fill=tk.X, pady=(0, 6))
        meta_card.columnconfigure(1, weight=1)
        meta_card.columnconfigure(3, weight=1)

        # Row 0: Pengirim, To
        ttk.Label(meta_card, text="Pengirim:", font=("Segoe UI", 8, "bold"), foreground="#1565C0").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        self.hist_det_sender = ttk.Label(meta_card, text="-", font=("Segoe UI", 8))
        self.hist_det_sender.grid(row=0, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(meta_card, text="Penerima (To):", font=("Segoe UI", 8, "bold"), foreground="#1565C0").grid(row=0, column=2, sticky=tk.W, padx=4, pady=2)
        self.hist_det_to = ttk.Label(meta_card, text="-", font=("Segoe UI", 8))
        self.hist_det_to.grid(row=0, column=3, sticky=tk.W, padx=4, pady=2)

        # Row 1: Cc / Bcc, Waktu
        ttk.Label(meta_card, text="Tembusan:", font=("Segoe UI", 8, "bold"), foreground="#1565C0").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2)
        self.hist_det_cc = ttk.Label(meta_card, text="-", font=("Segoe UI", 8))
        self.hist_det_cc.grid(row=1, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(meta_card, text="Waktu Kirim:", font=("Segoe UI", 8, "bold"), foreground="#1565C0").grid(row=1, column=2, sticky=tk.W, padx=4, pady=2)
        self.hist_det_time = ttk.Label(meta_card, text="-", font=("Segoe UI", 8))
        self.hist_det_time.grid(row=1, column=3, sticky=tk.W, padx=4, pady=2)

        # Row 2: Subjek, Status Badge
        ttk.Label(meta_card, text="Subjek:", font=("Segoe UI", 8, "bold"), foreground="#1565C0").grid(row=2, column=0, sticky=tk.W, padx=4, pady=2)
        self.hist_det_subj = ttk.Label(meta_card, text="-", font=("Segoe UI", 8, "bold"))
        self.hist_det_subj.grid(row=2, column=1, sticky=tk.W, padx=4, pady=2)

        ttk.Label(meta_card, text="Status:", font=("Segoe UI", 8, "bold"), foreground="#1565C0").grid(row=2, column=2, sticky=tk.W, padx=4, pady=2)
        self.hist_det_status = ttk.Label(meta_card, text="-", font=("Segoe UI", 8, "bold"), foreground="#555555")
        self.hist_det_status.grid(row=2, column=3, sticky=tk.W, padx=4, pady=2)

        # Row 3: Lampiran
        ttk.Label(meta_card, text="Lampiran:", font=("Segoe UI", 8, "bold"), foreground="#1565C0").grid(row=3, column=0, sticky=tk.W, padx=4, pady=2)
        self.hist_det_att = ttk.Label(meta_card, text="-", font=("Segoe UI", 8))
        self.hist_det_att.grid(row=3, column=1, columnspan=3, sticky=tk.W, padx=4, pady=2)

        # Action bar on top of body viewer
        det_act_bar = ttk.Frame(detail_lf)
        det_act_bar.pack(fill=tk.X, pady=(0, 4))

        self.btn_hist_browser = self._create_primary_button(
            det_act_bar,
            text="🌐 Buka Tampilan Visual di Browser",
            command=self._open_history_in_browser,
        )
        self.btn_hist_browser.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(det_act_bar, text="📋 Salin Isi Email", command=self._copy_history_body).pack(side=tk.LEFT, padx=(0, 6))

        self.hist_hint_lbl = ttk.Label(
            det_act_bar,
            text="(Pilih baris pada tabel di atas untuk memuat rincian isi email)",
            font=("Segoe UI", 8, "italic"),
            foreground="#666666",
        )
        self.hist_hint_lbl.pack(side=tk.LEFT, padx=(8, 0))

        # Body Content ScrolledText
        self.hist_body_text = scrolledtext.ScrolledText(
            detail_lf,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#FAFAFA",
            fg="#212121",
        )
        self.hist_body_text.pack(fill=tk.BOTH, expand=True)
        self.hist_body_text.insert(tk.END, "Belum ada email yang dipilih.")
        self.hist_body_text.config(state=tk.DISABLED)

    def _on_history_filter_change(self, event=None):
        self._refresh_history_ui()

    def _reset_history_filter(self):
        self.hist_sender_filter.set("[Semua Akun Pengirim]")
        self.hist_status_filter.set("[Semua Status]")
        self.hist_search_var.set("")
        self._refresh_history_ui()

    def _refresh_history_ui(self):
        """Reload distinct senders, query database, and refresh Treeview table."""
        if not hasattr(self, "history_tree"):
            return

        # 1. Update sender dropdown choices
        try:
            senders = self.history_manager.get_distinct_senders()
            sender_opts = ["[Semua Akun Pengirim]"] + senders
            self.combo_hist_sender["values"] = sender_opts
            cur_val = self.hist_sender_filter.get()
            if cur_val not in sender_opts:
                self.hist_sender_filter.set("[Semua Akun Pengirim]")
        except Exception:
            pass

        # 2. Query history records
        sender_val = self.hist_sender_filter.get()
        status_val = self.hist_status_filter.get()
        search_val = self.hist_search_var.get().strip()

        try:
            rows = self.history_manager.get_history(
                sender_filter=sender_val,
                status_filter=status_val,
                search_query=search_val,
                include_dry_run=True,
                limit=3000,
            )
        except Exception as e:
            rows = []
            print(f"[WARN] Gagal membaca riwayat: {e}")

        # 3. Populate Treeview
        self.history_tree.delete(*self.history_tree.get_children())
        for r in rows:
            is_dry = bool(r.get("is_dry_run"))
            type_badge = "Simulasi" if is_dry else "Riil"

            stat_raw = (r.get("status") or "").upper()
            if stat_raw == "SUCCESS":
                stat_badge = "✓ Berhasil"
                tag_name = "SUCCESS"
            elif stat_raw in ["FAILED", "ERROR"]:
                stat_badge = "✗ Gagal"
                tag_name = "FAILED"
            else:
                stat_badge = "Dilewati"
                tag_name = "SKIPPED"

            prov_text = f"{r.get('provinsi_kode', '')} - {r.get('provinsi_nama', '')}".strip(" -")
            row_id = str(r["id"])
            self.history_tree.insert(
                "",
                tk.END,
                iid=row_id,
                values=(
                    r["id"],
                    r.get("timestamp", ""),
                    r.get("sender_email", ""),
                    prov_text or "-",
                    r.get("recipient_to", ""),
                    r.get("subject", ""),
                    stat_badge,
                    type_badge,
                ),
                tags=(tag_name,),
            )

        self.hist_count_lbl.config(text=f"Total: {len(rows)} riwayat ditemukan.")

        # Re-select previous ID if still present, else clear detail
        if self.selected_history_id and self.history_tree.exists(str(self.selected_history_id)):
            self.history_tree.selection_set(str(self.selected_history_id))
            self.history_tree.see(str(self.selected_history_id))
            self._on_history_row_selected()
        elif rows:
            first_id = str(rows[0]["id"])
            self.history_tree.selection_set(first_id)
            self._on_history_row_selected()
        else:
            self.selected_history_id = None
            self._clear_history_detail()

    def _clear_history_detail(self):
        self.hist_det_sender.config(text="-")
        self.hist_det_to.config(text="-")
        self.hist_det_cc.config(text="-")
        self.hist_det_time.config(text="-")
        self.hist_det_subj.config(text="-")
        self.hist_det_att.config(text="-")
        self.hist_det_status.config(text="-", foreground="#555555")
        self.hist_hint_lbl.config(text="(Belum ada riwayat pengiriman yang dipilih)")

        self.hist_body_text.config(state=tk.NORMAL)
        self.hist_body_text.delete("1.0", tk.END)
        self.hist_body_text.insert(tk.END, "Tidak ada data riwayat yang dipilih.")
        self.hist_body_text.config(state=tk.DISABLED)

    def _on_history_row_selected(self, event=None):
        selected = self.history_tree.selection()
        if not selected:
            return
        rec_id = int(selected[0])
        self.selected_history_id = rec_id
        rec = self.history_manager.get_record_by_id(rec_id)
        if not rec:
            return

        self.hist_det_sender.config(text=rec.get("sender_email", "-"))
        self.hist_det_to.config(text=rec.get("recipient_to", "-"))

        cc_bcc_list = []
        if rec.get("recipient_cc"):
            cc_bcc_list.append(f"Cc: {rec['recipient_cc']}")
        if rec.get("recipient_bcc"):
            cc_bcc_list.append(f"Bcc: {rec['recipient_bcc']}")
        self.hist_det_cc.config(text=" | ".join(cc_bcc_list) if cc_bcc_list else "-")

        self.hist_det_time.config(text=rec.get("timestamp", "-"))
        self.hist_det_subj.config(text=rec.get("subject", "-"))
        self.hist_det_att.config(text=rec.get("attachments", "-") or "Tidak ada lampiran")

        is_dry = bool(rec.get("is_dry_run"))
        stat_raw = (rec.get("status") or "").upper()
        if is_dry:
            status_text = "SIMULASI (DRY RUN)"
            status_color = "#E65100"
        elif stat_raw == "SUCCESS":
            status_text = "✓ BERHASIL DIKIRIM (SUCCESS)"
            status_color = "#2E7D32"
        else:
            err = rec.get("error_message")
            status_text = f"✗ GAGAL: {err}" if err else "✗ GAGAL (FAILED)"
            status_color = "#C62828"

        self.hist_det_status.config(text=status_text, foreground=status_color)
        self.hist_hint_lbl.config(text=f"Menampilkan detail ID #{rec['id']} ({rec.get('provinsi_nama', '')})")

        body_content = rec.get("body_html") or rec.get("body_text", "")
        self.hist_body_text.config(state=tk.NORMAL)
        self.hist_body_text.delete("1.0", tk.END)
        self.hist_body_text.insert(tk.END, body_content or "(Konten body kosong)")
        self.hist_body_text.config(state=tk.DISABLED)

    def _open_history_in_browser(self):
        if not self.selected_history_id:
            messagebox.showinfo("Pilih Riwayat", "Silakan klik salah satu baris riwayat email terlebih dahulu.")
            return
        rec = self.history_manager.get_record_by_id(self.selected_history_id)
        if not rec:
            messagebox.showwarning("Data Tidak Ditemukan", "Data riwayat email tidak ditemukan di database.")
            return

        subject = rec.get("subject", "Tanpa Subjek")
        sender = rec.get("sender_email", "-")
        to_email = rec.get("recipient_to", "-")
        cc_email = rec.get("recipient_cc", "")
        bcc_email = rec.get("recipient_bcc", "")
        provinsi = f"{rec.get('provinsi_kode', '')} - {rec.get('provinsi_nama', '')}".strip(" -")
        timestamp = rec.get("timestamp", "-")
        attachments = rec.get("attachments", "")
        body_content = rec.get("body_html", "") or rec.get("body_text", "")

        is_dry = bool(rec.get("is_dry_run"))
        is_success = rec.get("status") == "SUCCESS"
        badge_class = "badge-dryrun" if is_dry else ("badge-success" if is_success else "badge-failed")
        badge_label = "SIMULASI (DRY RUN)" if is_dry else ("BERHASIL (SUCCESS)" if is_success else "GAGAL (FAILED)")

        if "<html" not in body_content.lower() and "<body" not in body_content.lower() and "<p" not in body_content.lower():
            body_rendered = f"<pre style='font-family: Segoe UI, sans-serif; white-space: pre-wrap; font-size: 14px;'>{body_content}</pre>"
        else:
            body_rendered = body_content

        full_html = (
            "<!DOCTYPE html>\n"
            "<html>\n"
            "<head>\n"
            "<meta charset=\"utf-8\">\n"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
            f"<title>Riwayat: {subject}</title>\n"
            "<style>\n"
            "body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f8; margin: 0; padding: 25px; }\n"
            ".email-card { background: #ffffff; max-width: 820px; margin: 0 auto; border-radius: 8px; box-shadow: 0 4px 14px rgba(0,0,0,0.12); overflow: hidden; }\n"
            ".email-header { background: #1565C0; color: #ffffff; padding: 18px 24px; font-size: 13px; line-height: 1.6; }\n"
            ".email-header h2 { margin: 0 0 12px 0; font-size: 19px; color: #ffffff; }\n"
            ".meta-item { margin-bottom: 4px; }\n"
            ".meta-tag { color: #BBDEFB; font-weight: bold; width: 110px; display: inline-block; }\n"
            ".badge { display: inline-block; padding: 3px 10px; border-radius: 4px; font-size: 11px; font-weight: bold; }\n"
            ".badge-success { background: #2E7D32; color: #fff; }\n"
            ".badge-failed { background: #C62828; color: #fff; }\n"
            ".badge-dryrun { background: #E65100; color: #fff; }\n"
            ".email-body { padding: 30px; min-height: 250px; line-height: 1.65; color: #212121; font-size: 14px; }\n"
            ".att-container { margin: 20px 30px; padding: 12px 18px; background: #ECEFF1; border-left: 4px solid #1565C0; border-radius: 4px; font-size: 13px; color: #37474F; }\n"
            ".footer-note { background: #FAFAFA; border-top: 1px solid #E0E0E0; padding: 12px 30px; font-size: 12px; color: #757575; text-align: right; }\n"
            "</style>\n"
            "</head>\n"
            "<body>\n"
            "<div class=\"email-card\">\n"
            "  <div class=\"email-header\">\n"
            f"    <h2>{subject}</h2>\n"
            f"    <div class=\"meta-item\"><span class=\"meta-tag\">Pengirim:</span> <b>{sender}</b></div>\n"
            f"    <div class=\"meta-item\"><span class=\"meta-tag\">Penerima (To):</span> {to_email}</div>\n"
            + (f"    <div class=\"meta-item\"><span class=\"meta-tag\">Tembusan (Cc):</span> {cc_email}</div>\n" if cc_email else "")
            + (f"    <div class=\"meta-item\"><span class=\"meta-tag\">Blind Cc (Bcc):</span> {bcc_email}</div>\n" if bcc_email else "")
            + f"    <div class=\"meta-item\"><span class=\"meta-tag\">Provinsi:</span> {provinsi or '-'}</div>\n"
            + f"    <div class=\"meta-item\"><span class=\"meta-tag\">Waktu Kirim:</span> {timestamp} &nbsp;|&nbsp; Status: <span class=\"badge {badge_class}\">{badge_label}</span></div>\n"
            "  </div>\n"
            "  <div class=\"email-body\">\n"
            f"{body_rendered}\n"
            "  </div>\n"
            + (f"  <div class=\"att-container\">📎 <b>Berkas Lampiran:</b> {attachments}</div>\n" if attachments else "")
            + "  <div class=\"footer-note\">Dicatat oleh Automated Provincial Email Dispatcher - BPS RI</div>\n"
            "</div>\n"
            "</body>\n"
            "</html>\n"
        )

        try:
            with tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8") as tf:
                tf.write(full_html)
                tmp_name = tf.name
            webbrowser.open("file://" + os.path.abspath(tmp_name))
        except Exception as e:
            messagebox.showerror("Gagal Membuka Browser", f"Gagal membuka browser:\n{e}")

    def _copy_history_body(self):
        if not self.selected_history_id:
            messagebox.showinfo("Pilih Riwayat", "Silakan klik salah satu baris riwayat email terlebih dahulu.")
            return
        rec = self.history_manager.get_record_by_id(self.selected_history_id)
        if not rec:
            return
        content = rec.get("body_html") or rec.get("body_text", "")
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        messagebox.showinfo("Berhasil Disalin", "Isi pesan email berhasil disalin ke clipboard.")

    def _export_history_excel(self):
        sender_val = self.hist_sender_filter.get()
        status_val = self.hist_status_filter.get()
        search_val = self.hist_search_var.get().strip()

        default_name = f"Riwayat_Pengiriman_BPS_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path = filedialog.asksaveasfilename(
            title="Ekspor Riwayat Pengiriman Email ke Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Files (*.xlsx)", "*.xlsx"), ("All Files (*.*)", "*.*")],
        )
        if not out_path:
            return

        try:
            count = self.history_manager.export_to_excel(
                output_path=out_path,
                sender_filter=sender_val,
                status_filter=status_val,
                search_query=search_val,
            )
            messagebox.showinfo(
                "Ekspor Berhasil",
                f"Berhasil mengekspor {count} data riwayat ke berkas:\n{out_path}",
            )
        except Exception as e:
            messagebox.showerror("Gagal Ekspor", f"Terjadi kesalahan saat mengekspor ke Excel:\n{str(e)}")

    def _confirm_clear_history(self):
        sender_val = self.hist_sender_filter.get()
        is_specific = bool(sender_val and sender_val != "[Semua Akun Pengirim]")
        target_desc = f"khusus untuk akun pengirim '{sender_val}'" if is_specific else "SEMUA akun pengirim"

        confirm = messagebox.askyesno(
            "Konfirmasi Hapus Riwayat",
            f"Apakah Anda yakin ingin menghapus seluruh riwayat pengiriman {target_desc}?\n\n"
            f"Tindakan ini permanen dan tidak dapat dibatalkan.",
            icon=messagebox.WARNING,
        )
        if not confirm:
            return

        try:
            deleted = self.history_manager.clear_history(sender_filter=sender_val if is_specific else None)
            messagebox.showinfo("Riwayat Dihapus", f"Berhasil menghapus {deleted} catatan riwayat.")
            self._refresh_history_ui()
        except Exception as e:
            messagebox.showerror("Gagal Menghapus", f"Gagal menghapus riwayat:\n{str(e)}")

    def _create_metric_card(self, parent, title: str, init_val: str, color: str):
        frame = ttk.Frame(parent, relief=tk.GROOVE, borderwidth=1, padding=8)
        lbl_title = ttk.Label(frame, text=title, font=("Segoe UI", 8), foreground="#555555")
        lbl_title.pack(anchor=tk.CENTER)
        lbl_val = ttk.Label(frame, text=init_val, font=("Segoe UI", 16, "bold"), foreground=color)
        lbl_val.pack(anchor=tk.CENTER)
        frame.val_label = lbl_val  # attach reference
        return frame

    def _build_statusbar(self):
        self.status_bar = ttk.Frame(self.root, relief=tk.SUNKEN, padding=(10, 3))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.status_lbl = ttk.Label(self.status_bar, text="Siap.", font=("Segoe UI", 8))
        self.status_lbl.pack(side=tk.LEFT)

        version_lbl = ttk.Label(self.status_bar, text="v1.0 (Portable) | BPS RI", font=("Segoe UI", 8), foreground="#777777")
        version_lbl.pack(side=tk.RIGHT)

    # --- ACTION HANDLERS ---
    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Pilih File Database Kontak BPS",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if path:
            self.excel_path.set(path)
            self._load_excel_data()

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Pilih Folder Lampiran Berkas")
        if folder:
            self.custom_attachment_files = []
            display_text = f"{folder} (Semua berkas di folder)"
            self.folder_path.set(display_text)
            self._scan_attachment_folder(actual_path=folder)

    def _browse_files(self):
        files = filedialog.askopenfilenames(
            title="Pilih Berkas Lampiran (Tahan tombol Ctrl / Shift untuk Memilih Banyak Sekaligus)",
            filetypes=[
                ("Semua Berkas", "*.*"),
                ("Excel & Dokumen Office", "*.xlsx;*.xls;*.pdf;*.docx;*.zip;*.csv"),
            ],
        )
        if files:
            # Handle splitlist if tkinter returns a formatted string
            if isinstance(files, str):
                try:
                    file_list = list(self.root.tk.splitlist(files))
                except Exception:
                    file_list = [files]
            else:
                file_list = list(files)

            # Append new files into custom_attachment_files (deduplicate)
            for f in file_list:
                f_norm = os.path.normpath(f)
                if f_norm not in self.custom_attachment_files:
                    self.custom_attachment_files.append(f_norm)

            count = len(self.custom_attachment_files)
            names = [os.path.basename(f) for f in self.custom_attachment_files]
            if count == 1:
                display_text = names[0]
            elif count == 2:
                display_text = f"{names[0]}, {names[1]}"
            else:
                display_text = f"[{count} Berkas Terpilih]: {names[0]}, {names[1]}, dst."

            self.folder_path.set(display_text)
            self._scan_custom_files()

    def _reset_attachments(self):
        """Clear all selected attachment files/folder."""
        self.custom_attachment_files = []
        self.folder_path.set("")
        self.attachment_engine = None
        self.row_checked_state.clear()
        self.folder_status_lbl.config(
            text="Belum ada folder atau berkas lampiran yang dipilih.",
            style="SubHeader.TLabel",
        )
        self._refresh_preview()

    def _load_excel_data(self):
        path = self.excel_path.get().strip()
        if not path or not os.path.exists(path):
            return

        try:
            parser = BPSExcelParser(path)
            records = parser.load()
            self.records = records
            self.email_columns = parser.email_columns
            self.dynamic_tags = parser.dynamic_tags
            self._update_tag_ui()
            self.row_checked_state.clear()

            valid_count = sum(1 for r in records if r["is_valid"])
            invalid_count = len(records) - valid_count

            if invalid_count > 0:
                self.excel_status_lbl.config(
                    text=f"✓ Terbaca {len(records)} provinsi ({valid_count} valid, {invalid_count} ada peringatan)",
                    style="BadgeWarn.TLabel",
                )
            else:
                self.excel_status_lbl.config(
                    text=f"✓ Berhasil memuat {len(records)} provinsi siap kirim.",
                    style="Badge.TLabel",
                )
            self._refresh_preview()
        except Exception as e:
            self.records = []
            self.excel_status_lbl.config(text=f"✗ Gagal membaca Excel: {str(e)}", style="BadgeWarn.TLabel")
            messagebox.showerror("Kesalahan Excel", f"Gagal membaca file Excel:\n{str(e)}")

    def _scan_attachment_folder(self, actual_path: Optional[str] = None):
        path = actual_path
        if not path:
            raw_path = self.folder_path.get().strip()
            if " (Semua berkas di folder)" in raw_path:
                path = raw_path.replace(" (Semua berkas di folder)", "").strip()
            else:
                path = raw_path

        if not path or not os.path.isdir(path):
            return

        self.custom_attachment_files = []
        self.attachment_engine = AttachmentEngine(folder_path=path)
        self._finish_attachment_scan()

    def _scan_custom_files(self):
        if not self.custom_attachment_files:
            return
        self.attachment_engine = AttachmentEngine(files=self.custom_attachment_files)
        self._finish_attachment_scan()

    def _finish_attachment_scan(self):
        if not self.attachment_engine:
            return

        total_files, matched_count = self.attachment_engine.scan()

        if self.attachment_engine.scan_error:
            self.folder_status_lbl.config(
                text=f"✗ {self.attachment_engine.scan_error}",
                style="BadgeWarn.TLabel",
            )
        elif total_files > 0 and matched_count == 0:
            self.folder_status_lbl.config(
                text=f"⚠️ {total_files} file ditemukan, tetapi TIDAK ADA yang diawali kode provinsi (misal: '11', '12') atau folder kode provinsi.",
                style="BadgeWarn.TLabel",
            )
            unmatched_names = [os.path.basename(f) for f in self.attachment_engine.unmatched_files[:3]]
            unmatched_str = ", ".join([f"'{n}'" for n in unmatched_names])
            messagebox.showwarning(
                "Format Nama File / Folder Tidak Cocok",
                f"Sistem menemukan {total_files} file di folder:\n"
                f"• {unmatched_str}...\n\n"
                f"Namun tidak ada file atau folder yang cocok dengan format kode wilayah BPS:\n"
                f"1. Awalan/akhiran '00' atau kata 'Nasional' (misal: '00_Tabel.xlsx', '00_Naskah.pdf', atau folder '00_Nasional/')\n"
                f"2. Memuat kode provinsi bebas di awal/akhir (misal: '11_Laporan.pdf', 'Tabel Profil IMK - 16.xlsx', atau folder '11/')\n\n"
                f"Silakan sesuaikan nama berkas/folder atau pilih berkas dengan tombol '📄 + Pilih / Tambah File...'.",
            )
        else:
            nat_cnt = len(self.attachment_engine.national_files)
            prov_cnt = len(self.attachment_engine.provincial_files)
            matched_codes = sorted(list(self.attachment_engine.provincial_files.keys()))
            if matched_codes:
                sample_codes = ", ".join(matched_codes[:6]) + ("..." if len(matched_codes) > 6 else "")
                detail_str = f"{prov_cnt} provinsi cocok ({sample_codes})"
            else:
                detail_str = "0 provinsi cocok"
            self.folder_status_lbl.config(
                text=f"✓ {total_files} file ditemukan ({nat_cnt} berkas nasional '00', {detail_str})",
                style="Badge.TLabel",
            )
        self.row_checked_state.clear()
        self._refresh_preview()

    def _toggle_pass_visibility(self):
        if self.smtp_show_pass.get():
            self.pass_entry.config(show="")
        else:
            self.pass_entry.config(show="*")

    def _insert_tag(self, tag: str):
        """Insert a dynamic tag placeholder at current cursor position in subject or body."""
        target = getattr(self, "last_focused_text_widget", self.body_text)
        try:
            target.insert(tk.INSERT, tag)
            target.focus_set()
        except Exception:
            try:
                self.body_text.insert(tk.INSERT, tag)
                self.body_text.focus_set()
            except Exception:
                pass

    def _update_tag_ui(self):
        """Update dynamic tag guidance label and clickable buttons based on detected Excel columns."""
        tags = getattr(self, "dynamic_tags", [])
        if not tags:
            tags = ["kode", "provinsi", "kepala_prov", "nama_ketua", "nama_koordinator", "nama_staf", "no"]

        # Clean standard order without duplicates:
        # User requested: remove duplicate (ketua), (staf), (koordinator), and add (kepala prov)
        standard_keys = [
            "kode",
            "provinsi",
            "kepala_prov",
            "nama_ketua",
            "nama_koordinator",
            "nama_staf",
            "no",
        ]

        duplicate_to_omit = {
            "ketua", "koordinator", "koor", "staff", "staf", "nama_staff",
            "kepala", "nama_kepala", "nama_provinsi", "kode_provinsi", "kode_provisi",
            "provisi", "nomor", "nama_tim_prod", "nama_pjk"
        }

        display_tags = []
        for sk in standard_keys:
            if sk not in display_tags:
                display_tags.append(sk)

        # Append any extra custom columns from Excel file (excluding duplicate aliases)
        for t in tags:
            t_clean = t.strip().lower()
            if t_clean not in display_tags and t_clean not in duplicate_to_omit:
                display_tags.append(t_clean)

        tag_placeholders = [f"{{{{{t}}}}}" for t in display_tags]
        guide_text = "Tag dinamis tersedia: " + ", ".join(tag_placeholders) + " (juga mendukung {{nama_kepala}})"
        if hasattr(self, "tag_guide_lbl"):
            self.tag_guide_lbl.config(text=guide_text)

        if hasattr(self, "tag_buttons_frame"):
            for w in self.tag_buttons_frame.winfo_children():
                w.destroy()

            max_cols = 4
            for idx, tag in enumerate(display_tags):
                r = idx // max_cols
                c = idx % max_cols
                tag_str = f"{{{{{tag}}}}}"
                btn = tk.Button(
                    self.tag_buttons_frame,
                    text=f"+ {tag_str}",
                    font=("Segoe UI", 8, "bold"),
                    bg="#E3F2FD",
                    fg="#0D47A1",
                    activebackground="#BBDEFB",
                    activeforeground="#0D47A1",
                    relief=tk.GROOVE,
                    bd=1,
                    padx=6,
                    pady=3,
                    cursor="hand2",
                    command=lambda t=tag_str: self._insert_tag(t),
                )
                btn.grid(row=r, column=c, padx=3, pady=2, sticky="ew")

            for col_i in range(max_cols):
                self.tag_buttons_frame.columnconfigure(col_i, weight=1)

    def _sync_sender_with_username(self):
        u = self.smtp_username.get().strip()
        if u:
            self.sender_email.set(u)
            self._save_config(show_popup=False)
            messagebox.showinfo(
                "Email Pengirim Diselaraskan",
                f"Alamat Email Pengirim (From) berhasil disamakan dengan Username:\n{u}",
                parent=self.root,
            )
        else:
            messagebox.showwarning(
                "Username Kosong",
                "Silakan isi kolom Username SMTP / Login terlebih dahulu.",
                parent=self.root,
            )

    def _test_smtp(self):
        host = self.smtp_host.get().strip()
        port_str = self.smtp_port.get().strip()
        username = self.smtp_username.get().strip()
        password = self.smtp_pass.get().strip()
        sender = self.sender_email.get().strip()

        if not host:
            messagebox.showwarning(
                "Host Kosong",
                "Silakan masukkan Host SMTP (contoh: mail.bps.go.id atau smtp.gmail.com) terlebih dahulu."
            )
            return

        try:
            port = int(port_str) if port_str else 587
        except ValueError:
            messagebox.showerror("Port Salah", "Port SMTP harus berupa angka (misal 587, 465, atau 25).")
            return

        # Auto-save current config
        self._save_config(show_popup=False)

        self.smtp_result_lbl.config(text="Menguji koneksi...", foreground="#1565C0")
        self.root.update_idletasks()

        def run_test():
            ok, msg = test_smtp_connection(host, port, username, password, timeout=self.smtp_timeout.get())
            self.root.after(0, lambda: self._on_smtp_test_result(ok, msg, host, username, sender))

        threading.Thread(target=run_test, daemon=True).start()

    def _on_smtp_test_result(self, ok: bool, msg: str, host: str, username: str, sender: str):
        if ok:
            is_compat, warnings, recommended = check_smtp_sender_compatibility(host, username, sender)
            if is_compat:
                self.smtp_result_lbl.config(text="✓ Terhubung & Selaras!", foreground="#2E7D32")
                messagebox.showinfo(
                    "Koneksi & Konfigurasi Sempurna",
                    f"{msg}\n\n"
                    f"✓ Konfigurasi Pengirim Selaras:\n"
                    f"Email Pengirim ('{sender}') telah sesuai dengan Host/Akun {host}.\n\n"
                    f"Aplikasi siap melakukan pengiriman email riil.",
                    parent=self.root,
                )
            else:
                self.smtp_result_lbl.config(text="⚠️ Terhubung, Pengirim Tidak Selaras!", foreground="#E65100")
                warn_text = "\n\n• ".join(warnings)
                recom_tip = f"\n\n💡 Rekomendasi Solusi:\nUbah 'Email Pengirim (From)' menjadi: '{recommended}'\natau klik tombol '📋 Samakan dg Username'." if recommended else ""
                messagebox.showwarning(
                    "⚠️ Peringatan: Konfigurasi Pengirim Berbeda",
                    f"✓ Login & Koneksi ke server SMTP BERHASIL!\n\n"
                    f"⚠️ NAMUN ADA PERINGATAN KETIDAKCOCOKAN:\n• {warn_text}{recom_tip}\n\n"
                    f"PENTING: Jangan terkecoh dengan koneksi yang berhasil. Jika Email Pengirim tidak selaras dengan Host/Username, email kemungkinan besar akan DITOLAK server atau TIDAK SAMPAI ke inbox penerima.",
                    parent=self.root,
                )
        else:
            self.smtp_result_lbl.config(text="✗ Gagal Terhubung", foreground="#C62828")
            messagebox.showerror("Koneksi Gagal", msg, parent=self.root)

    def _configure_tree_columns(self):
        """Configure treeview columns dynamically to match email and name columns found in the loaded Excel file."""
        base_cols = [
            ("check", "[✓] Kirim", 65, tk.CENTER),
            ("no", "No", 40, tk.CENTER),
            ("kode", "Kode", 50, tk.CENTER),
            ("provinsi", "Provinsi", 140, tk.W),
        ]

        contact_cols = []
        if self.email_columns:
            for c in self.email_columns:
                # 1. Add paired Name column if available
                if c.get("name_key") and c.get("name_col_idx", -1) != -1:
                    n_key = c["name_key"]
                    n_label = c.get("name_label") or f"{c['role']} - Nama"
                    calc_n_width = max(190, len(n_label) * 8)
                    contact_cols.append((n_key, n_label, calc_n_width, tk.W))

                # 2. Add Email column with full width so headers are never truncated
                e_key = c["key"]
                e_label = c["label"]
                calc_e_width = max(260, len(e_label) * 9)
                contact_cols.append((e_key, e_label, calc_e_width, tk.W))
        else:
            contact_cols = [
                ("nama_kepala", "Kepala Prov - Nama", 190, tk.W),
                ("email_kepala", "Kepala Prov - E-mail [TO]", 260, tk.W),
                ("nama_ketua", "Ketua Tim Statistik Produksi - Nama", 230, tk.W),
                ("email_pjk", "Ketua Tim Statistik Produksi - Email [CC]", 270, tk.W),
            ]

        tail_cols = [
            ("national_att", "Nasional (00)", 95, tk.CENTER),
            ("prov_att", "Daerah", 70, tk.CENTER),
            ("total_att", "Total Berkas", 85, tk.CENTER),
            ("status", "Status Validasi", 160, tk.W),
        ]

        all_cols = base_cols + contact_cols + tail_cols
        col_ids = [c[0] for c in all_cols]

        self.tree["columns"] = col_ids
        for cid, heading, width, anchor in all_cols:
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, minwidth=max(35, width - 30), anchor=anchor, stretch=False)

    def _go_to_preview(self):
        self._refresh_preview()
        self.notebook.select(self.tab_preview)

    def _refresh_preview(self):
        # Configure columns dynamically according to detected email columns
        self._configure_tree_columns()

        # Clear treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.records:
            if self.excel_path.get().strip():
                self._load_excel_data()
            if not self.records:
                self.preview_stats_lbl.config(text="Pilih file Excel terlebih dahulu pada Tab 1.")
                return

        # Ensure attachment engine is ready
        if not self.attachment_engine and self.folder_path.get().strip():
            self._scan_attachment_folder()

        ready_count = 0
        warning_count = 0
        error_count = 0

        for idx, rec in enumerate(self.records, start=1):
            kode = rec["kode"]
            prov = rec["provinsi"]
            to_email = (rec.get("to_email") or rec.get("email_kepala") or "").strip()
            cc_email = (rec.get("cc_email") or rec.get("email_pjk") or "").strip()
            bcc_email = (rec.get("bcc_email") or "").strip()

            nat_count = 0
            prov_count = 0
            total_att = 0

            if self.attachment_engine:
                atts = self.attachment_engine.get_attachments_for_province(kode)
                total_att = len(atts)
                nat_count = len(self.attachment_engine.national_files)
                prov_count = len(self.attachment_engine.provincial_files.get(kode, []))

            # Validation logic
            tag = "valid"
            status = "Siap Kirim"

            if not rec.get("is_valid", True):
                tag = "error"
                status = "Error: " + "; ".join(rec.get("errors", []))
                error_count += 1
            elif not to_email:
                tag = "error"
                status = "Error: Seluruh kolom email kosong"
                error_count += 1
            elif total_att == 0 and self.attachment_engine:
                tag = "warning"
                status = "Perhatian: Lampiran (0)"
                warning_count += 1
            else:
                ready_count += 1

            # Checkbox state:
            # Default behavior: Baris dengan status "Siap Kirim" (memiliki berkas nasional '00' atau daerah dan email valid) otomatis tercentang [✓]
            row_id_str = str(idx - 1)
            is_ready = (rec.get("is_valid", True) and total_att > 0 and bool(to_email))
            if row_id_str in self.row_checked_state:
                is_checked = self.row_checked_state[row_id_str]
            else:
                is_checked = is_ready
                self.row_checked_state[row_id_str] = is_checked

            check_sym = "[✓]" if is_checked else "[ ]"

            row_vals = [
                check_sym,
                idx,
                kode,
                prov,
            ]
            if self.email_columns:
                for c in self.email_columns:
                    # Name column
                    if c.get("name_key") and c.get("name_col_idx", -1) != -1:
                        name_val = rec.get(c["name_key"], "")
                        if not name_val:
                            if "kepala" in c["key"]:
                                name_val = rec.get("nama_kepala", "")
                            elif "tim" in c["key"] or "ketua" in c["key"]:
                                name_val = rec.get("nama_ketua", "")
                            elif "koor" in c["key"]:
                                name_val = rec.get("nama_koordinator", "")
                            elif "staf" in c["key"]:
                                name_val = rec.get("nama_staf", "")
                        row_vals.append(name_val if name_val else "-")

                    # Email column
                    val = rec.get(c["key"], "")
                    row_vals.append(val if val else "-")
            else:
                row_vals.append(rec.get("nama_kepala", "") or "-")
                row_vals.append(to_email or "(KOSONG)")
                row_vals.append(rec.get("nama_ketua", "") or "-")
                row_vals.append(cc_email or "-")

            row_vals.extend([
                nat_count,
                prov_count,
                total_att,
                status,
            ])

            self.tree.insert(
                "",
                tk.END,
                iid=row_id_str,
                values=tuple(row_vals),
                tags=(tag,),
            )

        self.preview_stats_lbl.config(
            text=f"Total: {len(self.records)} Provinsi | Siap: {ready_count} | Perhatian: {warning_count} | Error: {error_count}"
        )
        self._update_target_banner()

    def _get_checked_records(self) -> List[Dict[str, Any]]:
        """Return list of records currently checked [✓] in the Treeview."""
        checked = []
        for idx, rec in enumerate(self.records):
            if self.row_checked_state.get(str(idx), False):
                checked.append(rec)
        return checked

    def _toggle_checkbox(self, iid: str):
        if not iid:
            return
        cur = self.row_checked_state.get(iid, False)
        new_val = not cur
        self.row_checked_state[iid] = new_val
        vals = list(self.tree.item(iid, "values"))
        if vals:
            vals[0] = "[✓]" if new_val else "[ ]"
            self.tree.item(iid, values=vals)
        self._update_target_banner()

    def _select_all_rows(self):
        for iid in self.tree.get_children():
            self.row_checked_state[iid] = True
            vals = list(self.tree.item(iid, "values"))
            if vals:
                vals[0] = "[✓]"
                self.tree.item(iid, values=vals)
        self._update_target_banner()

    def _deselect_all_rows(self):
        for iid in self.tree.get_children():
            self.row_checked_state[iid] = False
            vals = list(self.tree.item(iid, "values"))
            if vals:
                vals[0] = "[ ]"
                self.tree.item(iid, values=vals)
        self._update_target_banner()

    def _on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            row_id = self.tree.identify_row(event.y)
            if row_id and col == "#1":  # Column #1 is 'check'
                self._toggle_checkbox(row_id)

    def _on_tree_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self._toggle_checkbox(row_id)

    def _toggle_selected_checkboxes(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        all_checked = all(self.row_checked_state.get(iid, False) for iid in selected)
        new_state = not all_checked
        for iid in selected:
            self.row_checked_state[iid] = new_state
            vals = list(self.tree.item(iid, "values"))
            if vals:
                vals[0] = "[✓]" if new_state else "[ ]"
                self.tree.item(iid, values=vals)
        self._update_target_banner()

    def _proceed_to_execution(self):
        checked = self._get_checked_records()
        if not checked:
            messagebox.showwarning(
                "Pilihan Kosong",
                "Silakan pilih/centang [✓] minimal satu baris data di tabel terlebih dahulu."
            )
            return
        self.target_records = checked
        self.target_mode_desc = f"{len(checked)} Provinsi Terpilih [✓]"
        self._update_target_banner()
        self.notebook.select(self.tab_exec)

    def _update_target_banner(self):
        """Update the target indicator banner and total card in Tab 3."""
        targets = self._get_checked_records()
        count = len(targets)
        self.target_records = targets
        self.target_mode_desc = f"{count} Provinsi Terpilih [✓]"
        if hasattr(self, "target_banner_lbl"):
            self.target_banner_lbl.config(
                text=f"Target Pengiriman: {count} Provinsi Terpilih [✓] (Menyiapkan pengiriman untuk {count} provinsi terpilih)"
            )
        if hasattr(self, "card_total"):
            self.card_total.val_label.config(text=str(count))

    def _on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return

        if len(selected) > 1:
            preview_lines = [
                f"=== SELEKSI GANDA: {len(selected)} Baris Dipilih ===",
                "Tekan tombol Spasi atau klik kolom pertama untuk mencentang/membatalkan centang [✓].\n",
                "Daftar Sasaran Terpilih:",
            ]
            for iid in selected:
                idx = int(iid)
                if idx < len(self.records):
                    r = self.records[idx]
                    nat_c = len(self.attachment_engine.national_files) if self.attachment_engine else 0
                    prov_c = len(self.attachment_engine.provincial_files.get(r["kode"], [])) if self.attachment_engine else 0
                    chk = "[✓]" if self.row_checked_state.get(iid, False) else "[ ]"
                    preview_lines.append(
                        f"  {chk} [{r['kode']}] {r['provinsi']} -> To: {r['email_kepala']} (Cc: {r['email_pjk'] or '-'}) | Total Berkas: {nat_c + prov_c}"
                    )

            self.detail_text.config(state=tk.NORMAL)
            self.detail_text.delete("1.0", tk.END)
            self.detail_text.insert(tk.END, "\n".join(preview_lines))
            self.detail_text.config(state=tk.DISABLED)
            return

        rec_idx = int(selected[0])
        if rec_idx >= len(self.records):
            return

        rec = self.records[rec_idx]
        kode = rec["kode"]

        # Prepare attachments info
        att_lines = []
        if self.attachment_engine:
            files = self.attachment_engine.get_attachments_for_province(kode)
            if files:
                for f in files:
                    sz = get_file_size_formatted(f)
                    name = os.path.basename(f)
                    category = "NASIONAL (00)" if name.startswith("00") else f"DAERAH ({kode})"
                    att_lines.append(f"  • [{category}] {name} ({sz})")
            else:
                att_lines.append("  (Tidak ada file lampiran yang cocok)")
        else:
            att_lines.append("  (Folder lampiran belum dipilih)")

        # Render preview subject and body
        subj_preview = replace_template(self.email_subject.get(), rec)
        body_preview = replace_template(self.body_text.get("1.0", tk.END).strip(), rec)
        chk_status = "TERCENTANG [✓] Siap Kirim" if self.row_checked_state.get(selected[0], False) else "TIDAK TERCANGKAU [ ]"

        # Format each email recipient cleanly
        email_recipient_lines = []
        if self.email_columns:
            for col_info in self.email_columns:
                lbl = col_info["label"]
                ekey = col_info["key"]
                nkey = col_info.get("name_key")
                em_val = rec.get(ekey, "").strip()
                nm_val = rec.get(nkey, "").strip() if nkey else ""

                if nm_val and em_val:
                    email_recipient_lines.append(f"  • {lbl:<26}: {nm_val} <{em_val}>")
                elif em_val:
                    email_recipient_lines.append(f"  • {lbl:<26}: <{em_val}>")
                elif nm_val:
                    email_recipient_lines.append(f"  • {lbl:<26}: {nm_val} (Tidak ada email / -)")
                else:
                    email_recipient_lines.append(f"  • {lbl:<26}: (Tidak ada / -)")
        else:
            to_val = rec.get("to_email") or rec.get("email_kepala", "")
            cc_val = rec.get("cc_email") or rec.get("email_pjk", "")
            email_recipient_lines.append(f"  • Kepala BPS (To)           : {rec.get('nama_kepala', '')} <{to_val or 'KOSONG'}>")
            email_recipient_lines.append(f"  • PJK / Tim (Cc)            : {rec.get('nama_ketua') or rec.get('nama_pjk', '')} <{cc_val or '-'}>")

        to_disp = rec.get("to_email") or "(KOSONG - Seluruh email baris ini kosong)"
        to_role_disp = f"[{rec.get('to_role')}]" if rec.get("to_role") else ""
        cc_disp = rec.get("cc_email") or "-"
        bcc_disp = f"{rec.get('bcc_email')} (Staf Seksi - Tersembunyi/Blind)" if rec.get("bcc_email") else "-"

        alloc_lines = [
            f"  👉 Alamat Tujuan (To)     : {to_disp} {to_role_disp}",
            f"  👉 Tembusan (Cc)          : {cc_disp}",
            f"  👉 Tembusan Buta (Bcc)    : {bcc_disp}",
        ]

        preview_content = (
            f"=== DETAIL TARGET: [{kode}] {rec['provinsi']} ({chk_status}) ===\n"
            f"--- ALOKASI PENGIRIMAN (TO / CC / BCC) ---\n"
            + "\n".join(alloc_lines)
            + "\n\n"
            f"--- DAFTAR EMAIL SUMBER EXCEL ---\n"
            + "\n".join(email_recipient_lines)
            + "\n\n"
            f"--- LAMPIRAN BERKAS ({len(att_lines)} berkas) ---\n"
            + "\n".join(att_lines)
            + "\n\n"
            f"--- SUBJEK EMAIL ---\n{subj_preview}\n\n"
            f"--- PRATINJAU ISI PESAN ---\n{body_preview}\n"
        )

        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, preview_content)
        self.detail_text.config(state=tk.DISABLED)

    # --- DISPATCH LOGIC ---
    def _start_dry_run(self):
        if not self._validate_preflight(require_smtp=False):
            return
        self.notebook.select(self.tab_exec)
        self._execute_dispatch(dry_run=True)

    def _confirm_and_start_real_send(self):
        if not self._validate_preflight(require_smtp=True):
            return

        targets = self._get_checked_records()
        valid_count = sum(1 for r in targets if r["is_valid"] and (r.get("to_email") or r.get("email_kepala")))
        confirm = messagebox.askyesno(
            "Konfirmasi Pengiriman Email",
            f"PERINGATAN PENGIRIMAN RIIL:\n\n"
            f"Menyiapkan pengiriman untuk {len(targets)} provinsi terpilih [✓].\n"
            f"Jumlah Siap Kirim : {valid_count} BPS Provinsi\n\n"
            f"• Email Pengirim (From) : {self.sender_email.get().strip()}\n"
            f"• Login SMTP (Username) : {self.smtp_username.get().strip()}\n"
            f"• Server SMTP           : {self.smtp_host.get().strip()}:{self.smtp_port.get().strip()}\n\n"
            f"Pastikan seluruh lampiran dan alamat email sasaran sudah diperiksa.\n\n"
            f"Apakah Anda yakin ingin memulai pengiriman sekarang?",
            icon=messagebox.WARNING,
        )
        if not confirm:
            return

        self.notebook.select(self.tab_exec)
        self._execute_dispatch(dry_run=False)

    def _validate_preflight(self, require_smtp: bool = True) -> bool:
        if not self.records:
            messagebox.showwarning("Data Kosong", "Silakan pilih file Excel database kontak BPS terlebih dahulu.")
            self.notebook.select(self.tab_config)
            return False

        targets = self._get_checked_records()
        if not targets:
            messagebox.showwarning(
                "Tidak Ada Sasaran Tercentang",
                "Silakan pilih/centang [✓] minimal satu baris data di Tab 2 (Preview & Verifikasi) terlebih dahulu."
            )
            self.notebook.select(self.tab_preview)
            return False

        if not self.attachment_engine or not self.attachment_engine.all_files:
            messagebox.showwarning("Folder Kosong", "Silakan pilih folder berkas lampiran yang valid.")
            self.notebook.select(self.tab_config)
            return False

        if require_smtp:
            host = self.smtp_host.get().strip()
            username = self.smtp_username.get().strip()
            sender = self.sender_email.get().strip()
            password = self.smtp_pass.get().strip()

            if not host:
                messagebox.showwarning(
                    "Host SMTP Kosong",
                    "Silakan masukkan Host SMTP (contoh: mail.bps.go.id atau smtp.bps.go.id) di Tab 1.",
                )
                self.notebook.select(self.tab_config)
                return False

            # Auto-sync sender and username if one is provided
            if not sender and username and "@" in username:
                sender = username
                self.sender_email.set(sender)
            elif not username and sender:
                username = sender
                self.smtp_username.set(username)

            if not sender:
                messagebox.showwarning(
                    "Email Pengirim Kosong",
                    "Silakan isi Email Pengirim (From) di Tab 1.",
                    parent=self.root,
                )
                self.notebook.select(self.tab_config)
                return False

            is_compat, warnings, recommended = check_smtp_sender_compatibility(host, username, sender)
            if not is_compat:
                warn_text = "\n• ".join(warnings)
                recom_tip = f"\n\n💡 Rekomendasi:\nUbah Email Pengirim menjadi '{recommended}' (atau klik '📋 Samakan dg Username' di Tab 1)." if recommended else ""
                ans = messagebox.askyesno(
                    "⚠️ Peringatan Ketidakcocokan Akun Pengirim",
                    f"Terdeteksi potensi masalah pada konfigurasi email pengirim Anda:\n\n"
                    f"• {warn_text}{recom_tip}\n\n"
                    f"Jika dilanjutkan, server SMTP berpotensi besar akan MENOLAK email ini atau masuk ke spam penerima.\n\n"
                    f"Apakah Anda yakin ingin TETAP MELANJUTKAN pengiriman email riil?",
                    icon=messagebox.WARNING,
                    parent=self.root,
                )
                if not ans:
                    self.notebook.select(self.tab_config)
                    return False

        return True

    def _execute_dispatch(self, dry_run: bool):
        if self.is_running:
            messagebox.showinfo("Sedang Berjalan", "Proses pengiriman sedang berlangsung.")
            return

        targets = self._get_checked_records()
        if not targets:
            messagebox.showwarning(
                "Tidak Ada Sasaran Tercentang",
                "Silakan pilih/centang [✓] minimal satu baris data di Tab 2 terlebih dahulu."
            )
            return

        # Reset counters
        self.card_success.val_label.config(text="0")
        self.card_skipped.val_label.config(text="0")
        self.card_error.val_label.config(text="0")

        self.progressbar["maximum"] = len(targets)
        self.progressbar["value"] = 0

        self.btn_send.config(state=tk.DISABLED, bg="#E0E0E0", fg="#9E9E9E", cursor="arrow")
        self.btn_dry_run.config(state=tk.DISABLED, bg="#E0E0E0", fg="#9E9E9E", cursor="arrow")
        self.btn_stop.config(state=tk.NORMAL, bg="#C62828", fg="#FFFFFF", cursor="hand2")
        self.is_running = True

        mode_text = "[SIMULASI / DRY RUN]" if dry_run else "[PENGIRIMAN ASLI]"
        self._log_to_console("INFO", f"=== MEMULAI {mode_text} - Menyiapkan pengiriman untuk {len(targets)} provinsi terpilih ===")

        host = self.smtp_host.get().strip()
        port = int(self.smtp_port.get().strip())
        username = self.smtp_username.get().strip()
        sender = self.sender_email.get().strip()
        pwd = self.smtp_pass.get().strip()
        delay = self.email_delay.get()

        # Save config
        self._save_config(show_popup=False)

        self.dispatcher = EmailDispatcher(
            host=host,
            port=port,
            username=username,
            password=pwd,
            delay_seconds=delay,
            sender_email=sender,
            timeout=self.smtp_timeout.get(),
            history_manager=self.history_manager,
        )

        subj_tmpl = self.email_subject.get()
        body_tmpl = self.body_text.get("1.0", tk.END).strip()

        def worker():
            stats = self.dispatcher.dispatch_batch(
                items=targets,
                subject_template=subj_tmpl,
                body_template=body_tmpl,
                attachment_engine=self.attachment_engine,
                status_callback=self._queue_status,
                progress_callback=self._queue_progress,
                dry_run=dry_run,
            )
            self.log_queue.put(("DONE", stats, None))

        threading.Thread(target=worker, daemon=True).start()

    def _queue_status(self, status: str, msg: str, record: Dict[str, Any]):
        self.log_queue.put(("LOG", (status, msg), record))

    def _queue_progress(self, current: int, total: int):
        self.log_queue.put(("PROG", (current, total), None))

    def _stop_dispatch(self):
        if self.dispatcher and self.is_running:
            self.dispatcher.cancel()
            self._log_to_console("INFO", "Mengirim sinyal pembatalan ke sistem...")
            self.btn_stop.config(state=tk.DISABLED, bg="#E0E0E0", fg="#9E9E9E", cursor="arrow")

    def _process_log_queue(self):
        """Consume messages from the worker thread safely on the main GUI thread."""
        try:
            while True:
                msg_type, data, extra = self.log_queue.get_nowait()
                if msg_type == "LOG":
                    status, text = data
                    self._log_to_console(status, text)
                elif msg_type == "PROG":
                    current, total = data
                    self.progressbar["value"] = current
                    pct = int((current / total) * 100) if total > 0 else 0
                    self.prog_lbl.config(text=f"Memproses {current} dari {total} provinsi ({pct}%)...")
                elif msg_type == "DONE":
                    stats = data
                    self.is_running = False
                    self.btn_send.config(state=tk.NORMAL, bg="#1565C0", fg="#FFFFFF", cursor="hand2")
                    self.btn_dry_run.config(state=tk.NORMAL, bg="#37474F", fg="#FFFFFF", cursor="hand2")
                    self.btn_stop.config(state=tk.DISABLED, bg="#E0E0E0", fg="#9E9E9E", cursor="arrow")

                    self.card_success.val_label.config(text=str(stats.get("success", 0)))
                    self.card_skipped.val_label.config(text=str(stats.get("skipped", 0)))
                    self.card_error.val_label.config(text=str(stats.get("error", 0)))

                    self.prog_lbl.config(text=f"Selesai! {stats.get('success', 0)} berhasil, {stats.get('skipped', 0)} dilewati, {stats.get('error', 0)} gagal.")
                    self._log_to_console(
                        "INFO",
                        f"=== SELESAI: Berhasil: {stats.get('success', 0)} | Dilewati: {stats.get('skipped', 0)} | Gagal: {stats.get('error', 0)} ===",
                    )
                    messagebox.showinfo(
                        "Proses Selesai",
                        f"Proses pengiriman selesai!\n\n"
                        f"Berhasil: {stats.get('success', 0)}\n"
                        f"Dilewati: {stats.get('skipped', 0)}\n"
                        f"Gagal   : {stats.get('error', 0)}",
                    )
                    try:
                        self._refresh_history_ui()
                    except Exception:
                        pass
        except queue.Empty:
            pass

        try:
            if hasattr(self, "root") and self.root.winfo_exists():
                self._after_id = self.root.after(100, self._process_log_queue)
        except Exception:
            pass

    def _log_to_console(self, status: str, message: str):
        timestamp = time.strftime("[%H:%M:%S] ")
        self.log_console.insert(tk.END, timestamp, "TIMESTAMP")
        tag_name = status if status in ["SUCCESS", "SKIPPED", "ERROR", "INFO"] else "INFO"
        self.log_console.insert(tk.END, f"[{status}] ", tag_name)
        self.log_console.insert(tk.END, f"{message}\n")
        self.log_console.see(tk.END)

    def _clear_log(self):
        self.log_console.delete("1.0", tk.END)


def main():
    root = tk.Tk()
    app = BPSDispatcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
