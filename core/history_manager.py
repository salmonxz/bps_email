"""
History Manager Module for BPS Email Dispatcher.
Provides SQLite-based persistent storage for sent email history, multi-sender filtering,
search, full message retrieval, and Excel reporting.
"""

import os
import sys
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import List, Dict, Any, Optional


def get_default_db_path() -> str:
    """Get absolute path to history.db, compatible with script and PyInstaller executable."""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, "history.db")


class HistoryManager:
    """Manages persistent dispatch history using a local SQLite database."""

    def __init__(self, db_path: Optional[str] = None):
        self.db_path = db_path or get_default_db_path()
        self._init_db()

    @contextmanager
    def _get_connection(self):
        """Create and yield a connection to SQLite database, ensuring proper close."""
        conn = sqlite3.connect(self.db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def _init_db(self):
        """Create table and indexes if they do not exist."""
        # Ensure parent directory exists
        parent_dir = os.path.dirname(os.path.abspath(self.db_path))
        if parent_dir and not os.path.exists(parent_dir):
            os.makedirs(parent_dir, exist_ok=True)

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS email_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    sender_email TEXT NOT NULL,
                    recipient_to TEXT NOT NULL,
                    recipient_cc TEXT DEFAULT '',
                    recipient_bcc TEXT DEFAULT '',
                    provinsi_kode TEXT DEFAULT '',
                    provinsi_nama TEXT DEFAULT '',
                    subject TEXT NOT NULL,
                    body_html TEXT DEFAULT '',
                    body_text TEXT DEFAULT '',
                    attachments TEXT DEFAULT '',
                    status TEXT NOT NULL,
                    error_message TEXT DEFAULT '',
                    is_dry_run INTEGER DEFAULT 0
                );
                """
            )
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_sender ON email_history(sender_email);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_timestamp ON email_history(timestamp);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_status ON email_history(status);")
            conn.commit()

    def add_record(
        self,
        sender_email: str,
        recipient_to: str,
        subject: str,
        status: str,
        recipient_cc: str = "",
        recipient_bcc: str = "",
        provinsi_kode: str = "",
        provinsi_nama: str = "",
        body_html: str = "",
        body_text: str = "",
        attachments: str = "",
        error_message: str = "",
        is_dry_run: bool = False,
        timestamp: Optional[str] = None,
    ) -> int:
        """
        Record an email dispatch attempt (Success or Failure).
        Returns the inserted record ID.
        """
        if not timestamp:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        query = """
        INSERT INTO email_history (
            timestamp, sender_email, recipient_to, recipient_cc, recipient_bcc,
            provinsi_kode, provinsi_nama, subject, body_html, body_text,
            attachments, status, error_message, is_dry_run
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """
        params = (
            timestamp,
            (sender_email or "").strip(),
            (recipient_to or "").strip(),
            (recipient_cc or "").strip(),
            (recipient_bcc or "").strip(),
            str(provinsi_kode or "").strip(),
            str(provinsi_nama or "").strip(),
            str(subject or "").strip(),
            str(body_html or ""),
            str(body_text or ""),
            str(attachments or "").strip(),
            str(status or "SUCCESS").upper(),
            str(error_message or "").strip(),
            1 if is_dry_run else 0,
        )

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            conn.commit()
            return cursor.lastrowid

    def get_distinct_senders(self) -> List[str]:
        """Return a sorted list of unique sender email addresses in history."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT DISTINCT sender_email
                FROM email_history
                WHERE sender_email IS NOT NULL AND TRIM(sender_email) != ''
                ORDER BY sender_email ASC;
                """
            )
            rows = cursor.fetchall()
            return [r["sender_email"] for r in rows if r["sender_email"]]

    def get_history(
        self,
        sender_filter: Optional[str] = None,
        status_filter: Optional[str] = None,
        search_query: Optional[str] = None,
        include_dry_run: bool = True,
        limit: int = 2000,
        offset: int = 0,
    ) -> List[Dict[str, Any]]:
        """
        Retrieve dispatch history rows matching filters, sorted newest to oldest.
        """
        conditions = []
        params: List[Any] = []

        if sender_filter and sender_filter.strip() and sender_filter.strip().lower() not in ["semua", "[semua akun pengirim]"]:
            conditions.append("LOWER(sender_email) = LOWER(?)")
            params.append(sender_filter.strip())

        if status_filter and status_filter.strip() and status_filter.strip().lower() not in ["semua", "[semua status]"]:
            clean_stat = status_filter.strip().upper()
            if "SUCCESS" in clean_stat or "BERHASIL" in clean_stat:
                conditions.append("status = 'SUCCESS'")
            elif "FAILED" in clean_stat or "GAGAL" in clean_stat or "ERROR" in clean_stat:
                conditions.append("status IN ('FAILED', 'ERROR')")
            elif "SKIPPED" in clean_stat or "LEWATI" in clean_stat:
                conditions.append("status = 'SKIPPED'")
            else:
                conditions.append("status = ?")
                params.append(clean_stat)

        if not include_dry_run:
            conditions.append("is_dry_run = 0")

        if search_query and search_query.strip():
            kw = f"%{search_query.strip()}%"
            conditions.append(
                "(provinsi_nama LIKE ? OR provinsi_kode LIKE ? OR recipient_to LIKE ? OR subject LIKE ? OR sender_email LIKE ?)"
            )
            params.extend([kw, kw, kw, kw, kw])

        where_clause = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        query = f"""
            SELECT id, timestamp, sender_email, recipient_to, recipient_cc, recipient_bcc,
                   provinsi_kode, provinsi_nama, subject, attachments, status, error_message, is_dry_run
            FROM email_history
            {where_clause}
            ORDER BY id DESC
            LIMIT ? OFFSET ?;
        """
        params.extend([limit, offset])

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            rows = cursor.fetchall()
            return [dict(r) for r in rows]

    def get_record_by_id(self, record_id: int) -> Optional[Dict[str, Any]]:
        """Retrieve full details of a specific email dispatch record including body content."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM email_history WHERE id = ?;", (record_id,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def clear_history(self, sender_filter: Optional[str] = None) -> int:
        """
        Delete history records.
        If sender_filter is specified and not 'semua', only delete records for that sender.
        Returns the number of deleted records.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if sender_filter and sender_filter.strip() and sender_filter.strip().lower() not in ["semua", "[semua akun pengirim]"]:
                cursor.execute(
                    "DELETE FROM email_history WHERE LOWER(sender_email) = LOWER(?);",
                    (sender_filter.strip(),),
                )
            else:
                cursor.execute("DELETE FROM email_history;")
            deleted_count = cursor.rowcount
            conn.commit()
            return deleted_count

    def export_to_excel(
        self,
        output_path: str,
        sender_filter: Optional[str] = None,
        status_filter: Optional[str] = None,
        search_query: Optional[str] = None,
    ) -> int:
        """
        Export matching history records to an Excel (.xlsx) file with professional formatting.
        Returns the count of exported records.
        """
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        records = self.get_history(
            sender_filter=sender_filter,
            status_filter=status_filter,
            search_query=search_query,
            include_dry_run=True,
            limit=50000,
        )

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Riwayat Pengiriman"

        # Headers
        headers = [
            "No",
            "Waktu Kirim",
            "Akun Pengirim",
            "Kode Prov",
            "Nama Provinsi",
            "Penerima (To)",
            "Tembusan (Cc)",
            "Tembusan Rahasia (Bcc)",
            "Subjek Surat",
            "Daftar Lampiran",
            "Status",
            "Tipe",
            "Pesan Error / Keterangan",
        ]
        ws.append(headers)

        # Style header row
        header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Append records
        success_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        failed_fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")

        for idx, rec in enumerate(records, start=1):
            status_str = "BERHASIL" if rec["status"] == "SUCCESS" else "GAGAL"
            tipe_str = "Simulasi" if rec.get("is_dry_run") else "Pengiriman Riil"

            row_data = [
                idx,
                rec.get("timestamp", ""),
                rec.get("sender_email", ""),
                rec.get("provinsi_kode", ""),
                rec.get("provinsi_nama", ""),
                rec.get("recipient_to", ""),
                rec.get("recipient_cc", ""),
                rec.get("recipient_bcc", ""),
                rec.get("subject", ""),
                rec.get("attachments", ""),
                status_str,
                tipe_str,
                rec.get("error_message", ""),
            ]
            ws.append(row_data)

            # Apply cell styles
            row_idx = idx + 1
            is_success = rec["status"] == "SUCCESS"
            curr_fill = success_fill if is_success else failed_fill

            for col_num in range(1, len(headers) + 1):
                c = ws.cell(row=row_idx, column=col_num)
                c.border = thin_border
                c.font = Font(name="Calibri", size=10)
                if col_num in [1, 2, 4, 11, 12]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

                if col_num == 11:
                    c.fill = curr_fill
                    c.font = Font(
                        name="Calibri",
                        size=10,
                        bold=True,
                        color="2E7D32" if is_success else "C62828",
                    )

        # Set column widths
        col_widths = {
            1: 6,   # No
            2: 20,  # Waktu
            3: 25,  # Pengirim
            4: 12,  # Kode
            5: 22,  # Nama Prov
            6: 28,  # To
            7: 28,  # Cc
            8: 24,  # Bcc
            9: 35,  # Subjek
            10: 25, # Lampiran
            11: 14, # Status
            12: 15, # Tipe
            13: 35, # Keterangan
        }
        for col_idx, width in col_widths.items():
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        wb.save(output_path)
        return len(records)
