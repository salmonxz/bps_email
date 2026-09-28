"""
Unit tests for HistoryManager in core/history_manager.py.
"""

import os
import tempfile
import unittest
import openpyxl
from core.history_manager import HistoryManager


class TestHistoryManager(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_history.db")
        self.hm = HistoryManager(self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_database_initialization(self):
        self.assertTrue(os.path.exists(self.db_path))
        senders = self.hm.get_distinct_senders()
        self.assertEqual(senders, [])

    def test_add_and_retrieve_record(self):
        rec_id = self.hm.add_record(
            sender_email="sender_a@bps.go.id",
            recipient_to="kepala.aceh@bps.go.id",
            recipient_cc="ketua.aceh@bps.go.id",
            recipient_bcc="staf.aceh@bps.go.id",
            provinsi_kode="11",
            provinsi_nama="Aceh",
            subject="Surat Tugas IMK - Aceh",
            body_html="<p>Halo Kepala BPS Aceh</p>",
            body_text="Halo Kepala BPS Aceh",
            attachments="tabel_11.xlsx, pedoman.pdf",
            status="SUCCESS",
        )
        self.assertGreater(rec_id, 0)

        record = self.hm.get_record_by_id(rec_id)
        self.assertIsNotNone(record)
        self.assertEqual(record["sender_email"], "sender_a@bps.go.id")
        self.assertEqual(record["recipient_to"], "kepala.aceh@bps.go.id")
        self.assertEqual(record["provinsi_kode"], "11")
        self.assertEqual(record["provinsi_nama"], "Aceh")
        self.assertEqual(record["subject"], "Surat Tugas IMK - Aceh")
        self.assertEqual(record["status"], "SUCCESS")
        self.assertEqual(record["body_html"], "<p>Halo Kepala BPS Aceh</p>")

    def test_multi_sender_filtering(self):
        # Add records for Sender A
        self.hm.add_record(
            sender_email="admin_a@bps.go.id",
            recipient_to="aceh@bps.go.id",
            subject="Subj A1",
            status="SUCCESS",
            provinsi_nama="Aceh",
        )
        self.hm.add_record(
            sender_email="admin_a@bps.go.id",
            recipient_to="sumut@bps.go.id",
            subject="Subj A2",
            status="SUCCESS",
            provinsi_nama="Sumatera Utara",
        )

        # Add records for Sender B
        self.hm.add_record(
            sender_email="admin_b@bps.go.id",
            recipient_to="bali@bps.go.id",
            subject="Subj B1",
            status="FAILED",
            error_message="Connection timed out",
            provinsi_nama="Bali",
        )

        # Distinct senders
        senders = self.hm.get_distinct_senders()
        self.assertEqual(senders, ["admin_a@bps.go.id", "admin_b@bps.go.id"])

        # Filter by Sender A
        hist_a = self.hm.get_history(sender_filter="admin_a@bps.go.id")
        self.assertEqual(len(hist_a), 2)
        for h in hist_a:
            self.assertEqual(h["sender_email"], "admin_a@bps.go.id")

        # Filter by Sender B
        hist_b = self.hm.get_history(sender_filter="admin_b@bps.go.id")
        self.assertEqual(len(hist_b), 1)
        self.assertEqual(hist_b[0]["sender_email"], "admin_b@bps.go.id")
        self.assertEqual(hist_b[0]["status"], "FAILED")

        # All senders
        all_hist = self.hm.get_history(sender_filter="[Semua Akun Pengirim]")
        self.assertEqual(len(all_hist), 3)

    def test_status_filter_and_search(self):
        self.hm.add_record(
            sender_email="sender@bps.go.id",
            recipient_to="aceh@bps.go.id",
            subject="Pemberitahuan Rekonsiliasi",
            status="SUCCESS",
            provinsi_nama="Aceh",
        )
        self.hm.add_record(
            sender_email="sender@bps.go.id",
            recipient_to="papua@bps.go.id",
            subject="Pemberitahuan Rekonsiliasi",
            status="FAILED",
            error_message="SMTP 550 Mailbox not found",
            provinsi_nama="Papua",
        )

        # Status filter SUCCESS
        successes = self.hm.get_history(status_filter="Berhasil")
        self.assertEqual(len(successes), 1)
        self.assertEqual(successes[0]["provinsi_nama"], "Aceh")

        # Status filter FAILED
        failures = self.hm.get_history(status_filter="Gagal")
        self.assertEqual(len(failures), 1)
        self.assertEqual(failures[0]["provinsi_nama"], "Papua")

        # Search query by provinsi
        papua_res = self.hm.get_history(search_query="Papua")
        self.assertEqual(len(papua_res), 1)
        self.assertEqual(papua_res[0]["provinsi_nama"], "Papua")

        # Search query by subject keyword
        rekon_res = self.hm.get_history(search_query="Rekonsiliasi")
        self.assertEqual(len(rekon_res), 2)

    def test_export_to_excel(self):
        self.hm.add_record(
            sender_email="export_test@bps.go.id",
            recipient_to="test1@bps.go.id",
            subject="Export Subj 1",
            status="SUCCESS",
            provinsi_kode="31",
            provinsi_nama="DKI Jakarta",
        )
        self.hm.add_record(
            sender_email="export_test@bps.go.id",
            recipient_to="test2@bps.go.id",
            subject="Export Subj 2",
            status="FAILED",
            error_message="Auth error",
            provinsi_kode="32",
            provinsi_nama="Jawa Barat",
        )

        excel_out = os.path.join(self.temp_dir.name, "export_history.xlsx")
        count = self.hm.export_to_excel(excel_out, sender_filter="export_test@bps.go.id")
        self.assertEqual(count, 2)
        self.assertTrue(os.path.exists(excel_out))

        # Check excel content
        wb = openpyxl.load_workbook(excel_out)
        ws = wb.active
        self.assertEqual(ws.title, "Riwayat Pengiriman")
        self.assertEqual(ws.cell(row=1, column=1).value, "No")
        self.assertEqual(ws.cell(row=2, column=5).value, "Jawa Barat")
        self.assertEqual(ws.cell(row=3, column=5).value, "DKI Jakarta")

    def test_clear_history(self):
        self.hm.add_record(sender_email="a@bps.go.id", recipient_to="1@b.com", subject="S1", status="SUCCESS")
        self.hm.add_record(sender_email="b@bps.go.id", recipient_to="2@b.com", subject="S2", status="SUCCESS")

        # Delete only sender a
        deleted = self.hm.clear_history(sender_filter="a@bps.go.id")
        self.assertEqual(deleted, 1)
        remaining = self.hm.get_history()
        self.assertEqual(len(remaining), 1)
        self.assertEqual(remaining[0]["sender_email"], "b@bps.go.id")

        # Delete all
        deleted_all = self.hm.clear_history()
        self.assertEqual(deleted_all, 1)
        self.assertEqual(len(self.hm.get_history()), 0)


if __name__ == "__main__":
    unittest.main()
