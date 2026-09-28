"""
Automated GUI flow test verifying app instantiation, data loading, treeview population,
checkbox selection, and dry-run simulation execution.
"""

import os
import sys
import time
import unittest
import tkinter as tk

# Ensure current directory is in path
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, base_dir)

from unittest.mock import patch
from app import BPSDispatcherApp


class TestGUIIntegration(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide root during test
        self.patcher_info = patch("tkinter.messagebox.showinfo")
        self.patcher_warn = patch("tkinter.messagebox.showwarning")
        self.patcher_err = patch("tkinter.messagebox.showerror")
        self.mock_info = self.patcher_info.start()
        self.mock_warn = self.patcher_warn.start()
        self.mock_err = self.patcher_err.start()
        self.app = BPSDispatcherApp(self.root)

    def tearDown(self):
        self.patcher_info.stop()
        self.patcher_warn.stop()
        self.patcher_err.stop()
        try:
            self.app.cleanup()
            self.root.destroy()
        except Exception:
            pass

    def test_full_gui_load_and_dry_run(self):
        excel_file = os.path.join(base_dir, "sample_data", "sample_kontak_bps.xlsx")
        att_dir = os.path.join(base_dir, "sample_data", "attachments")

        self.assertTrue(os.path.exists(excel_file))
        self.assertTrue(os.path.isdir(att_dir))

        # 1. Set paths
        self.app.excel_path.set(excel_file)
        self.app._load_excel_data()

        self.app.folder_path.set(att_dir)
        self.app._scan_attachment_folder()

        self.assertGreaterEqual(len(self.app.records), 8)
        self.assertIsNotNone(self.app.attachment_engine)

        # 2. Check Treeview items
        items = self.app.tree.get_children()
        self.assertEqual(len(items), len(self.app.records))

        # First item (Aceh - 11) has valid email and attachments -> Siap Kirim -> should be auto-checked [✓]
        first_vals = self.app.tree.item(items[0], "values")
        self.assertEqual(first_vals[0], "[✓]")
        self.assertEqual(first_vals[2], "11")

        # 3. Simulate row detail view
        self.app.tree.selection_set(items[0])
        self.app._on_tree_select(None)
        detail_content = self.app.detail_text.get("1.0", tk.END)
        self.assertIn("DETAIL TARGET", detail_content)
        self.assertIn("11", detail_content)
        self.assertIn("Aceh", detail_content)

        # 4. Trigger dry-run simulation
        self.app._start_dry_run()

        # Pump events until finished (max 5 seconds)
        for _ in range(50):
            self.root.update()
            if not self.app.is_running:
                break
            time.sleep(0.1)

        self.assertFalse(self.app.is_running)
        success_val = int(self.app.card_success.val_label.cget("text"))
        skipped_val = int(self.app.card_skipped.val_label.cget("text"))

        # Maluku has invalid email so it is NOT auto-checked [✓]. Only checked valid rows run.
        self.assertGreaterEqual(success_val, 7)
        self.assertEqual(skipped_val, 0)

    def test_gui_config_save_and_load(self):
        # Test modifying config and saving/loading
        self.app.smtp_host.set("smtp-relay.brevo.com")
        self.app.smtp_port.set("587")
        self.app.smtp_username.set("b7cfd1001@smtp-brevo.com")
        self.app.sender_email.set("ryehanalfiansyah24@gmail.com")
        self.app.email_delay.set(2.5)

        self.app._save_config(show_popup=False)

        # Reset values
        self.app.smtp_username.set("")
        self.app.sender_email.set("")

        # Reload
        self.app._load_config()

        self.assertEqual(self.app.smtp_host.get(), "smtp-relay.brevo.com")
        self.assertEqual(self.app.smtp_username.get(), "b7cfd1001@smtp-brevo.com")
        self.assertEqual(self.app.sender_email.get(), "ryehanalfiansyah24@gmail.com")
        self.assertEqual(self.app.email_delay.get(), 2.5)

    def test_checkbox_selection_and_selective_dispatch(self):
        excel_file = os.path.join(base_dir, "sample_data", "sample_kontak_bps.xlsx")
        att_dir = os.path.join(base_dir, "sample_data", "attachments")

        self.app.excel_path.set(excel_file)
        self.app._load_excel_data()
        self.app.folder_path.set(att_dir)
        self.app._scan_attachment_folder()

        items = self.app.tree.get_children()
        self.assertGreaterEqual(len(items), 2)

        # Deselect all rows
        self.app._deselect_all_rows()
        self.assertEqual(len(self.app._get_checked_records()), 0)
        for iid in items:
            self.assertEqual(self.app.tree.item(iid, "values")[0], "[ ]")

        # Trying to proceed with 0 checked rows should show warning
        self.app._proceed_to_execution()
        self.mock_warn.assert_called()
        self.assertIn("Pilihan Kosong", self.mock_warn.call_args[0][0])

        # Check only first 2 rows (Aceh & Sumut)
        self.app._toggle_checkbox(items[0])
        self.app._toggle_checkbox(items[1])
        self.assertEqual(len(self.app._get_checked_records()), 2)
        self.assertEqual(self.app.tree.item(items[0], "values")[0], "[✓]")
        self.assertEqual(self.app.tree.item(items[1], "values")[0], "[✓]")

        # Proceed to execution
        self.app._proceed_to_execution()
        self.assertEqual(len(self.app.target_records), 2)
        self.assertIn("2 Provinsi Terpilih [✓]", self.app.target_mode_desc)
        self.assertIn("Menyiapkan pengiriman untuk 2 provinsi terpilih", self.app.target_banner_lbl.cget("text"))

        # Run dry-run simulation for just the 2 selected rows
        self.app._start_dry_run()

        for _ in range(50):
            self.root.update()
            if not self.app.is_running:
                break
            time.sleep(0.1)

        self.assertFalse(self.app.is_running)
        success_val = int(self.app.card_success.val_label.cget("text"))
        self.assertEqual(success_val, 2)

        # Select all rows
        self.app._select_all_rows()
        self.assertEqual(len(self.app._get_checked_records()), len(items))
        for iid in items:
            self.assertEqual(self.app.tree.item(iid, "values")[0], "[✓]")

    def test_flexible_attachment_input(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            f1 = os.path.join(tmpdir, "Tabel Profil IMK 2025 - Provinsi 16.xlsx")
            f2 = os.path.join(tmpdir, "11_Ringkasan_Aceh.docx")
            f3 = os.path.join(tmpdir, "00_Pedoman_Nasional.pdf")
            for f in [f1, f2, f3]:
                with open(f, "w") as fp:
                    fp.write("dummy")

            # 1. Test folder mode
            self.app._scan_attachment_folder(actual_path=tmpdir)
            self.assertIsNotNone(self.app.attachment_engine)
            self.assertEqual(len(self.app.attachment_engine.all_files), 3)

            # 2. Test custom multi-file mode
            self.app.custom_attachment_files = [f1, f2, f3]
            self.app._scan_custom_files()
            self.assertIsNotNone(self.app.attachment_engine)

            # Check matching for 16 (Sumsel) includes both 00 file and suffix 16 file
            sumsel_atts = self.app.attachment_engine.get_attachments_for_province("16")
            sumsel_names = [os.path.basename(f) for f in sumsel_atts]
            self.assertTrue(any("16" in name for name in sumsel_names))
            self.assertTrue(any("00" in name for name in sumsel_names))

    def test_dynamic_email_columns_in_treeview(self):
        official_excel = os.path.join(base_dir, "Daftar E-mail Kepala Provinsi dan PJK Statistik Industri Provinsi.xlsx")
        self.assertTrue(os.path.exists(official_excel))

        # 1. Load official BPS Excel file
        self.app.excel_path.set(official_excel)
        self.app._load_excel_data()

        # Check that 4 email columns were detected
        self.assertEqual(len(self.app.email_columns), 4)
        col_keys = [c["key"] for c in self.app.email_columns]
        self.assertEqual(col_keys, ["email_kepala", "email_tim_prod", "email_koor", "email_staf"])

        # Check that Treeview columns match
        tree_cols = list(self.app.tree["columns"])
        self.assertIn("email_kepala", tree_cols)
        self.assertIn("email_tim_prod", tree_cols)
        self.assertIn("email_koor", tree_cols)
        self.assertIn("email_staf", tree_cols)

        # Check Aceh row in treeview
        items = self.app.tree.get_children()
        self.assertGreaterEqual(len(items), 38)
        aceh_vals = self.app.tree.item(items[0], "values")
        # Structure follows Excel: check, no, kode, prov, nama_kep, email_kep, nama_ket, email_ket, nama_koor, email_koor, nama_staf, email_staf, nat, prov, tot, status
        self.assertEqual(aceh_vals[2], "11")
        self.assertEqual(aceh_vals[3], "Aceh")
        self.assertTrue(len(aceh_vals[4]) > 0)  # nama_kepala (e.g. 'andre')
        self.assertTrue("@" in aceh_vals[5])    # email_kepala
        self.assertTrue(len(aceh_vals[6]) > 0)  # nama_ketua (e.g. 'Hendra Gunawan')
        self.assertTrue("@" in aceh_vals[7])    # email_tim_prod
        self.assertTrue(len(aceh_vals[8]) > 0)  # nama_koor (e.g. 'Melly')
        self.assertTrue("@" in aceh_vals[9])    # email_koor
        self.assertTrue(len(aceh_vals[10]) > 0) # nama_staf (e.g. 'Rodiah')
        self.assertTrue("@" in aceh_vals[11])   # email_staf

        # 2. Check Detail Inspection view
        self.app.tree.selection_set(items[0])
        self.app._on_tree_select(None)
        detail_content = self.app.detail_text.get("1.0", tk.END)
        self.assertTrue(any("Kepala Prov" in line and "[TO]" in line for line in detail_content.splitlines()))
        self.assertTrue(any("Ketua Tim" in line and "[CC]" in line for line in detail_content.splitlines()))
        self.assertTrue(any("Koordinator" in line and "[CC]" in line for line in detail_content.splitlines()))
        self.assertTrue(any("Staf Seksi" in line and "[BCC]" in line for line in detail_content.splitlines()))
        self.assertIn("ALOKASI PENGIRIMAN (TO / CC / BCC)", detail_content)

        # Check template UI properties (height >= 13 and tag insertion)
        self.assertGreaterEqual(int(self.app.body_text.cget("height")), 13)
        self.app._insert_tag("{{provinsi}}")
        self.assertIn("{{provinsi}}", self.app.body_text.get("1.0", tk.END))

        # Check tag guidance label & buttons in tag_box
        guide_str = self.app.tag_guide_lbl.cget("text")
        self.assertIn("{{kode}}", guide_str)
        self.assertIn("{{nama_kepala}}", guide_str)
        self.assertIn("{{nama_staf}}", guide_str)
        self.assertIn("{{no}}", guide_str)
        self.assertGreaterEqual(len(self.app.tag_buttons_frame.winfo_children()), 7)

        # Test inserting into subject entry when focused
        self.app.last_focused_text_widget = self.app.subject_entry
        self.app._insert_tag("{{kode}}")
        self.assertIn("{{kode}}", self.app.subject_entry.get())

        # 3. Test dynamic transition when switching to a 2-column Excel file
        sample_excel = os.path.join(base_dir, "sample_data", "sample_kontak_bps.xlsx")
        self.app.excel_path.set(sample_excel)
        self.app._load_excel_data()

        # Should adapt to 2 email columns
        self.assertEqual(len(self.app.email_columns), 2)
        new_tree_cols = list(self.app.tree["columns"])
        self.assertTrue(any("email" in c for c in new_tree_cols))
        self.assertNotIn("email_tim_prod", new_tree_cols)

    def test_national_attachment_auto_checks_all_provinces(self):
        import tempfile
        official_excel = os.path.join(base_dir, "Daftar E-mail Kepala Provinsi dan PJK Statistik Industri Provinsi.xlsx")
        self.app.excel_path.set(official_excel)
        self.app._load_excel_data()

        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a single national file with prefix 00_
            nat_file = os.path.join(tmpdir, "00_Tabel_IMK_Nasional.xlsx")
            with open(nat_file, "w") as fp:
                fp.write("dummy")

            # Scan the custom national file
            self.app.custom_attachment_files = [nat_file]
            self.app._scan_custom_files()

            # Verify that all 38 provinces are now valid and auto-checked [✓],
            # because Bengkulu (17) which has an empty Kepala Prov email now falls back to Ketua Tim!
            checked_records = self.app._get_checked_records()
            self.assertEqual(len(checked_records), 38)

            # Check that Bengkulu (row with empty Kepala email) successfully falls back to Ketua Tim and IS checked
            bengkulu_item = next(iid for iid in self.app.tree.get_children() if self.app.tree.item(iid, "values")[2] == "17")
            self.assertEqual(self.app.tree.item(bengkulu_item, "values")[0], "[✓]")
            bengkulu_rec = next(r for r in checked_records if r["kode"] == "17")
            self.assertEqual(bengkulu_rec["to_email"], "fitriar@bps.go.id")
            self.assertIn("Ketua Tim Statistik Produksi", bengkulu_rec["to_role"])
            self.assertEqual(bengkulu_rec["bcc_email"], "amalela@bps.go.id")

            # Check that Aceh (11) and Jawa Timur (35) ARE checked
            aceh_item = next(iid for iid in self.app.tree.get_children() if self.app.tree.item(iid, "values")[2] == "11")
            jatim_item = next(iid for iid in self.app.tree.get_children() if self.app.tree.item(iid, "values")[2] == "35")
            self.assertEqual(self.app.tree.item(aceh_item, "values")[0], "[✓]")
            self.assertEqual(self.app.tree.item(jatim_item, "values")[0], "[✓]")

    def test_five_tab_navigation_and_history(self):
        # 1. Verify Notebook has 5 tabs
        tabs = self.app.notebook.tabs()
        self.assertEqual(len(tabs), 5)
        tab_texts = [self.app.notebook.tab(t, "text").strip() for t in tabs]
        self.assertTrue(any("1. Konfigurasi Sistem" in t for t in tab_texts))
        self.assertTrue(any("2. Template Pesan & Editor HTML" in t for t in tab_texts))
        self.assertTrue(any("3. Preview & Verifikasi" in t for t in tab_texts))
        self.assertTrue(any("4. Eksekusi & Log Pengiriman" in t for t in tab_texts))
        self.assertTrue(any("5. 📜 Riwayat Pengiriman" in t for t in tab_texts))

        # 2. Test delay presets
        self.app._set_delay_preset(0.0)
        self.assertEqual(self.app.email_delay.get(), 0.0)
        self.app._set_delay_preset(1.5)
        self.assertEqual(self.app.email_delay.get(), 1.5)

        # 3. Test WYSIWYG helpers
        self.app.body_text.delete("1.0", tk.END)
        self.app.body_text.insert(tk.END, "Halo Dunia")
        # Select "Halo"
        self.app.body_text.tag_add(tk.SEL, "1.0", "1.4")
        self.app._format_bold()
        content = self.app.body_text.get("1.0", tk.END)
        self.assertIn("<b>Halo</b>", content)

        # Test insert table
        self.app._insert_table_dialog()
        tbl_content = self.app.body_text.get("1.0", tk.END)
        self.assertIn("<table", tbl_content)

        # Test clean html
        with patch("tkinter.messagebox.askyesno", return_value=True):
            self.app._clean_html_tags()
        cleaned_content = self.app.body_text.get("1.0", tk.END)
        self.assertNotIn("<b>", cleaned_content)
        self.assertNotIn("<table", cleaned_content)

    def test_html_file_loader(self):
        import tempfile
        sample_html = "<html><head><title>Surat Resmi BPS 2026</title></head><body><h1>Pengumuman</h1><p>Isi naskah</p></body></html>"
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8") as tf:
            tf.write(sample_html)
            tmp_path = tf.name

        try:
            self.app.html_file_path.set(tmp_path)
            self.app._load_html_file_content(tmp_path)

            loaded_body = self.app.body_text.get("1.0", tk.END)
            self.assertIn("<h1>Pengumuman</h1>", loaded_body)
            self.assertEqual(self.app.email_subject.get(), "Surat Resmi BPS 2026")
            self.assertIn("berhasil dimuat", self.app.html_status_lbl.cget("text"))

            # Test reset
            self.app._reset_html_file()
            self.assertEqual(self.app.html_file_path.get(), "")
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_tab5_history_gui_workflow(self):
        # 1. Clear any pre-existing records and add records to history
        self.app.history_manager.clear_history()
        self.app.history_manager.add_record(
            sender_email="akun_a@bps.go.id",
            recipient_to="aceh@bps.go.id",
            subject="Surat Tugas Aceh",
            body_html="<p>Isi Surat Aceh</p>",
            status="SUCCESS",
            provinsi_kode="11",
            provinsi_nama="Aceh",
        )
        self.app.history_manager.add_record(
            sender_email="akun_b@bps.go.id",
            recipient_to="bali@bps.go.id",
            subject="Surat Tugas Bali",
            body_html="<p>Isi Surat Bali</p>",
            status="FAILED",
            error_message="SMTP connection dropped",
            provinsi_kode="51",
            provinsi_nama="Bali",
        )

        # 2. Refresh UI
        self.app._refresh_history_ui()
        items = self.app.history_tree.get_children()
        self.assertEqual(len(items), 2)

        # 3. Filter by akun_a
        self.app.hist_sender_filter.set("akun_a@bps.go.id")
        self.app._on_history_filter_change()
        items_a = self.app.history_tree.get_children()
        self.assertEqual(len(items_a), 1)
        row_vals = self.app.history_tree.item(items_a[0], "values")
        self.assertEqual(row_vals[2], "akun_a@bps.go.id")
        self.assertEqual(row_vals[4], "aceh@bps.go.id")

        # 4. Detail selection
        self.app.history_tree.selection_set(items_a[0])
        self.app._on_history_row_selected()
        self.assertEqual(self.app.hist_det_sender.cget("text"), "akun_a@bps.go.id")
        self.assertEqual(self.app.hist_det_to.cget("text"), "aceh@bps.go.id")
        body_shown = self.app.hist_body_text.get("1.0", tk.END)
        self.assertIn("Isi Surat Aceh", body_shown)

        # 5. Reset filter
        self.app._reset_history_filter()
        self.assertEqual(len(self.app.history_tree.get_children()), 2)


if __name__ == "__main__":
    unittest.main()
