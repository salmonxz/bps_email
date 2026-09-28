"""
Unit tests for BPS Provincial Email Dispatcher core components.
"""

import os
import unittest
from core.excel_parser import BPSExcelParser, format_kode_provinsi, is_valid_email
from core.attachment_engine import AttachmentEngine
from core.mailer import (
    replace_template,
    build_email_message,
    EmailDispatcher,
    extract_envelope_recipients,
    check_smtp_sender_compatibility,
)


class TestBPSDispatcher(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        cls.sample_excel = os.path.join(cls.base_dir, "sample_data", "sample_kontak_bps.xlsx")
        cls.sample_att_dir = os.path.join(cls.base_dir, "sample_data", "attachments")

    def test_kode_formatting(self):
        self.assertEqual(format_kode_provinsi(11), "11")
        self.assertEqual(format_kode_provinsi("11.0"), "11")
        self.assertEqual(format_kode_provinsi(1), "01")
        self.assertEqual(format_kode_provinsi(" 31 "), "31")
        self.assertIsNone(format_kode_provinsi(""))
        self.assertIsNone(format_kode_provinsi("Total"))

    def test_email_validation(self):
        self.assertTrue(is_valid_email("kepala@bps.go.id"))
        self.assertTrue(is_valid_email("user.test+extra@domain.com"))
        self.assertFalse(is_valid_email("not_an_email"))
        self.assertFalse(is_valid_email(""))

    def test_excel_parser_multilevel(self):
        self.assertTrue(os.path.exists(self.sample_excel), f"Excel not found: {self.sample_excel}")
        parser = BPSExcelParser(self.sample_excel)
        records = parser.load()

        self.assertGreaterEqual(len(records), 8)

        # Check Aceh record
        aceh = next((r for r in records if r["kode"] == "11"), None)
        self.assertIsNotNone(aceh)
        self.assertEqual(aceh["provinsi"], "Aceh")
        self.assertEqual(aceh["nama_kepala"], "Dr. Ahmad, M.Si")
        self.assertEqual(aceh["email_kepala"], "ahmad.aceh@bps.go.id")
        self.assertEqual(aceh["nama_pjk"], "Budi Santoso, S.Si")
        self.assertEqual(aceh["email_pjk"], "budi.aceh@bps.go.id")
        self.assertTrue(aceh["is_valid"])

        # Check Sumbar record (No Cc)
        sumbar = next((r for r in records if r["kode"] == "13"), None)
        self.assertIsNotNone(sumbar)
        self.assertEqual(sumbar["email_pjk"], "")
        self.assertTrue(sumbar["is_valid"])

        # Check Maluku record (Invalid email)
        maluku = next((r for r in records if r["kode"] == "81"), None)
        self.assertIsNotNone(maluku)
        self.assertFalse(maluku["is_valid"])
        self.assertTrue(any("tidak valid" in err for err in maluku["errors"]))

    def test_attachment_engine(self):
        engine = AttachmentEngine(self.sample_att_dir)
        total_files, matched_count = engine.scan()

        self.assertGreater(total_files, 0)
        self.assertEqual(len(engine.national_files), 2)  # 00_Naskah_Nasional, 00_Pedoman

        # Check Aceh (11): 2 national + 2 provincial = 4 files
        aceh_files = engine.get_attachments_for_province("11")
        self.assertEqual(len(aceh_files), 4)
        base_names = [os.path.basename(f) for f in aceh_files]
        self.assertIn("00_Naskah_Nasional_2026.pdf", base_names)
        self.assertIn("00_Pedoman_Teknis_Pelaporan.pdf", base_names)
        self.assertIn("11_Naskah_Statistik_Aceh.pdf", base_names)
        self.assertIn("11_Lampiran_Tabel_Aceh.pdf", base_names)

        # Check Sumut (12): 2 national + 1 provincial = 3 files
        sumut_files = engine.get_attachments_for_province("12")
        self.assertEqual(len(sumut_files), 3)

        # Check Jatim (35): 2 national + 0 provincial = 2 files
        jatim_files = engine.get_attachments_for_province("35")
        self.assertEqual(len(jatim_files), 2)

    def test_template_replacement(self):
        template_subj = "Distribusi Naskah Statistik - BPS Provinsi {provinsi} ({kode})"
        template_body = "Yth. Bapak/Ibu {nama_kepala},\nCc: {nama_pjk}\nBerikut terlampir dokumen provinsi {provinsi}."

        data = {
            "kode": "31",
            "provinsi": "DKI Jakarta",
            "nama_kepala": "Dr. Nur Indah",
            "nama_pjk": "Farhan Maulana",
        }

        subj = replace_template(template_subj, data)
        body = replace_template(template_body, data)

        self.assertEqual(subj, "Distribusi Naskah Statistik - BPS Provinsi DKI Jakarta (31)")
        self.assertIn("Yth. Bapak/Ibu Dr. Nur Indah,", body)
        self.assertIn("Cc: Farhan Maulana", body)
        self.assertIn("provinsi DKI Jakarta", body)

    def test_dry_run_simulation(self):
        parser = BPSExcelParser(self.sample_excel)
        records = parser.load()

        engine = AttachmentEngine(self.sample_att_dir)
        engine.scan()

        dispatcher = EmailDispatcher("smtp.test.com", 587, "test@bps.go.id", "secret")

        logs = []
        progress_calls = []

        def on_status(status, msg, item):
            logs.append((status, msg))

        def on_progress(cur, tot):
            progress_calls.append((cur, tot))

        stats = dispatcher.dispatch_batch(
            items=records,
            subject_template="Uji Coba - {provinsi}",
            body_template="Halo {nama_kepala}",
            attachment_engine=engine,
            status_callback=on_status,
            progress_callback=on_progress,
            dry_run=True,
        )

        # Maluku (invalid email) should be skipped, others succeed
        self.assertGreaterEqual(stats["success"], 7)
        self.assertGreaterEqual(stats["skipped"], 1)
        self.assertEqual(stats["error"], 0)
        self.assertEqual(len(progress_calls), len(records))

    def test_dispatcher_separate_username_and_from_email(self):
        # Relay Brevo test case
        login_user = "b7cfd1001@smtp-brevo.com"
        official_from = "ryehanalfiansyah24@gmail.com"

        dispatcher = EmailDispatcher(
            host="smtp-relay.brevo.com",
            port=587,
            username=login_user,
            password="secret_relay_key",
            sender_email=official_from,
        )

        self.assertEqual(dispatcher.username, login_user)
        self.assertEqual(dispatcher.sender_email, official_from)

        # Verify built message sets msg['From'] to official_from
        msg = build_email_message(
            sender=dispatcher.sender_email,
            to_email="tujuan@bps.go.id",
            cc_email="tembusan@bps.go.id",
            subject="Test Subject",
            body="Test Body",
            attachment_paths=[],
        )
        self.assertEqual(msg["From"], official_from)
        self.assertEqual(msg["To"], "tujuan@bps.go.id")
        self.assertEqual(msg["Cc"], "tembusan@bps.go.id")

    def test_config_json_structure(self):
        import json
        config_path = os.path.join(self.base_dir, "config.json")
        self.assertTrue(os.path.exists(config_path))
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.assertIn("smtp_host", data)
        self.assertIn("smtp_port", data)
        self.assertIn("smtp_username", data)
        self.assertIn("sender_email", data)

    def test_official_bps_excel_and_multi_cc(self):
        official_name = "Daftar E-mail Kepala Provinsi dan PJK Statistik Industri Provinsi.xlsx"
        official_excel = os.path.join(self.base_dir, official_name)
        self.assertTrue(os.path.exists(official_excel), f"File not found: {official_excel}")
        parser = BPSExcelParser(official_excel)
        records = parser.load()

        self.assertEqual(len(records), 38)
        labels = [c["label"] for c in parser.email_columns]
        self.assertTrue(any("Kepala Prov" in l and "[TO]" in l for l in labels))
        self.assertTrue(any("Ketua Tim" in l and "[CC]" in l for l in labels))
        self.assertTrue(any("Koordinator" in l and "[CC]" in l for l in labels))
        self.assertTrue(any("Staf Seksi" in l and "[BCC]" in l for l in labels))

        # Check Aceh record (11)
        aceh = next((r for r in records if r["kode"] == "11"), None)
        self.assertIsNotNone(aceh)
        self.assertEqual(aceh["provinsi"], "Aceh")
        self.assertTrue("@" in aceh["email_kepala"])
        self.assertTrue(len(aceh["email_tim_prod"]) > 0)
        self.assertTrue(len(aceh["email_koor"]) > 0)
        self.assertTrue(len(aceh["email_staf"]) > 0)

        # Cc & Bcc check: Ketua Tim and Koordinator in Cc, Staf in Bcc
        self.assertIn(aceh["email_tim_prod"], aceh["cc_email"])
        self.assertIn(aceh["email_koor"], aceh["cc_email"])
        self.assertEqual(aceh["email_staf"], aceh["bcc_email"])

        # Check Jambi (15) with 2 Ccs
        jambi = next((r for r in records if r["kode"] == "15"), None)
        self.assertIsNotNone(jambi)
        self.assertEqual(jambi["provinsi"], "Jambi")
        self.assertEqual(jambi["email_kepala"], "aidil@bps.go.id")
        self.assertEqual(jambi["email_tim_prod"], "riani@bps.go.id")
        self.assertEqual(jambi["email_koor"], "oemar.wibisono@bps.go.id")
        self.assertIn("riani@bps.go.id", jambi["cc_email"])
        self.assertIn("oemar.wibisono@bps.go.id", jambi["cc_email"])

        # Check Sumbar (13) with no Cc
        sumbar = next((r for r in records if r["kode"] == "13"), None)
        self.assertIsNotNone(sumbar)
        self.assertEqual(sumbar["cc_email"], "")

    def test_multi_format_attachment_engine(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            f1 = os.path.join(tmpdir, "Tabel Profil IMK 2025 - Provinsi 16.xlsx")
            f2 = os.path.join(tmpdir, "11_Laporan Aceh.docx")
            f3 = os.path.join(tmpdir, "Pedoman Teknis Nasional 2025.pdf")
            f4 = os.path.join(tmpdir, "Arsip Data - 31.zip")
            f5 = os.path.join(tmpdir, "00_Tabel_Nasional.xlsx")
            for f in [f1, f2, f3, f4, f5]:
                with open(f, "w") as fp:
                    fp.write("dummy")

            engine = AttachmentEngine(folder_path=tmpdir)
            total_files, matched = engine.scan()

            self.assertEqual(total_files, 5)
            self.assertEqual(matched, 5)
            # Both 'Pedoman Teknis Nasional 2025.pdf' and '00_Tabel_Nasional.xlsx' must be national
            self.assertEqual(len(engine.national_files), 2)
            nat_names = [os.path.basename(f) for f in engine.national_files]
            self.assertIn("00_Tabel_Nasional.xlsx", nat_names)
            self.assertIn("Pedoman Teknis Nasional 2025.pdf", nat_names)

            # Sumsel (16) should match 2 national + suffix 'Tabel Profil IMK 2025 - Provinsi 16.xlsx' = 3 files
            sumsel_files = engine.get_attachments_for_province("16")
            self.assertEqual(len(sumsel_files), 3)
            sumsel_names = [os.path.basename(f) for f in sumsel_files]
            self.assertIn("00_Tabel_Nasional.xlsx", sumsel_names)
            self.assertIn("Pedoman Teknis Nasional 2025.pdf", sumsel_names)
            self.assertIn("Tabel Profil IMK 2025 - Provinsi 16.xlsx", sumsel_names)

            # Aceh (11) should match 2 national + prefix '11_Laporan Aceh.docx' = 3 files
            aceh_files = engine.get_attachments_for_province("11")
            self.assertEqual(len(aceh_files), 3)
            aceh_names = [os.path.basename(f) for f in aceh_files]
            self.assertIn("00_Tabel_Nasional.xlsx", aceh_names)
            self.assertIn("11_Laporan Aceh.docx", aceh_names)

    def test_attachment_folder_level_detection(self):
        import tempfile
        import shutil

        with tempfile.TemporaryDirectory() as tmpdir:
            # Create subfolders named with province codes
            folder_11 = os.path.join(tmpdir, "11_Aceh")
            folder_12 = os.path.join(tmpdir, "12")
            folder_00 = os.path.join(tmpdir, "00_Nasional")
            folder_31 = os.path.join(tmpdir, "31 - DKI Jakarta")
            os.makedirs(folder_11)
            os.makedirs(folder_12)
            os.makedirs(folder_00)
            os.makedirs(folder_31)

            # Files inside subfolders DO NOT have province codes in their filename
            with open(os.path.join(folder_11, "dokumen_naskah.pdf"), "w") as f:
                f.write("test aceh")
            with open(os.path.join(folder_12, "tabel_data.xlsx"), "w") as f:
                f.write("test sumut")
            with open(os.path.join(folder_00, "pedoman_umum.docx"), "w") as f:
                f.write("test nasional")
            with open(os.path.join(folder_31, "arsip_lengkap.zip"), "w") as f:
                f.write("test dki")

            # 1. Test scanning parent directory
            engine = AttachmentEngine(folder_path=tmpdir)
            total, matched = engine.scan()
            self.assertEqual(total, 4)
            self.assertEqual(matched, 4)

            self.assertEqual(len(engine.national_files), 1)
            self.assertIn("11", engine.provincial_files)
            self.assertIn("12", engine.provincial_files)
            self.assertIn("31", engine.provincial_files)

            # 2. Test scanning root folder itself that starts with code (e.g. 11_Aceh)
            engine_single = AttachmentEngine(folder_path=folder_11)
            t_s, m_s = engine_single.scan()
            self.assertEqual(t_s, 1)
            self.assertEqual(m_s, 1)
            self.assertIn("11", engine_single.provincial_files)

    def test_38_provinces_official_excel(self):
        official_name = "Daftar E-mail Kepala Provinsi dan PJK Statistik Industri Provinsi.xlsx"
        official_file = os.path.join(self.base_dir, official_name)
        if os.path.exists(official_file):
            parser = BPSExcelParser(official_file)
            records = parser.load()
            self.assertEqual(len(records), 38)
            codes = [r["kode"] for r in records]
            self.assertIn("11", codes)  # Aceh
            self.assertIn("12", codes)  # Sumut
            self.assertIn("31", codes)  # DKI
            self.assertIn("97", codes)  # Papua Pegunungan

            # Verify that Aceh has To, Cc, Bcc allocated properly
            aceh = next((r for r in records if r["kode"] == "11"), None)
            self.assertIsNotNone(aceh)
            self.assertEqual(aceh["to_email"], "mhmmdandri280@gmail.com")
            self.assertIn("windapratiwi4704@gmail.com", aceh["cc_email"])
            self.assertIn("tyodisu51@gmail.com", aceh["cc_email"])
            self.assertEqual(aceh["bcc_email"], "ryehanalfiansyah24@gmail.com")
            self.assertTrue(aceh["is_valid"])

    def test_dynamic_tag_placeholders(self):
        data = {
            "kode": "32",
            "provinsi": "Jawa Barat",
            "nama_kepala": "Drs. Agus Sudarto, M.Si",
            "nama_ketua": "Siti Rahmawati, S.Si",
        }

        # Test double curly braces {{tag}}
        subj_curly = replace_template("Laporan {{provinsi}} ({{kode}})", data)
        body_curly = replace_template("Yth. {{nama_kepala}}, Cc: {{nama_ketua}}", data)
        self.assertEqual(subj_curly, "Laporan Jawa Barat (32)")
        self.assertEqual(body_curly, "Yth. Drs. Agus Sudarto, M.Si, Cc: Siti Rahmawati, S.Si")

        # Test double angle brackets <<tag>> with spaces / synonyms
        subj_angle = replace_template("Laporan <<nama provinsi>> (<<kode provinsi>>)", data)
        body_angle = replace_template("Kepada ketua tim produksi <<nama provinsi>>, Yth. <<nama kepala>>, Cc: <<nama ketua tim>>", data)
        self.assertEqual(subj_angle, "Laporan Jawa Barat (32)")
        self.assertIn("Kepada ketua tim produksi Jawa Barat", body_angle)
        self.assertIn("Yth. Drs. Agus Sudarto, M.Si", body_angle)
        self.assertIn("Cc: Siti Rahmawati, S.Si", body_angle)

        # Test single curly braces {tag} backward compatibility
        subj_single = replace_template("Laporan {provinsi} ({kode})", data)
        self.assertEqual(subj_single, "Laporan Jawa Barat (32)")

    def test_envelope_recipients_and_no_bcc_header(self):
        to_email = "kepala@bps.go.id"
        cc_email = "ketua@bps.go.id, koor@bps.go.id"
        bcc_email = "staf@bps.go.id"

        envelope = extract_envelope_recipients(to_email, cc_email, bcc_email)
        self.assertEqual(len(envelope), 4)
        self.assertIn("kepala@bps.go.id", envelope)
        self.assertIn("ketua@bps.go.id", envelope)
        self.assertIn("koor@bps.go.id", envelope)
        self.assertIn("staf@bps.go.id", envelope)

        msg = build_email_message(
            sender="admin@bps.go.id",
            to_email=to_email,
            cc_email=cc_email,
            subject="Test Subject",
            body="Test Body",
            attachment_paths=[],
        )

        # Ensure To and Cc headers are present
        self.assertEqual(msg["To"], to_email)
        self.assertEqual(msg["Cc"], cc_email)
        # Ensure Bcc header is NOT added to the message
        self.assertIsNone(msg["Bcc"])

    def test_fallback_priority_and_allocation(self):
        import pandas as pd
        import tempfile

        # Create a mock Excel file with 4 email columns to test all fallback combinations
        data = [
            [None, None, None, None, None, None, None, None, None, None, None, None],
            [None, None, None, None, "Daftar Kontak", None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None, None, None, None, None, None],
            [None, "No", "Kode", "Provinsi", "Kepala Prov", None, "Ketua Tim Statistik Produksi", None, "Koordinator Statistik Industri", None, "Staf Seksi Statistik Industri", None],
            [None, None, None, None, "Nama", "E-mail", "Nama", "Email", "Nama", "E-mail", "Nama", "E-mail"],
            [None, "-1", "-2", "-3", "-4", "-5", "-6", "-7", "-8", "-9", "-10", "-11"],
            # Row 1: All present -> To: Kepala, Cc: Ketua + Koor, Bcc: Staf
            [None, 1, 11, "Prov A", "Kep A", "kepala@bps.go.id", "Ket A", "ketua@bps.go.id", "Koor A", "koor@bps.go.id", "Staf A", "staf@bps.go.id"],
            # Row 2: Kepala empty -> Fallback To: Ketua, Cc: Koor, Bcc: Staf
            [None, 2, 12, "Prov B", "Kep B", None, "Ket B", "ketua@bps.go.id", "Koor B", "koor@bps.go.id", "Staf B", "staf@bps.go.id"],
            # Row 3: Kepala & Ketua empty -> Fallback To: Koor, Cc: empty, Bcc: Staf
            [None, 3, 13, "Prov C", "Kep C", None, "Ket C", None, "Koor C", "koor@bps.go.id", "Staf C", "staf@bps.go.id"],
            # Row 4: Kepala, Ketua, Koor empty -> Fallback To: Staf, Cc: empty, Bcc: empty
            [None, 4, 14, "Prov D", "Kep D", None, "Ket D", None, "Koor D", None, "Staf D", "staf@bps.go.id"],
            # Row 5: All empty -> invalid, error
            [None, 5, 15, "Prov E", "Kep E", None, "Ket E", None, "Koor E", None, "Staf E", None],
        ]

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            df = pd.DataFrame(data)
            df.to_excel(tmp_path, header=False, index=False)

            parser = BPSExcelParser(tmp_path)
            records = parser.load()

            self.assertEqual(len(records), 5)

            # Row 1: All present
            r1 = records[0]
            self.assertTrue(r1["is_valid"])
            self.assertEqual(r1["to_email"], "kepala@bps.go.id")
            self.assertIn("ketua@bps.go.id", r1["cc_email"])
            self.assertIn("koor@bps.go.id", r1["cc_email"])
            self.assertEqual(r1["bcc_email"], "staf@bps.go.id")

            # Row 2: Kepala empty -> To is Ketua
            r2 = records[1]
            self.assertTrue(r2["is_valid"])
            self.assertEqual(r2["to_email"], "ketua@bps.go.id")
            self.assertIn("koor@bps.go.id", r2["cc_email"])
            self.assertNotIn("ketua@bps.go.id", r2["cc_email"])
            self.assertEqual(r2["bcc_email"], "staf@bps.go.id")

            # Row 3: Kepala & Ketua empty -> To is Koor
            r3 = records[2]
            self.assertTrue(r3["is_valid"])
            self.assertEqual(r3["to_email"], "koor@bps.go.id")
            self.assertEqual(r3["cc_email"], "")
            self.assertEqual(r3["bcc_email"], "staf@bps.go.id")

            # Row 4: Only Staf present -> To is Staf
            r4 = records[3]
            self.assertTrue(r4["is_valid"])
            self.assertEqual(r4["to_email"], "staf@bps.go.id")
            self.assertEqual(r4["cc_email"], "")
            self.assertEqual(r4["bcc_email"], "")

            # Row 5: All empty -> is_valid False
            r5 = records[4]
            self.assertFalse(r5["is_valid"])
            self.assertEqual(r5["to_email"], "")
            self.assertTrue(any("kosong" in e.lower() for e in r5["errors"]))

        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


    def test_dynamic_tags_detection_official_excel(self):
        official_path = os.path.join(self.base_dir, "Daftar E-mail Kepala Provinsi dan PJK Statistik Industri Provinsi.xlsx")
        if not os.path.exists(official_path):
            self.skipTest("Official Excel not found")

        parser = BPSExcelParser(official_path)
        records = parser.load()

        # Dynamic tags detected
        expected_tags = ["kode", "provinsi", "nama_kepala", "nama_ketua", "nama_koordinator", "nama_staf", "no"]
        for t in expected_tags:
            self.assertIn(t, parser.dynamic_tags)

        # Ensure no email columns are in dynamic_tags
        for t in parser.dynamic_tags:
            self.assertNotIn("email", t.lower())
            self.assertNotIn("e-mail", t.lower())

        # Check Aceh record has all tags filled
        aceh = next(r for r in records if r["kode"] == "11")
        self.assertEqual(aceh["no"], "1")
        self.assertEqual(aceh["kode"], "11")
        self.assertEqual(aceh["provinsi"], "Aceh")
        self.assertEqual(aceh["nama_kepala"], "andre")
        self.assertEqual(aceh["nama_ketua"], "Hendra Gunawan")
        self.assertEqual(aceh["nama_koordinator"], "Melly")
        self.assertEqual(aceh["nama_staf"], "Rodiah")

        # Check Bengkulu record (empty kepala)
        bengkulu = next(r for r in records if r["kode"] == "17")
        self.assertEqual(bengkulu["nama_kepala"], "")
        self.assertEqual(bengkulu["nama_ketua"], "Fitri Aryati")
        self.assertEqual(bengkulu["nama_koordinator"], "Boby Fernando")
        self.assertEqual(bengkulu["nama_staf"], "Amalela Neti")

    def test_dynamic_template_replacement_case_insensitive_and_multisyntax(self):
        data = {
            "no": "1",
            "kode": "11",
            "provinsi": "Aceh",
            "nama_kepala": "andre",
            "nama_ketua": "Hendra Gunawan",
            "nama_koordinator": "Melly",
            "nama_staf": "Rodiah",
        }

        # Multi-syntax: {{...}}, <<...>>, {...}
        # Case variations: UPPERCASE, lowercase, TitleCase
        # Space variations: <<nama provinsi>>, {{ nama_kepala }}
        template_subj = "Surat <<NAMA PROVINSI>> (<<kode>>) - No: {{no}}"
        template_body = (
            "Kepada Yth. <<Nama_Kepala>>,\n"
            "Ditempat: {{PROVINSI}}\n"
            "Ketua Tim: <<nama_ketua>>\n"
            "Koordinator: {{ nama_koordinator }}\n"
            "Staf: <<nama_staf>>\n"
            "Kode: {kode}"
        )

        subj = replace_template(template_subj, data)
        body = replace_template(template_body, data)

        self.assertEqual(subj, "Surat Aceh (11) - No: 1")
        self.assertIn("Kepada Yth. andre,", body)
        self.assertIn("Ditempat: Aceh", body)
        self.assertIn("Ketua Tim: Hendra Gunawan", body)
        self.assertIn("Koordinator: Melly", body)
        self.assertIn("Staf: Rodiah", body)
        self.assertIn("Kode: 11", body)

    def test_dynamic_template_nan_replacement(self):
        # Empty and NaN values must be replaced by "" without producing 'nan'
        data = {
            "no": None,
            "kode": "17",
            "provinsi": "Bengkulu",
            "nama_kepala": float("nan"),
            "nama_ketua": "Fitri Aryati",
            "nama_koordinator": "",
            "nama_staf": None,
        }

        tmpl = "No: {{no}} | Kepala: {{nama_kepala}} | Koor: <<nama_koordinator>> | Staf: {nama_staf} | Ketua: {{nama_ketua}}"
        res = replace_template(tmpl, data)

        self.assertNotIn("nan", res.lower())
        self.assertNotIn("none", res.lower())
        self.assertEqual(res, "No:  | Kepala:  | Koor:  | Staf:  | Ketua: Fitri Aryati")

    def test_explicit_header_role_detection(self):
        from core.excel_parser import detect_explicit_email_role
        self.assertEqual(detect_explicit_email_role("Alamat Email [TO]"), "to")
        self.assertEqual(detect_explicit_email_role("To: Email Kepala"), "to")
        self.assertEqual(detect_explicit_email_role("Ketua Tim - Email (Cc)"), "cc")
        self.assertEqual(detect_explicit_email_role("CC: Koordinator"), "cc")
        self.assertEqual(detect_explicit_email_role("Email Staf [Bcc]"), "bcc")
        self.assertEqual(detect_explicit_email_role("BCC: Tim Arsip"), "bcc")
        self.assertIsNone(detect_explicit_email_role("Alamat Email Pegawai"))

    def test_html_email_mime_structure(self):
        from core.mailer import is_html_content
        self.assertTrue(is_html_content("<div>Halo BPS</div>"))
        self.assertTrue(is_html_content("<p>Surat Resmi<br>Tanggal 2026</p>"))
        self.assertFalse(is_html_content("Surat biasa tanpa tag HTML"))

        html_body = "<h2>Pemberitahuan</h2><p>Yth. <b>Kepala BPS</b></p><br><a href='https://bps.go.id'>Situs</a>"
        msg = build_email_message(
            sender="admin@bps.go.id",
            to_email="kepala@bps.go.id",
            cc_email="ketua@bps.go.id",
            subject="Uji HTML Email",
            body=html_body,
            attachment_paths=[],
        )
        self.assertTrue(msg.is_multipart())
        payloads = msg.get_payload()
        # Should contain multipart/alternative with plain text and html
        content_types = [p.get_content_type() for p in payloads]
        if any("multipart/alternative" in ct for ct in content_types):
            alt_part = next(p for p in payloads if "multipart/alternative" in p.get_content_type())
            alt_cts = [p.get_content_type() for p in alt_part.get_payload()]
            self.assertIn("text/plain", alt_cts)
            self.assertIn("text/html", alt_cts)
        else:
            self.assertIn("text/plain", content_types)
            self.assertIn("text/html", content_types)


    def test_safe_get_functionality(self):
        import pandas as pd
        from core.excel_parser import safe_get

        s = pd.Series(["11", "Aceh", float("nan"), "test@bps.go.id"])
        # In-bounds valid values
        self.assertEqual(safe_get(s, 0), "11")
        self.assertEqual(safe_get(s, 1), "Aceh")
        self.assertEqual(safe_get(s, 3), "test@bps.go.id")

        # NaN and None handling
        self.assertEqual(safe_get(s, 2), "")
        self.assertEqual(safe_get(s, None), "")

        # Out of bounds and negative indices - must never raise IndexError
        self.assertEqual(safe_get(s, 4), "")
        self.assertEqual(safe_get(s, 10), "")
        self.assertEqual(safe_get(s, 99), "")
        self.assertEqual(safe_get(s, -1), "")

        # List / tuple handling
        lst = ["31", "DKI Jakarta"]
        self.assertEqual(safe_get(lst, 0), "31")
        self.assertEqual(safe_get(lst, 5), "")

    def test_deleted_columns_no_index_error_and_dynamic_hierarchy(self):
        """
        Test that when 'Kepala Prov' is deleted (reducing table from 12 to 10 columns):
        1. No 'index out of bounds' error occurs.
        2. Ketua Tim (1st leftmost) automatically becomes 'To'.
        3. Koordinator (middle) automatically becomes 'Cc'.
        4. Staf (last rightmost) automatically becomes 'Bcc'.
        """
        import pandas as pd
        import tempfile

        # 10 columns total (No, Kode, Provinsi, Ketua Tim Nama, Ketua Tim Email, Koor Nama, Koor Email, Staf Nama, Staf Email)
        data = [
            [None] * 10,
            [None, None, None, "Daftar Kontak Tanpa Kepala Prov", None, None, None, None, None, None],
            [None] * 10,
            # Row 4 (index 3): Header Level 1
            [None, "No", "Kode", "Provinsi", "Ketua Tim Statistik", None, "Koordinator Industri", None, "Staf Seksi", None],
            # Row 5 (index 4): Header Level 2
            [None, None, None, None, "Nama", "E-mail", "Nama", "E-mail", "Nama", "E-mail"],
            # Row 6 (index 5): Numbering row
            [None, "-1", "-2", "-3", "-4", "-5", "-6", "-7", "-8", "-9"],
            # Row 7 (index 6): Data row 1 - All emails present
            [None, 1, 11, "Aceh", "Budi Ketua", "ketua.aceh@bps.go.id", "Rian Koor", "koor.aceh@bps.go.id", "Siti Staf", "staf.aceh@bps.go.id"],
            # Row 8 (index 7): Data row 2 - Ketua empty -> Koordinator shifts to To, Staf is Bcc
            [None, 2, 12, "Sumatera Utara", "Andi Ketua", None, "Dedi Koor", "koor.sumut@bps.go.id", "Lina Staf", "staf.sumut@bps.go.id"],
            # Row 9 (index 8): Data row 3 - Both Ketua & Koordinator empty -> Staf shifts to To
            [None, 3, 13, "Sumatera Barat", "Hendra Ketua", None, "Maya Koor", None, "Rini Staf", "staf.sumbar@bps.go.id"],
            # Row 10 (index 9): Data row 4 - All email columns empty -> invalid row
            [None, 4, 14, "Riau", "Doni Ketua", None, "Rudi Koor", None, "Eka Staf", None],
        ]

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_file = tmp.name

        try:
            df = pd.DataFrame(data)
            df.to_excel(tmp_file, header=False, index=False)

            parser = BPSExcelParser(tmp_file)
            records = parser.load()

            self.assertEqual(len(records), 4)

            # Check email column role detection
            # 3 email columns detected: Ketua Tim, Koordinator, Staf
            self.assertEqual(len(parser.email_columns), 3)
            roles = [c["role_type"] for c in parser.email_columns]
            self.assertEqual(roles[0], "to")   # Leftmost = To
            self.assertEqual(roles[1], "cc")   # Middle = Cc
            self.assertEqual(roles[2], "bcc")  # Rightmost = Bcc

            # Check labels have proper role tags
            self.assertTrue(any("Ketua Tim" in c["label"] and "[TO]" in c["label"] for c in parser.email_columns))
            self.assertTrue(any("Koordinator" in c["label"] and "[CC]" in c["label"] for c in parser.email_columns))
            self.assertTrue(any("Staf Seksi" in c["label"] and "[BCC]" in c["label"] for c in parser.email_columns))

            # Row 1 (Aceh): All present
            r1 = records[0]
            self.assertTrue(r1["is_valid"])
            self.assertEqual(r1["to_email"], "ketua.aceh@bps.go.id")
            self.assertEqual(r1["cc_email"], "koor.aceh@bps.go.id")
            self.assertEqual(r1["bcc_email"], "staf.aceh@bps.go.id")

            # Row 2 (Sumut): Ketua empty -> Koordinator shifts to To, Staf is Bcc
            r2 = records[1]
            self.assertTrue(r2["is_valid"])
            self.assertEqual(r2["to_email"], "koor.sumut@bps.go.id")
            self.assertEqual(r2["cc_email"], "")
            self.assertEqual(r2["bcc_email"], "staf.sumut@bps.go.id")

            # Row 3 (Sumbar): Ketua & Koor empty -> Staf shifts to To
            r3 = records[2]
            self.assertTrue(r3["is_valid"])
            self.assertEqual(r3["to_email"], "staf.sumbar@bps.go.id")
            self.assertEqual(r3["cc_email"], "")
            self.assertEqual(r3["bcc_email"], "")

            # Row 4 (Riau): All empty -> invalid row
            r4 = records[3]
            self.assertFalse(r4["is_valid"])
            self.assertEqual(r4["to_email"], "")
            self.assertTrue(any("kosong" in e.lower() for e in r4["errors"]))

        finally:
            if os.path.exists(tmp_file):
                os.remove(tmp_file)

    def test_explicit_tags_override_default_hierarchy(self):
        """Test that explicit [TO], [CC], [BCC] tags in headers override default left-to-right hierarchy."""
        import pandas as pd
        import tempfile

        data = [
            [None, "No", "Kode", "Provinsi", "Tim Teknis [BCC]", "Koordinator [TO]", "Staf Arsip [CC]"],
            [None, 1, 31, "DKI Jakarta", "teknis@bps.go.id", "koor@bps.go.id", "arsip@bps.go.id"],
        ]

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_file = tmp.name

        try:
            df = pd.DataFrame(data)
            df.to_excel(tmp_file, header=False, index=False)

            parser = BPSExcelParser(tmp_file)
            records = parser.load()

            self.assertEqual(len(records), 1)
            r = records[0]
            self.assertTrue(r["is_valid"])
            # Despite Koordinator being in the middle, explicit [TO] makes it To
            self.assertEqual(r["to_email"], "koor@bps.go.id")
            # Despite Tim Teknis being leftmost, explicit [BCC] makes it Bcc
            self.assertEqual(r["bcc_email"], "teknis@bps.go.id")
            # Staf Arsip has explicit [CC]
            self.assertEqual(r["cc_email"], "arsip@bps.go.id")

        finally:
            if os.path.exists(tmp_file):
                os.remove(tmp_file)

    def test_comprehensive_dynamic_tag_aliases(self):
        """
        Verify that all user-requested tag variants:
        - nama ketua / ketua
        - koordinator
        - staff / staf / nama staff
        - kode provisi / kode provinsi / kode
        - provisi / provinsi
        - kepala / nama kepala
        - no / nomor
        across various syntaxes ({{...}}, <<...>>, &lt;&lt;...&gt;&gt;, {...}, [[...]])
        are all automatically connected and replaced with data from the Excel record.
        """
        data = {
            "no": "1",
            "kode": "11",
            "provinsi": "Aceh",
            "nama_ketua": "Budi Santoso",
            "nama_koordinator": "Ir. Melly",
            "nama_staf": "Rodiah Rahma",
            "nama_kepala": "Dr. Ahmad",
        }

        # 1. Ketua variants
        self.assertEqual(replace_template("Halo {{nama_ketua}}", data), "Halo Budi Santoso")
        self.assertEqual(replace_template("Halo {{nama ketua}}", data), "Halo Budi Santoso")
        self.assertEqual(replace_template("Halo {{ketua}}", data), "Halo Budi Santoso")
        self.assertEqual(replace_template("Halo <<ketua>>", data), "Halo Budi Santoso")
        self.assertEqual(replace_template("Halo {ketua}", data), "Halo Budi Santoso")

        # 2. Koordinator variants
        self.assertEqual(replace_template("Yth. {{koordinator}}", data), "Yth. Ir. Melly")
        self.assertEqual(replace_template("Yth. <<koordinator>>", data), "Yth. Ir. Melly")
        self.assertEqual(replace_template("Yth. &lt;&lt;koordinator&gt;&gt;", data), "Yth. Ir. Melly")
        self.assertEqual(replace_template("Yth. {{nama_koordinator}}", data), "Yth. Ir. Melly")
        self.assertEqual(replace_template("Yth. {{nama koordinator}}", data), "Yth. Ir. Melly")

        # 3. Staff / Staf variants
        self.assertEqual(replace_template("Kepada {{staff}}", data), "Kepada Rodiah Rahma")
        self.assertEqual(replace_template("Kepada <<staff>>", data), "Kepada Rodiah Rahma")
        self.assertEqual(replace_template("Kepada [[staff]]", data), "Kepada Rodiah Rahma")
        self.assertEqual(replace_template("Kepada {{staf}}", data), "Kepada Rodiah Rahma")
        self.assertEqual(replace_template("Kepada {{nama_staff}}", data), "Kepada Rodiah Rahma")
        self.assertEqual(replace_template("Kepada {{nama staff}}", data), "Kepada Rodiah Rahma")
        self.assertEqual(replace_template("Kepada {{nama_staf}}", data), "Kepada Rodiah Rahma")

        # 4. Kode Provisi / Provinsi variants
        self.assertEqual(replace_template("Kode: {{kode provisi}}", data), "Kode: 11")
        self.assertEqual(replace_template("Kode: {{kode_provisi}}", data), "Kode: 11")
        self.assertEqual(replace_template("Kode: <<kode provisi>>", data), "Kode: 11")
        self.assertEqual(replace_template("Kode: {{kode_provinsi}}", data), "Kode: 11")
        self.assertEqual(replace_template("Kode: {{kode provinsi}}", data), "Kode: 11")
        self.assertEqual(replace_template("Kode: {{kode}}", data), "Kode: 11")

        # 5. Provisi / Provinsi name variants
        self.assertEqual(replace_template("Prov: {{provisi}}", data), "Prov: Aceh")
        self.assertEqual(replace_template("Prov: {{nama_provisi}}", data), "Prov: Aceh")
        self.assertEqual(replace_template("Prov: {{provinsi}}", data), "Prov: Aceh")
        self.assertEqual(replace_template("Prov: {{nama_provinsi}}", data), "Prov: Aceh")

        # 6. Combined realistic email body test
        tmpl = (
            "Yth. Bapak/Ibu Kepala BPS {{provinsi}} ({{kepala}}),\n"
            "Cc: Ketua Tim {{nama ketua}} dan Koordinator {{koordinator}},\n"
            "Staf Pengelola: {{staff}}.\n"
            "Dokumen untuk Provinsi {{provisi}} (Kode: {{kode provisi}}) No: {{nomor}}."
        )
        rendered = replace_template(tmpl, data)
        self.assertIn("Kepala BPS Aceh (Dr. Ahmad)", rendered)
        self.assertIn("Ketua Tim Budi Santoso dan Koordinator Ir. Melly", rendered)
        self.assertIn("Staf Pengelola: Rodiah Rahma", rendered)
        self.assertIn("Dokumen untuk Provinsi Aceh (Kode: 11) No: 1", rendered)

    def test_dispatcher_records_to_history_manager(self):
        import tempfile
        from core.history_manager import HistoryManager

        with tempfile.TemporaryDirectory() as tmp_dir:
            db_path = os.path.join(tmp_dir, "hist_test.db")
            hm = HistoryManager(db_path)

            dispatcher = EmailDispatcher(
                host="smtp.fake.com",
                port=587,
                username="pengirim@bps.go.id",
                password="pwd",
                sender_email="pengirim@bps.go.id",
                history_manager=hm,
            )

            dummy_items = [
                {
                    "kode": "11",
                    "provinsi": "Aceh",
                    "to_email": "aceh@bps.go.id",
                    "cc_email": "cc.aceh@bps.go.id",
                    "bcc_email": "bcc.aceh@bps.go.id",
                    "nama_kepala": "Dr. Ahmad",
                    "is_valid": True,
                }
            ]

            class DummyAttEngine:
                def get_attachments_for_province(self, code):
                    return []

            # Run dry-run dispatch
            stats = dispatcher.dispatch_batch(
                items=dummy_items,
                subject_template="Subjek IMK {{provinsi}}",
                body_template="Halo {{nama_kepala}}",
                attachment_engine=DummyAttEngine(),
                status_callback=lambda *args: None,
                progress_callback=lambda *args: None,
                dry_run=True,
            )

            self.assertEqual(stats["success"], 1)
            hist = hm.get_history(sender_filter="pengirim@bps.go.id")
            self.assertEqual(len(hist), 1)
            self.assertEqual(hist[0]["provinsi_nama"], "Aceh")
            self.assertEqual(hist[0]["recipient_to"], "aceh@bps.go.id")
            self.assertEqual(hist[0]["subject"], "Subjek IMK Aceh")

    def test_create_flexible_ssl_context(self):
        import ssl
        from core.mailer import create_flexible_ssl_context
        ctx = create_flexible_ssl_context()
        self.assertIsInstance(ctx, ssl.SSLContext)
        self.assertFalse(ctx.check_hostname)
        self.assertEqual(ctx.verify_mode, ssl.CERT_NONE)

    def test_test_smtp_connection_validation(self):
        from core.mailer import test_smtp_connection
        # Missing host
        ok, msg = test_smtp_connection("", 587)
        self.assertFalse(ok)
        self.assertIn("Host SMTP tidak boleh kosong", msg)

    def test_dispatcher_handles_refused_recipients(self):
        import tempfile
        from unittest.mock import MagicMock
        from core.history_manager import HistoryManager

        with tempfile.TemporaryDirectory() as tmp_dir:
            db_path = os.path.join(tmp_dir, "hist_refused.db")
            hm = HistoryManager(db_path)

            dispatcher = EmailDispatcher(
                host="mail.bps.go.id",
                port=587,
                username="dinas@bps.go.id",
                password="secret_password",
                sender_email="dinas@bps.go.id",
                history_manager=hm,
            )

            mock_server = MagicMock()
            # Simulate server refusing recipient
            mock_server.send_message.return_value = {
                "invalid@bps.go.id": (550, b"User mailbox does not exist")
            }
            dispatcher._connect = MagicMock(return_value=mock_server)

            dummy_items = [
                {
                    "kode": "11",
                    "provinsi": "Aceh",
                    "to_email": "invalid@bps.go.id",
                    "is_valid": True,
                }
            ]

            class DummyAttEngine:
                def get_attachments_for_province(self, code):
                    return []

            errors_logged = []
            def status_cb(status, msg, record):
                if status == "ERROR":
                    errors_logged.append(msg)

            stats = dispatcher.dispatch_batch(
                items=dummy_items,
                subject_template="Subjek IMK {{provinsi}}",
                body_template="Halo",
                attachment_engine=DummyAttEngine(),
                status_callback=status_cb,
                progress_callback=lambda *args: None,
                dry_run=False,
            )

            self.assertEqual(stats["error"], 1)
            self.assertEqual(stats["success"], 0)
            self.assertTrue(any("Server SMTP menolak penerima" in e for e in errors_logged))
            hist = hm.get_history(sender_filter="dinas@bps.go.id")
            self.assertEqual(len(hist), 1)
            self.assertEqual(hist[0]["status"], "FAILED")
            self.assertIn("550 User mailbox does not exist", hist[0]["error_message"])


class TestSMTPSenderCompatibility(unittest.TestCase):
    """Test suite for validating SMTP host and sender address compatibility."""

    def test_empty_sender(self):
        ok, warnings, recom = check_smtp_sender_compatibility("mail.bps.go.id", "user@bps.go.id", "")
        self.assertFalse(ok)
        self.assertTrue(any("kosong" in w for w in warnings))
        self.assertEqual(recom, "user@bps.go.id")

    def test_bps_host_with_gmail_sender_mismatch(self):
        ok, warnings, recom = check_smtp_sender_compatibility(
            host="mail.bps.go.id",
            username="arya.pratama@bps.go.id",
            sender_email="ryehanalfiansyah24@gmail.com",
        )
        self.assertFalse(ok)
        self.assertTrue(any("Anti-Spoofing" in w or "bukan @bps.go.id" in w for w in warnings))
        self.assertEqual(recom, "arya.pratama@bps.go.id")

    def test_bps_host_with_matching_bps_sender(self):
        ok, warnings, recom = check_smtp_sender_compatibility(
            host="mail.bps.go.id",
            username="arya.pratama@bps.go.id",
            sender_email="arya.pratama@bps.go.id",
        )
        self.assertTrue(ok)
        self.assertEqual(len(warnings), 0)

    def test_bps_host_with_different_bps_username(self):
        ok, warnings, recom = check_smtp_sender_compatibility(
            host="mail.bps.go.id",
            username="admin@bps.go.id",
            sender_email="staf@bps.go.id",
        )
        self.assertFalse(ok)
        self.assertTrue(any("berbeda" in w for w in warnings))

    def test_gmail_host_with_bps_sender_mismatch(self):
        ok, warnings, recom = check_smtp_sender_compatibility(
            host="smtp.gmail.com",
            username="myuser@gmail.com",
            sender_email="admin@bps.go.id",
        )
        self.assertFalse(ok)
        self.assertTrue(any("bukan @gmail.com" in w for w in warnings))
        self.assertEqual(recom, "myuser@gmail.com")

    def test_gmail_host_with_matching_sender(self):
        ok, warnings, recom = check_smtp_sender_compatibility(
            host="smtp.gmail.com",
            username="myuser@gmail.com",
            sender_email="myuser@gmail.com",
        )
        self.assertTrue(ok)
        self.assertEqual(len(warnings), 0)

    def test_generic_different_domains(self):
        ok, warnings, recom = check_smtp_sender_compatibility(
            host="smtp.office365.com",
            username="user@company-a.com",
            sender_email="user@company-b.com",
        )
        self.assertFalse(ok)
        self.assertTrue(any("Domain akun login" in w for w in warnings))


if __name__ == "__main__":
    unittest.main()


