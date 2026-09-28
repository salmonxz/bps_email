"""
Script untuk menghasilkan data dummy uji coba sistem Email Blast BPS:
1. File Excel 'Daftar_Email_Dummy_10.xlsx' dengan Multi-Index Header Resmi BPS (12 Kolom, Header Baris 4 & 5, Data Mulai Baris 7)
2. Folder 'dummy_lampiran' dengan file simulasi berbagai format (PDF, Excel, Word, ZIP)
"""

import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def create_dummy_pdf(filepath: str, title: str):
    """Buat file dummy PDF yang valid."""
    content = (
        b"%PDF-1.4\n"
        b"1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n"
        b"2 0 obj<</Type/Pages/Count 1/Kids[3 0 R]>>endobj\n"
        b"3 0 obj<</Type/Page/MediaBox[0 0 595 842]/Parent 2 0 R/Resources<<>>>>endobj\n"
        b"xref\n0 4\n"
        b"0000000000 65535 f \n"
        b"0000000009 00000 n \n"
        b"0000000052 00000 n \n"
        b"0000000102 00000 n \n"
        b"trailer<</Size 4/Root 1 0 R>>\n"
        b"startxref\n178\n%%EOF\n"
    )
    with open(filepath, "wb") as f:
        f.write(content)


def create_dummy_text_file(filepath: str, text: str):
    """Buat file dummy teks/format lain."""
    with open(filepath, "wb") as f:
        f.write(text.encode("utf-8"))


def generate_dummy_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # 3 Email aktif sesuai instruksi (tanpa spasi)
    active_emails = [
        "iguardrehan@gmail.com",
        "ryehanalfiansyah24@gmail.com",
        "windapratiwi4704@gmail.com",
    ]

    # Data 10 provinsi
    provinces_data = [
        ("11", "Aceh", "Dr. H. Ahmad Fauzi, M.Si", "Rahmat Hidayat, S.Stat", "Dewi Sartika, S.Si", "Fikri Maulana, A.Md"),
        ("12", "Sumut", "Dr. Ir. Nurul Huda, M.Si", "Rian Pratama, S.Si", "Siti Rahma, S.E", ""),
        ("13", "Sumatera Barat", "Ir. Herianto, M.E", "", "", ""),  # Tanpa Cc
        ("14", "Riau", "Drs. Asep Supriatna, M.Si", "Dedi Kurniawan, S.Stat", "", "Budi Santoso, S.Kom"),
        ("15", "Jambi", "Hj. Sunarti, S.E., M.Si", "Farhan Maulana, S.Tr.Stat", "Andi Wijaya, S.Si", "Nurlina, S.Stat"),
        ("16", "Sumsel", "Dr. Mohamad Ihsan, M.Si", "Eko Prasetyo, S.Kom", "", ""),
        ("17", "Bengkulu", "Ir. Winarno, M.M", "", "Yusuf Habibie, S.T", ""),
        ("18", "Lampung", "Dr. Zulkipli, M.Stat", "Agus Setiawan, S.Si", "Maya Anggraini, M.Si", "Riko Pratama, S.Kom"),
        ("19", "Babel", "Dra. Siti Aminah, M.Si", "", "", ""),  # Tanpa Cc
        ("21", "Kepri", "Ir. Dadang Hendrawan, M.Si", "Tri Wahyuni, S.Stat", "Hendri Saputra, S.E", ""),
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daftar_Kontak"

    # Baris 1: Judul
    ws.append(["", "DAFTAR KONTAK DISTRIBUSI DOKUMEN & NASKAH STATISTIK BPS PROVINSI"])
    # Baris 2 & 3: Kosong
    ws.append([""])
    ws.append([""])

    # Baris 4 (Index row 3): Header Level 1
    header_l1 = [
        "",  # Col 0 (A): Margin
        "No",  # Col 1 (B)
        "Kode",  # Col 2 (C)
        "Provinsi",  # Col 3 (D)
        "Kepala Prov",  # Col 4 (E)
        "Kepala Prov",  # Col 5 (F) - To
        "Ketua Tim Statistik Produksi",  # Col 6 (G)
        "Ketua Tim Statistik Produksi",  # Col 7 (H) - Cc 1
        "Koordinator Statistik Industri",  # Col 8 (I)
        "Koordinator Statistik Industri",  # Col 9 (J) - Cc 2
        "Staf Seksi Statistik Industri",  # Col 10 (K)
        "Staf Seksi Statistik Industri",  # Col 11 (L) - Cc 3
    ]
    ws.append(header_l1)

    # Baris 5 (Index row 4): Header Level 2
    header_l2 = [
        "",
        "-1",
        "-2",
        "-3",
        "Nama",
        "E-mail",
        "Nama",
        "Email",
        "Nama",
        "Email",
        "Nama",
        "Email",
    ]
    ws.append(header_l2)

    # Baris 6 (Index row 5): Numbering
    header_l3 = [
        "",
        "(1)",
        "(2)",
        "(3)",
        "(4)",
        "(5)",
        "(6)",
        "(7)",
        "(8)",
        "(9)",
        "(10)",
        "(11)",
    ]
    ws.append(header_l3)

    # Style Header
    header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    header_font = Font(bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="B0C4DE"),
        right=Side(style="thin", color="B0C4DE"),
        top=Side(style="thin", color="B0C4DE"),
        bottom=Side(style="thin", color="B0C4DE"),
    )

    for row in ws.iter_rows(min_row=4, max_row=6, min_col=2, max_col=12):
        for cell in row:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    # Isi Data mulai Baris 7 (Index row 6)
    email_counter = 0
    for idx, (kode, prov, nama_kepala, nama_pjk, nama_koord, nama_staf) in enumerate(provinces_data, start=1):
        # Email To (Kepala Prov)
        to_email = active_emails[(idx - 1) % len(active_emails)].strip()

        # Email Cc 1 (Ketua Tim)
        if nama_pjk:
            cc1_email = active_emails[(email_counter + 1) % len(active_emails)].strip()
            email_counter += 1
        else:
            cc1_email = ""

        # Email Cc 2 (Koordinator)
        if nama_koord:
            cc2_email = active_emails[(email_counter + 2) % len(active_emails)].strip()
            email_counter += 1
        else:
            cc2_email = ""

        # Email Cc 3 (Staf)
        if nama_staf:
            cc3_email = active_emails[(email_counter) % len(active_emails)].strip()
            email_counter += 1
        else:
            cc3_email = ""

        row_data = [
            "",  # Col A
            idx,  # Col B (1)
            kode,  # Col C (2)
            prov,  # Col D (3)
            nama_kepala,  # Col E (4)
            to_email,  # Col F (5)
            nama_pjk,  # Col G (6)
            cc1_email,  # Col H (7)
            nama_koord,  # Col I (8)
            cc2_email,  # Col J (9)
            nama_staf,  # Col K (10)
            cc3_email,  # Col L (11)
        ]
        ws.append(row_data)

    # Format border & alignment untuk baris data
    for row in ws.iter_rows(min_row=7, max_row=len(provinces_data) + 6, min_col=2, max_col=12):
        for col_idx, cell in enumerate(row, start=2):
            cell.border = thin_border
            if col_idx in [2, 3]:
                cell.alignment = Alignment(horizontal="center")

    # Set column widths
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 25
    ws.column_dimensions["F"].width = 28
    ws.column_dimensions["G"].width = 24
    ws.column_dimensions["H"].width = 28
    ws.column_dimensions["I"].width = 24
    ws.column_dimensions["J"].width = 28
    ws.column_dimensions["K"].width = 24
    ws.column_dimensions["L"].width = 28

    excel_path = os.path.join(base_dir, "Daftar_Email_Dummy_10.xlsx")
    wb.save(excel_path)
    print(f"[OK] File Excel resmi BPS dibuat: {excel_path}")

    # Buat folder dummy_lampiran
    lampiran_dir = os.path.join(base_dir, "dummy_lampiran")
    os.makedirs(lampiran_dir, exist_ok=True)

    # 1. Lampiran Nasional (PDF dan XLSX)
    nat_pdf = "00 Naskah_IMK Tw 1 2024.pdf"
    nat_xlsx = "00 Pedoman_Teknis_Pelaporan.xlsx"
    create_dummy_pdf(os.path.join(lampiran_dir, nat_pdf), nat_pdf)
    create_dummy_text_file(os.path.join(lampiran_dir, nat_xlsx), "Dummy Excel Nasional Content")
    print(f"  + Lampiran Nasional: {nat_pdf}, {nat_xlsx}")

    # 2. Lampiran Spesifik Provinsi
    kodes = ["11", "12", "13", "14", "15", "16", "17", "18", "19", "21"]
    for k in kodes:
        fname = f"{k} Naskah_IMK Tw 1 2024.pdf"
        create_dummy_pdf(os.path.join(lampiran_dir, fname), fname)

    # 3. Lampiran multi-format (DOCX & ZIP)
    extra_aceh_docx = "11 Data_Tambahan_Aceh.docx"
    extra_sumut_zip = "12 Arsip_Tabel_Sumut.zip"
    create_dummy_text_file(os.path.join(lampiran_dir, extra_aceh_docx), "Dummy Word Aceh Content")
    create_dummy_text_file(os.path.join(lampiran_dir, extra_sumut_zip), "Dummy ZIP Sumut Content")
    print(f"  + Multi-format Lampiran: {extra_aceh_docx}, {extra_sumut_zip}")


if __name__ == "__main__":
    generate_dummy_data()
