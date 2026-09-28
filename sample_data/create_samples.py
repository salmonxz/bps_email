"""
Script to generate sample Excel with multi-level header and sample dummy PDF attachments.
"""

import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def create_sample_files():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(base_dir, exist_ok=True)
    att_dir = os.path.join(base_dir, "attachments")
    os.makedirs(att_dir, exist_ok=True)

    # 1. Create sample Excel file with multi-level merged header
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kontak_BPS"

    # Title row
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = "DAFTAR KONTAK BPS PROVINSI UNTUK DISTRIBUSI DOKUMEN NASKAH STATISTIK"
    title_cell.font = Font(bold=True, size=12, color="1F497D")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Header Level 1 (Row 2)
    headers_l1 = [
        ("A2", "No"),
        ("B2", "Kode"),
        ("C2", "Provinsi"),
        ("D2", "Kepala BPS Provinsi"),
        ("E2", ""),
        ("F2", "Ketua Tim / PJK"),
        ("G2", ""),
    ]
    for pos, text in headers_l1:
        ws[pos] = text

    ws.merge_cells("A2:A3")
    ws.merge_cells("B2:B3")
    ws.merge_cells("C2:C3")
    ws.merge_cells("D2:E2")
    ws.merge_cells("F2:G2")

    # Header Level 2 (Row 3)
    ws["D3"] = "Nama"
    ws["E3"] = "E-mail"
    ws["F3"] = "Nama"
    ws["G3"] = "Email"

    header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    header_font = Font(bold=True, size=10)
    thin_border = Border(
        left=Side(style="thin", color="B0C4DE"),
        right=Side(style="thin", color="B0C4DE"),
        top=Side(style="thin", color="B0C4DE"),
        bottom=Side(style="thin", color="B0C4DE"),
    )

    for row in ws["A2:G3"]:
        for cell in row:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    # Sample Data (Row 4 onwards)
    data = [
        (1, "11", "Aceh", "Dr. Ahmad, M.Si", "ahmad.aceh@bps.go.id", "Budi Santoso, S.Si", "budi.aceh@bps.go.id"),
        (2, "12", "Sumatera Utara", "Dr. Nurul Huda, M.Si", "nurul.sumut@bps.go.id", "Rian Pratama, S.Stat", "rian.sumut@bps.go.id"),
        (3, "13", "Sumatera Barat", "Ir. Herianto, M.E", "herianto.sumbar@bps.go.id", "", ""), # No Cc
        (4, "31", "DKI Jakarta", "Dr. Nur Indah, M.Si", "nur.jakarta@bps.go.id", "Farhan Maulana, S.Tr.Stat", "farhan.jakarta@bps.go.id"),
        (5, "32", "Jawa Barat", "Drs. Agus Sudarto, M.Si", "agus.jabar@bps.go.id", "Siti Rahmawati, S.Si", "siti.jabar@bps.go.id"),
        (6, "33", "Jawa Tengah", "Ir. Dadang, M.Si", "dadang.jateng@bps.go.id", "Tri Wahyuni, S.Stat", "tri.jateng@bps.go.id"),
        (7, "35", "Jawa Timur", "Dr. Zulkipli, M.Si", "zulkipli.jatim@bps.go.id", "Eko Prasetyo, S.Kom", "eko.jatim@bps.go.id"),
        (8, "73", "Sulawesi Selatan", "Hj. Sunarti, M.Si", "sunarti.sulsel@bps.go.id", "Andi Baso, S.Stat", "andi.sulsel@bps.go.id"),
        (9, "81", "Maluku", "Dr. Chris Pattikawa, M.Si", "chris.maluku_invalid_email", "Markus, S.Si", "markus@bps.go.id"), # Invalid email test
    ]

    for row_idx, record in enumerate(data, start=4):
        for col_idx, val in enumerate(record, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            if col_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="center")

    # Add notes at bottom
    footer_row = len(data) + 5
    ws.cell(row=footer_row, column=1, value="* Catatan: Data kontak bersumber dari Biro Humas dan Protokol BPS.")
    ws.cell(row=footer_row, column=1).font = Font(italic=True, size=9, color="7F7F7F")

    # Column widths
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 26
    ws.column_dimensions["G"].width = 28

    excel_path = os.path.join(base_dir, "sample_kontak_bps.xlsx")
    wb.save(excel_path)
    print(f"Sample Excel created: {excel_path}")

    # 2. Create sample dummy PDF files
    mock_files = [
        "00_Naskah_Nasional_2026.pdf",
        "00_Pedoman_Teknis_Pelaporan.pdf",
        "11_Naskah_Statistik_Aceh.pdf",
        "11_Lampiran_Tabel_Aceh.pdf",      # multi-attachment for Aceh
        "12_Naskah_Statistik_Sumut.pdf",
        "13_Naskah_Statistik_Sumbar.pdf",
        "31_Naskah_Statistik_DKI_Jakarta.pdf",
        "32_Naskah_Statistik_Jawa_Barat.pdf",
        "33_Naskah_Statistik_Jateng.pdf",
        # Note: 35 Jawa Timur deliberately has no province-specific file (to test national-only case)
        # Note: 73 has 73_Naskah_Sulsel.pdf
        "73_Naskah_Sulsel.pdf",
    ]

    for fname in mock_files:
        fpath = os.path.join(att_dir, fname)
        # Write minimal valid dummy PDF structure
        with open(fpath, "wb") as f:
            f.write(b"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Count 1/Kids[3 0 R]>>endobj 3 0 obj<</Type/Page/MediaBox[0 0 595 842]/Parent 2 0 R/Resources<<>>>>endobj\nxref\n0 4\n0000000000 65535 f \n0000000009 00000 n \n0000000052 00000 n \n0000000102 00000 n \ntrailer<</Size 4/Root 1 0 R>>\nstartxref\n178\n%%EOF\n")
        print(f"Created dummy PDF: {fname}")

    print("All sample data ready.")


if __name__ == "__main__":
    create_sample_files()
