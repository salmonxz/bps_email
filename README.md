# Panduan Penggunaan: Automated Provincial Email Dispatcher (BPS RI)

Aplikasi desktop portabel untuk mengotomatiskan pengiriman email personalisasi dokumen/naskah statistik ke **38 BPS Provinsi** dengan pencocokan lampiran dinamis berbasis kode wilayah.

---

## 🌟 Fitur Utama

1. **Pembacaan Excel Fleksibel & Logika Fallback Alamat 'To'**:
   - Mendukung format Excel resmi BPS dengan 4 kontak:
     * Col 5: Kepala Prov
     * Col 7: Ketua Tim Statistik Produksi
     * Col 9: Koordinator Statistik Industri
     * Col 11: Staf Seksi Statistik Industri
   - **Logika Penentuan 'To'**: Mengutamakan Kepala Prov. Jika kosong/NaN, otomatis mengambil email valid pertama sesuai urutan prioritas: **Ketua Tim $\rightarrow$ Koordinator $\rightarrow$ Staf**.
   - **Alokasi Cc & Bcc**: Sisa email dialokasikan ke **Cc** (Ketua Tim & Koordinator) dan **Bcc** (Staf Seksi). Pengiriman Bcc dimasukkan ke envelope SMTP (`to_addrs`), namun header `Bcc` disembunyikan agar alamat staf tetap terlindungi (*blind*).
   - Baris hanya ditandai error jika seluruh kolom email pada baris tersebut kosong.

2. **Pencocokan Berkas Lampiran Dinamis (Attachment Engine)**:
   - **Berkas Nasional (`00`)**: Setiap file dengan awalan/akhiran `00` atau kata `Nasional` otomatis dilampirkan ke **seluruh** 38 provinsi.
   - **Berkas Daerah (`{kode}`)**: File dengan kode provinsi terkait (contoh: `11_Naskah_Aceh.pdf`, `Tabel Profil IMK - 16.xlsx`, atau folder `11/`) hanya dilampirkan ke provinsi tersebut. Mendukung multi-file dan folder.

3. **Template Subjek & Pesan Dinamis & Area Input Responsif**:
   - Area teks Subjek Email diperbesar dan proporsional.
   - Area teks Isi Pesan Email (tinggi 10 baris) dilengkapi Vertical Scrollbar dan layout responsif (`sticky="nsew"`).
   - Tag dinamis yang didukung (format `{{tag}}`, `<<tag>>`, maupun `{tag}`):
     * `{{provinsi}}` atau `<<provinsi>>`: Nama provinsi (contoh: *Aceh*, *Jawa Barat*)
     * `{{kode}}` atau `<<kode>>`: Kode wilayah 2 digit (contoh: *11*, *32*)
     * `{{nama_kepala}}`: Nama Kepala BPS Provinsi
     * `{{nama_ketua}}`: Nama Ketua Tim
   - Tombol klik cepat untuk menyisipkan tag dan label petunjuk ringkas di bawah template.

4. **Tabel Pratinjau & Verifikasi Pre-Flight (Tab 2)**:
   - Menampilkan ringkasan seluruh target: No, Kode, Provinsi, kolom email sumber, jumlah lampiran nasional/daerah, dan status validasi.
   - Panel detail inspeksi menampilkan alokasi penerima secara transparan (**To / Cc / Bcc**) serta pratinjau hasil penggantian tag dinamis pada Subjek dan Isi Pesan.
   - Indikator warna:
     - 🟢 **Hijau (Siap Kirim)**: Data lengkap / memiliki fallback email valid dan berkas lampiran cocok (otomatis tercentang `[✓]`).
     - 🟡 **Kuning (Perhatian)**: Berkas daerah belum ditemukan (hanya berkas nasional).
     - 🔴 **Merah (Error)**: Seluruh email kosong atau format email tidak valid.

5. **Simulasi (Dry Run) & Pengiriman Aman (Tab 3)**:
   - **Simulasi (Dry Run)**: Menjalankan uji coba alur tanpa mengirim email sungguhan, memverifikasi keterikatan To, Cc, Bcc, dan berkas.
   - **Konfirmasi Pengiriman**: Dialog konfirmasi sebelum pengiriman riil dimulai.
   - **Toleransi Error (Fault-Tolerant)**: Pengiriman tetap berlanjut ke provinsi berikutnya jika terjadi kendala pada satu baris dan mencatat log *real-time*.
   - Tombol **Hentikan Proses** sewaktu-waktu.

---

## 🚀 Cara Menjalankan

### Opsi 1: Menggunakan File Portable Executable (.exe) - Tanpa Perlu Install Python
1. Buka folder `dist/`.
2. Klik dua kali pada `BPS_Email_Dispatcher.exe`.
3. Aplikasi langsung berjalan tanpa memerlukan instalasi software tambahan atau hak administrator.

### Opsi 2: Menjalankan dari Source Code (Python)
1. Pastikan Python 3.10+ sudah terpasang.
2. Pasang dependensi:
   ```bash
   pip install -r requirements.txt
   ```
3. Jalankan aplikasi:
   ```bash
   python app.py
   ```

---

## ⚙️ Petunjuk Konfigurasi Email (Gmail SMTP)

Jika Anda menggunakan akun Gmail pengirim:
1. Pastikan **Verifikasi 2 Langkah (2-Step Verification)** aktif pada Akun Google Anda.
2. Buka [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords).
3. Buat sandi aplikasi baru dengan nama aplikasi (misal: "BPS Dispatcher").
4. Salin sandi 16 karakter yang diberikan Google (misal: `abcd efgh ijkl mnop`).
5. Masukkan sandi tersebut pada kolom **Password / App Password** di aplikasi.
6. Klik tombol **⚡ Uji Koneksi SMTP** untuk memastikan status "Berhasil".

---

## 🧪 Data Uji Coba (Sample Data)

Folder `sample_data/` telah disediakan untuk latihan:
- `sample_data/sample_kontak_bps.xlsx`: Contoh database kontak dengan format multi-level header BPS.
- `sample_data/attachments/`: Kumpulan berkas mock PDF (`00_...`, `11_...`, `12_...`, `31_...`).
