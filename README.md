# Project-1-LAPORAN

## Deskripsi
Repositori ini berisi dokumen dan script pendukung untuk penyusunan **Laporan Anggota Dewan**
atas nama **Ir. Andreas Eddy Susetyo, M.M. (A-220)**.

Cakupan dokumen dalam project ini meliputi beberapa jenis laporan kerja, antara lain:
- Laporan Rumah Aspirasi
- Laporan Kunjungan Pengawasan
- Laporan 4 Pilar
- KUNDAPIL
- Screening Berita Ekonomi (nasional & global)

## Struktur File Utama (saat ini)

### Dokumen
- `Final Laporan Rumah Aspirasi Mei 2026.docx`
- `Final_Laporan_Kunjungan Pengawasan_Ke_1_2026.docx`
- `FINAL KUNDAPIL VI 15_17 MEI 2026.pdf`
- `7Pagi_Laporan Sos 4 Pilar.pdf`
- `7Pagi_Ringkasan Kegiatan.pdf`
- `7Pagi_Rilis 4 pilar.pdf`
- `7 Pagi presensii ST. Yohanes Pemandi.pdf`
- `Daftar_Hadir_4_Pilar.docx`

### Script Python
- `create_daftar_hadir_4pilar.py`
- `create_kundapil_docx.py`
- `create_screening_docx.py`
- `create_screening_global.py`

### File Konfigurasi/Konteks
- `BLACKBOX.md`
- `4_pilar_rilis.txt`

## Cara Menjalankan Script
Jalankan dari folder repo:

```bash
cd C:\Users\USER\Project-1-LAPORAN
python <nama_script.py>
```

Contoh:

```bash
python create_screening_docx.py
python create_screening_global.py
```

## Konvensi Output
Sesuai preferensi kerja user, semua file output baru disarankan disimpan ke:

`D:\BLACKBOX DATA`

Gunakan format nama file yang konsisten berdasarkan jenis laporan dan tanggal.

## Catatan Workflow
1. Siapkan data sumber (agenda kegiatan, poin pendahuluan, foto, referensi berita).
2. Jalankan script sesuai kebutuhan jenis laporan.
3. Lakukan pengecekan isi akhir (tanggal, nama, jabatan, lokasi, referensi).
4. Simpan hasil final ke folder output standar.
5. Commit perubahan ke Git setelah validasi dokumen selesai.
