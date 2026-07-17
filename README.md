# Project-1-LAPORAN

## Deskripsi
Repositori ini berisi dokumen dan script pendukung untuk penyusunan **Laporan Anggota Dewan**
atas nama **Ir. Andreas Eddy Susetyo, M.M. (A-220)**.

Cakupan dokumen dalam project ini meliputi beberapa jenis laporan kerja, antara lain:
- Laporan Rumah Aspirasi
- Laporan Kunjungan Pengawasan
- Laporan 4 Pilar
- KUNDAPIL
- Laporan Reses

Screening Berita Ekonomi (nasional & global) kini dikelola di repo terpisah:
[Screening-Berita-Ekonomi](https://github.com/puguht-ai/Screening-Berita-Ekonomi).

## Struktur File Utama (saat ini)

### Dokumen
- `FINAL KUNDAPIL VI 15_17 MEI 2026.pdf`
- `7Pagi_Laporan Sos 4 Pilar.pdf`
- `7Pagi_Ringkasan Kegiatan.pdf`
- `7Pagi_Rilis 4 pilar.pdf`
- `7 Pagi presensii ST. Yohanes Pemandi.pdf`
- `Daftar_Hadir_4_Pilar.docx`

Laporan final bulanan (Rumah Aspirasi, Kunjungan Pengawasan, KUNDAPIL, Reses, dst) **tidak
disimpan di repo git** karena ukurannya besar — semua diarsipkan per kategori/bulan di
`D:\DATA PAPA\LAPORAN\` (lihat subfolder `RUMAH ASPIRASI`, `KUNWAS`, `KUNDAPIL`, `RESES`,
`4 PILAR`, `BLACBOX POOL`).

### Script Python
- `create_daftar_hadir_4pilar.py`
- `create_kundapil_docx.py`

### File Konfigurasi/Konteks
- `BLACKBOX.md`
- `4_pilar_rilis.txt`

## Cara Menjalankan Script
Jalankan dari root repo:

```bash
cd C:\Users\USER
python <nama_script.py>
```

Contoh:

```bash
python create_kundapil_docx.py
python create_daftar_hadir_4pilar.py
```

## Konvensi Output
File hasil generate/final disimpan ke arsip terstruktur di:

`D:\DATA PAPA\LAPORAN\<KATEGORI>\<BULAN/PERIODE>\`

Gunakan format nama file yang konsisten berdasarkan jenis laporan dan tanggal.

## Catatan Workflow
1. Siapkan data sumber (agenda kegiatan, poin pendahuluan, foto, referensi berita).
2. Jalankan script sesuai kebutuhan jenis laporan.
3. Lakukan pengecekan isi akhir (tanggal, nama, jabatan, lokasi, referensi).
4. Simpan hasil final ke folder arsip `D:\DATA PAPA\LAPORAN\...` (bukan ke repo git).
5. Commit perubahan ke Git setelah validasi dokumen selesai (script/dokumen pendukung saja).
