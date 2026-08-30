# Generator Laporan Susunan Bal Stock Gudang

Skrip Python otomatis untuk menghasilkan laporan layout/susunan bal stock gudang dalam format spreadsheet Excel (`.xlsx`) menggunakan library `openpyxl`.

## 📦 Fitur
- **Layout Multi-Blok**: Mengatur pemetaan 16 blok gudang dengan saf dan tingkat susunan bal (Tingkat 1 - 7).
- **Format Cetak Rapi**: Disesuaikan dengan batas halaman cetak, garis pembatas (border), dan hierarki visual yang jelas.
- **Hemat Tinta & Kontras Tinggi**: Tampilan tabel bersih dan mudah dibaca saat dicetak fisik.

## 🚀 Persyaratan
- Python 3.8+
- Library `openpyxl`

Instalasi dependency:
```bash
pip install openpyxl
```

## 🛠️ Cara Menjalankan
Jalankan skrip generator:
```bash
python create_report.py
```
File laporan Excel akan dibuat di direktori yang sama.
