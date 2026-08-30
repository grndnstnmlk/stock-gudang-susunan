---
name: stock-gudang
description: >-
  Keahlian komprehensif untuk memproses, merekonsiliasi, dan mengotomasi laporan
  susunan bal stock gudang tembakau ke dalam format spreadsheet Excel (.xlsx) siap cetak A4.
  Mencakup pemetaan spasial blok gudang (Saf 1 Utara - Saf N Selatan), rekonsiliasi data master
  No Gudang dengan Nomor Barcode, Grade, dan Berat (Kg), penanganan anomali bal, layout multi-sheet,
  serta konfigurasi print setup anti-border tebal.
---

# Skill: Otomasi & Rekonsiliasi Stock Gudang Susunan Bal (`stock-gudang`)

Skill ini menyediakan standar operasional prosedur, arsitektur data, fungsi helper Python, dan pedoman desain untuk mengelola data susunan bal gudang tembakau.

---

## 1. Arsitektur Data & Model Spasial Gudang

### Hierarki Fisik Gudang:
1. **Blok (1 s/d 16)**: Area penumpukan bal tembakau di gudang.
2. **Saf (Kolom)**: Baris susunan dari sisi ke sisi.
   - Blok 01 – 05: 6 Saf
   - Blok 06: 5 Saf
   - Blok 07 – 16: 4 Saf
3. **Tingkat (7 Tingkat Kapasitas)**:
   - `T7 (Atas)`: Slot kosong (disiapkan untuk bal baru)
   - `T6`: Slot kosong
   - `T5`: Slot kosong
   - `T4`: Bal aktif tingkat 4
   - `T3`: Bal aktif tingkat 3
   - `T2`: Bal aktif tingkat 2
   - `T1 (Dasar)`: Bal aktif dasar

### Aturan Orientasi Arah:
- **Header Kolom**: Gunakan nama saf murni (`Saf 1`, `Saf 2`, `Saf 3`, dst.).
- **Banner Judul Blok**: Wajib mencantumkan arah hadap saf secara dinamis:
  - `BLOK 01 (Saf 1 Utara - Saf 6 Selatan)`
  - `BLOK 06 (Saf 1 Utara - Saf 5 Selatan)`
  - `BLOK 07 (Saf 1 Utara - Saf 4 Selatan)`

---

## 2. Rekonsiliasi Master Data & Lookup Barcode/Kg

Data nomor bal fisik di gudang merepresentasikan **No Gudang (No Gud)** yang dicocokkan ke file master `Dokumen_Rekap_NoGud_Barkot_Kg.xlsx`.

### Logika Smart Loader:
```python
def load_master_database(master_filepath='Dokumen_Rekap_NoGud_Barkot_Kg.xlsx'):
    master_dict = {}
    wb_master = openpyxl.load_workbook(master_filepath, data_only=True)
    
    def update_entry(no_gud, barkot, grade, kg, status, ket):
        if no_gud is None:
            return
        k = str(no_gud).strip()
        b_str = '' if barkot in [None, '', '-', ''] else str(barkot).strip()
        kg_val = '' if kg in [None, '', '-'] else kg
        grade_str = '' if grade in [None, '', '-'] else str(grade).strip()
        status_str = '' if status in [None, '', '-'] else str(status).strip()

        if k not in master_dict:
            master_dict[k] = {
                'no_gud': no_gud,
                'barkot': b_str,
                'grade': grade_str,
                'kg': kg_val,
                'status': status_str,
                'ket': str(ket or '').strip()
            }
        else:
            # Utamakan record dengan barcode valid (bukan status OUT)
            if b_str and not master_dict[k]['barkot']:
                master_dict[k] = {
                    'no_gud': no_gud,
                    'barkot': b_str,
                    'grade': grade_str or master_dict[k]['grade'],
                    'kg': kg_val if kg_val != '' else master_dict[k]['kg'],
                    'status': status_str or master_dict[k]['status'],
                    'ket': str(ket or '').strip()
                }

    # Baca sheet harian & master
    for sname in wb_master.sheetnames:
        if sname not in ['Semua Data (Master)', 'Ringkasan Per Tanggal']:
            ws = wb_master[sname]
            for r in range(5, ws.max_row + 1):
                update_entry(ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value, ws.cell(r, 7).value)

    if 'Semua Data (Master)' in wb_master.sheetnames:
        ws_m = wb_master['Semua Data (Master)']
        for r in range(5, ws_m.max_row + 1):
            update_entry(ws_m.cell(r, 3).value, ws_m.cell(r, 4).value, ws_m.cell(r, 5).value, ws_m.cell(r, 6).value, ws_m.cell(r, 7).value, ws_m.cell(r, 8).value)

    # Custom Override untuk koreksi manual spesifik
    custom_overrides = {
        '259': {'no_gud': 259, 'barkot': '30164', 'grade': '55', 'kg': 55, 'status': 'SELESAI', 'ket': ''},
    }
    for k, v in custom_overrides.items():
        master_dict[k] = v

    return master_dict
```

---

## 3. Standar 4 Sheet Output

Spreadsheet yang dihasilkan (`Stock_Susunan_Bal_*.xlsx` dan `c.xlsx`) harus menyertakan 4 tampilan standar:

### Sheet 1: `Stock Lengkap (NoGud+Barkot+Kg)`
- **Tujuan**: Tampilan visual lengkap operasional gudang.
- **Layout per Saf**: Setiap Saf memiliki 3 sub-kolom: `No Gud` | `Barkot` | `Kg`.
- **Page Break Budgeting (A4 Landscape)**:
  - **Halaman 1**: 2 Tabel (Blok 01 & 02) + Judul Utama (`STOCK TEMBAKAU DJARUM`) + Tanggal.
  - **Halaman 2 s/d 5**: Masing-masing tepat **3 Tabel** per lembar.
  - **Halaman 6**: 2 Tabel terakhir (Blok 15 & 16).
  - Page break ditambahkan pada blok: `[2, 5, 8, 11, 14]`.

### Sheet 2: `Stock Ringkas (No Gud Saja)`
- **Tujuan**: Pengecekan cepat susunan nomor bal fisik saat berjalan di lorong gudang.
- **Layout**: 1 kolom per Saf dengan font jumbo (14pt bold).
- **Page Break**: Tepat 4 blok per halaman portrait (Page break setelah Blok 4, 8, 12).

### Sheet 3: `Format Grid (c.xlsx)`
- **Tujuan**: Format grid mentah data matriks 3 kolom per saf dengan label nomor blok di sebelah kanan.

### Sheet 4: `Rekapitulasi & Cari Bal`
- **Tabel 1**: Ringkasan Kapasitas (7 Tkt), Terisi (Bal), Kosong (Slot), dan Total Berat (Kg) per blok.
- **Tabel 2**: Flat Database 300 bal dengan AutoFilter Excel:
  `No | No Gudang | Nomor Barkot | Grade | Berat (Kg) | Blok | Tingkat | Saf / Kolom | Status Bal`

---

## 4. Aturan Styling & Anti-Border Tebal (Print-Ready)

1. **Garis Border Tipis & Elegan**:
   - **JANGAN** gunakan `Side(border_style='thin', color='000000')` murni pada semua sel karena akan dirender tebal/bold saat print preview.
   - **GUNAKAN**:
     ```python
     border_side_clean = Side(border_style='thin', color='BFBFBF') # Light gray crisp border
     border_all = Border(left=border_side_clean, right=border_side_clean, top=border_side_clean, bottom=border_side_clean)
     ```
2. **Palette Hemat Tinta**:
   - Judul Utama: `Font(name='Segoe UI', size=16, bold=True)`
   - Banner Blok: Fill `PatternFill('E8EFF5')`
   - Header Kolom Saf: Fill `PatternFill('F2F5F8')`
   - Tingkat Label: Fill `PatternFill('F9FAFB')`
   - Data Sel: Background putih `FFFFFF` dengan teks hitam tegas.
3. **Konfigurasi Cetak Presisi**:
   ```python
   def configure_a4_print(ws, orientation='landscape', fit_w=1, fit_h=0):
       ws.page_setup.paperSize = ws.PAPERSIZE_A4
       ws.page_setup.orientation = orientation
       ws.sheet_properties.pageSetUpPr.fitToPage = True
       ws.page_setup.fitToWidth = fit_w
       ws.page_setup.fitToHeight = fit_h
       ws.page_margins.left = 0.25
       ws.page_margins.right = 0.25
       ws.page_margins.top = 0.3
       ws.page_margins.bottom = 0.3
       ws.print_options.horizontalCentered = True
   ```
