import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

# ==============================================================================
# DATA SUSUNAN BAL GUDANG (16 BLOK)
# Headers: Saf 1 s/d Saf N (Keterangan Utara & Selatan dihapus)
# Tiap blok saat ini berisi 4 tingkat aktif: [T4, T3, T2, T1 (Dasar)]
# Disisihkan 3 tingkat kosong di atas (T7, T6, T5) untuk penambahan bal baru
# ==============================================================================
blocks_data = {
    1: {
        "title": "BLOK 01",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4", "Saf 5", "Saf 6"],
        "data": [
            [494, 497, 453, 474, 116, 310],
            [283, 23, 467, "41/509", 177, 193],
            [508, 214, 460, 462, 110, 123],
            [305, 18, 133, 144, 234, 274],
        ]
    },
    2: {
        "title": "BLOK 02",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4", "Saf 5", "Saf 6"],
        "data": [
            [164, 288, 15, 227, 416, 99],
            [183, 368, 160, 491, 469, 504],
            [493, 439, 438, 98, 171, 124],
            [125, 444, 158, 448, 108, 112],
        ]
    },
    3: {
        "title": "BLOK 03",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4", "Saf 5", "Saf 6"],
        "data": [
            [408, 465, 296, 421, 463, 206],
            [222, 217, 502, 440, 487, 107],
            [180, 473, 432, 456, 282, 91],
            [230, 199, 503, 495, 32, 437],
        ]
    },
    4: {
        "title": "BLOK 04",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4", "Saf 5", "Saf 6"],
        "data": [
            [466, 447, 431, 277, 496, 468],
            [498, 501, 130, 489, 115, 19],
            [500, 436, 475, 201, 399, 111],
            [100, 492, 226, 506, 374, 370],
        ]
    },
    5: {
        "title": "BLOK 05",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4", "Saf 5", "Saf 6"],
        "data": [
            [499, 210, 309, 373, 332, 405],
            [381, 224, 397, 245, 413, 272],
            [257, 386, 246, 241, 313, 260],
            [412, 430, 451, 490, 109, 4],
        ]
    },
    6: {
        "title": "BLOK 06",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4", "Saf 5"],
        "data": [
            [211, 258, 433, 398, 122],
            [423, 389, 365, 263, 458],
            [472, 334, 315, 387, 103],
            [388, 367, 349, 336, 300],
        ]
    },
    7: {
        "title": "BLOK 07",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [320, 384, 371, 352],
            [377, 326, 339, 375],
            [380, 350, 330, 329],
            [318, 427, 395, 311],
        ]
    },
    8: {
        "title": "BLOK 08",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [325, 379, 324, 385],
            [425, 369, 383, 393],
            [340, 382, 333, 265],
            [267, 271, 90, 341],
        ]
    },
    9: {
        "title": "BLOK 09",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [321, 415, 434, 278],
            [275, 197, 331, 396],
            [419, 424, 254, 378],
            [418, 426, 337, 323],
        ]
    },
    10: {
        "title": "BLOK 10",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [421, 351, 255, 392],
            [400, 390, 328, 366],
            [420, 358, 50, 406],
            [488, 344, 428, 470],
        ]
    },
    11: {
        "title": "BLOK 11",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [289, 410, 357, 404],
            [422, 347, 327, 355],
            [154, 505, 259, 228],
            [247, 414, 346, 338],
        ]
    },
    12: {
        "title": "BLOK 12",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [394, 322, 335, 264],
            [312, 316, 407, 231],
            [445, 403, 361, 455],
            [342, 169, 216, 459],
        ]
    },
    13: {
        "title": "BLOK 13",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [173, 409, 442, 429],
            [140, 141, 178, 401],
            [233, 461, 446, 132],
            [63, 168, 372, 402],
        ]
    },
    14: {
        "title": "BLOK 14",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [450, 256, 376, 175],
            [49, 411, 319, 452],
            [219, 232, 454, 457],
            [273, 449, 441, 181],
        ]
    },
    15: {
        "title": "BLOK 15",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [179, 574, 150, 343],
            [559, 572, 356, 285],
            [575, 223, 220, 435],
            [276, 564, 149, 443],
        ]
    },
    16: {
        "title": "BLOK 16",
        "headers": ["Saf 1", "Saf 2", "Saf 3", "Saf 4"],
        "data": [
            [546, 573, 64, 560],
            [547, 554, 540, 538],
            [542, 557, 544, 561],
            [555, 558, 549, 307],
        ]
    },
}

all_tingkat_info = [
    ('T7 (Atas)', True),
    ('T6', True),
    ('T5', True),
    ('T4', False),
    ('T3', False),
    ('T2', False),
    ('T1 (Dasar)', False)
]

wb = openpyxl.Workbook()

# ==============================================================================
# STYLE DEFINITIONS (HEMAT TINTA, FONT BESAR, BORDER TEGAS)
# ==============================================================================
font_main_title = Font(name='Segoe UI', size=16, bold=True, color='000000')
font_subtitle = Font(name='Segoe UI', size=11, bold=True, color='333333')
font_date = Font(name='Segoe UI', size=10, italic=True, color='333333')

font_block_hdr = Font(name='Segoe UI', size=12, bold=True, color='000000')
font_subhdr = Font(name='Segoe UI', size=11, bold=True, color='000000')
font_data_jumbo = Font(name='Segoe UI', size=14, bold=True, color='000000')
font_tingkat_large = Font(name='Segoe UI', size=11, bold=True, color='000000')
font_data_regular = Font(name='Segoe UI', size=11, bold=False, color='000000')

# Fills hemat tinta (terang dan bersih)
fill_banner_eco = PatternFill(start_color='E8EFF5', end_color='E8EFF5', fill_type='solid') # Soft tint
fill_subhdr_eco = PatternFill(start_color='F2F5F8', end_color='F2F5F8', fill_type='solid')
fill_tingkat_eco = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
fill_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

# Solid black thin border
border_side_black = Side(border_style='thin', color='000000')
border_all_black = Border(left=border_side_black, right=border_side_black, top=border_side_black, bottom=border_side_black)

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_wrap_center = Alignment(horizontal='center', vertical='center', wrap_text=True)

def configure_a4_print(ws, orientation='portrait', fit_w=1, fit_h=4):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = orientation
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True
    ws.print_options.gridLines = False

# ==============================================================================
# SHEET 1: STOCK (PORTRAIT - TEPAT 4 BLOK PER HALAMAN TANPA TERPOTONG)
# Judul: STOCK | Tanggal: Minggu 30/8/2026 | Header: Saf 1 s/d Saf N murni
# ==============================================================================
ws_portrait = wb.active
ws_portrait.title = 'Stock (Cetak Portrait 4 Blok)'
ws_portrait.views.sheetView[0].showGridLines = True
configure_a4_print(ws_portrait, orientation='portrait', fit_w=1, fit_h=4)

# Title & Date at top of Sheet 1
ws_portrait['B2'] = 'STOCK'
ws_portrait['B2'].font = font_main_title
ws_portrait['B2'].alignment = align_left
ws_portrait.row_dimensions[2].height = 20

ws_portrait['B3'] = 'Minggu 30/8/2026'
ws_portrait['B3'].font = font_subtitle
ws_portrait['B3'].alignment = align_left
ws_portrait.row_dimensions[3].height = 16

ws_portrait.row_dimensions[4].height = 6

curr_r = 5
for b_idx in range(1, 17):
    b = blocks_data[b_idx]
    num_cols = len(b["headers"])
    end_col = 2 + num_cols
    
    # Block Header Banner (BLOK 01, BLOK 02, ...)
    ws_portrait.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=end_col)
    hdr_cell = ws_portrait.cell(curr_r, 2)
    hdr_cell.value = b['title']
    hdr_cell.font = font_block_hdr
    hdr_cell.alignment = align_center
    for c in range(2, end_col + 1):
        ws_portrait.cell(curr_r, c).fill = fill_banner_eco
        ws_portrait.cell(curr_r, c).border = border_all_black
    ws_portrait.row_dimensions[curr_r].height = 21
    
    # Table Column Headers (Tingkat, Saf 1, Saf 2, ...)
    curr_r += 1
    t_hdr = ws_portrait.cell(curr_r, 2, 'Tingkat')
    t_hdr.font = font_subhdr
    t_hdr.fill = fill_subhdr_eco
    t_hdr.alignment = align_center
    t_hdr.border = border_all_black
    
    for c_idx, h_name in enumerate(b["headers"]):
        col_num = 3 + c_idx
        c_cell = ws_portrait.cell(curr_r, col_num, h_name)
        c_cell.font = font_subhdr
        c_cell.fill = fill_subhdr_eco
        c_cell.alignment = align_center
        c_cell.border = border_all_black
    ws_portrait.row_dimensions[curr_r].height = 20
    
    # Data Rows (7 Tiers)
    data_idx = 0
    for t_name, is_empty in all_tingkat_info:
        curr_r += 1
        t_cell = ws_portrait.cell(curr_r, 2, t_name)
        t_cell.font = font_tingkat_large
        t_cell.fill = fill_tingkat_eco
        t_cell.alignment = align_center
        t_cell.border = border_all_black
        
        if is_empty:
            for c_idx in range(num_cols):
                col_num = 3 + c_idx
                d_cell = ws_portrait.cell(curr_r, col_num, "")
                d_cell.fill = fill_white
                d_cell.border = border_all_black
        else:
            row_vals = b["data"][data_idx]
            data_idx += 1
            for c_idx, val in enumerate(row_vals):
                col_num = 3 + c_idx
                d_cell = ws_portrait.cell(curr_r, col_num, val)
                d_cell.font = font_data_jumbo
                d_cell.fill = fill_white
                d_cell.alignment = align_center
                d_cell.border = border_all_black
        ws_portrait.row_dimensions[curr_r].height = 18
        
    curr_r += 1
    ws_portrait.row_dimensions[curr_r].height = 6 # Spacer
    
    # Page Break tepat setelah Blok 4, Blok 8, Blok 12 (Tepat 4 blok per halaman)
    if b_idx in [4, 8, 12]:
        ws_portrait.row_breaks.append(Break(id=curr_r))
        
    curr_r += 1

# Column widths for Portrait Sheet (Semua kolom B s/d H sama rata: 12.5)
ws_portrait.column_dimensions['A'].width = 2
for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
    ws_portrait.column_dimensions[c].width = 12.5

# ==============================================================================
# SHEET 2: CETAK A4 LANDSCAPE (2x2 - 4 BLOK PER HALAMAN DENGAN FONT JUMBO 14PT)
# ==============================================================================
ws_jumbo = wb.create_sheet(title='Cetak A4 (Landscape 2x2)')
ws_jumbo.views.sheetView[0].showGridLines = True
configure_a4_print(ws_jumbo, orientation='landscape', fit_w=1, fit_h=4)

# Title & Date at top of Sheet 2
ws_jumbo['B2'] = 'STOCK'
ws_jumbo['B2'].font = font_main_title
ws_jumbo['B2'].alignment = align_left
ws_jumbo.row_dimensions[2].height = 20

ws_jumbo['B3'] = 'Minggu 30/8/2026'
ws_jumbo['B3'].font = font_subtitle
ws_jumbo['B3'].alignment = align_left
ws_jumbo.row_dimensions[3].height = 16

ws_jumbo.row_dimensions[4].height = 6

curr_r = 5
page_groups = [
    (1, [1, 2, 3, 4]),
    (2, [5, 6, 7, 8]),
    (3, [9, 10, 11, 12]),
    (4, [13, 14, 15, 16])
]

for page_num, b_ids in page_groups:
    for row_pair_idx in range(2):
        left_id = b_ids[row_pair_idx * 2]
        right_id = b_ids[row_pair_idx * 2 + 1]
        
        b_left = blocks_data[left_id]
        b_right = blocks_data[right_id]
        
        num_cols_l = len(b_left['headers'])
        end_col_l = 2 + num_cols_l
        
        num_cols_r = len(b_right['headers'])
        end_col_r = 10 + num_cols_r
        
        # Block Banners
        ws_jumbo.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=end_col_l)
        hdr_l = ws_jumbo.cell(curr_r, 2, b_left['title'])
        hdr_l.font = font_block_hdr
        hdr_l.alignment = align_center
        for c in range(2, end_col_l + 1):
            ws_jumbo.cell(curr_r, c).fill = fill_banner_eco
            ws_jumbo.cell(curr_r, c).border = border_all_black
            
        ws_jumbo.merge_cells(start_row=curr_r, start_column=10, end_row=curr_r, end_column=end_col_r)
        hdr_r = ws_jumbo.cell(curr_r, 10, b_right['title'])
        hdr_r.font = font_block_hdr
        hdr_r.alignment = align_center
        for c in range(10, end_col_r + 1):
            ws_jumbo.cell(curr_r, c).fill = fill_banner_eco
            ws_jumbo.cell(curr_r, c).border = border_all_black
            
        ws_jumbo.row_dimensions[curr_r].height = 21
        curr_r += 1
        
        # Column Headers
        ws_jumbo.cell(curr_r, 2, 'Tingkat').font = font_subhdr
        ws_jumbo.cell(curr_r, 2).fill = fill_subhdr_eco
        ws_jumbo.cell(curr_r, 2).alignment = align_center
        ws_jumbo.cell(curr_r, 2).border = border_all_black
        for c_idx, h in enumerate(b_left['headers']):
            c_cell = ws_jumbo.cell(curr_r, 3 + c_idx, h)
            c_cell.font = font_subhdr
            c_cell.fill = fill_subhdr_eco
            c_cell.alignment = align_center
            c_cell.border = border_all_black
            
        ws_jumbo.cell(curr_r, 10, 'Tingkat').font = font_subhdr
        ws_jumbo.cell(curr_r, 10).fill = fill_subhdr_eco
        ws_jumbo.cell(curr_r, 10).alignment = align_center
        ws_jumbo.cell(curr_r, 10).border = border_all_black
        for c_idx, h in enumerate(b_right['headers']):
            c_cell = ws_jumbo.cell(curr_r, 11 + c_idx, h)
            c_cell.font = font_subhdr
            c_cell.fill = fill_subhdr_eco
            c_cell.alignment = align_center
            c_cell.border = border_all_black
            
        ws_jumbo.row_dimensions[curr_r].height = 20
        
        # 7 Tiers Data
        data_idx_l = 0
        data_idx_r = 0
        for t_name, is_empty in all_tingkat_info:
            curr_r += 1
            
            # Left Block Row
            t_l = ws_jumbo.cell(curr_r, 2, t_name)
            t_l.font = font_tingkat_large
            t_l.fill = fill_tingkat_eco
            t_l.alignment = align_center
            t_l.border = border_all_black
            if is_empty:
                for c_idx in range(num_cols_l):
                    d_c = ws_jumbo.cell(curr_r, 3 + c_idx, '')
                    d_c.fill = fill_white
                    d_c.border = border_all_black
            else:
                row_vals = b_left['data'][data_idx_l]
                data_idx_l += 1
                for c_idx, val in enumerate(row_vals):
                    d_c = ws_jumbo.cell(curr_r, 3 + c_idx, val)
                    d_c.font = font_data_jumbo
                    d_c.fill = fill_white
                    d_c.alignment = align_center
                    d_c.border = border_all_black
                    
            # Right Block Row
            t_r = ws_jumbo.cell(curr_r, 10, t_name)
            t_r.font = font_tingkat_large
            t_r.fill = fill_tingkat_eco
            t_r.alignment = align_center
            t_r.border = border_all_black
            if is_empty:
                for c_idx in range(num_cols_r):
                    d_c = ws_jumbo.cell(curr_r, 11 + c_idx, '')
                    d_c.fill = fill_white
                    d_c.border = border_all_black
            else:
                row_vals = b_right['data'][data_idx_r]
                data_idx_r += 1
                for c_idx, val in enumerate(row_vals):
                    d_c = ws_jumbo.cell(curr_r, 11 + c_idx, val)
                    d_c.font = font_data_jumbo
                    d_c.fill = fill_white
                    d_c.alignment = align_center
                    d_c.border = border_all_black
                    
            ws_jumbo.row_dimensions[curr_r].height = 23
            
        curr_r += 1
        ws_jumbo.row_dimensions[curr_r].height = 8 # Spacer
        curr_r += 1
        
    # Page Break after every page
    if page_num < 4:
        ws_jumbo.row_breaks.append(Break(id=curr_r - 1))

# Column widths for Sheet 2 (Semua kolom B-H dan J-P dibuat SAMA ukurannya: 12.0)
ws_jumbo.column_dimensions['A'].width = 2
ws_jumbo.column_dimensions['I'].width = 2
for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N', 'O', 'P']:
    ws_jumbo.column_dimensions[col].width = 12.0

# ==============================================================================
# SHEET 3: FORMAT GRID (c.xlsx)
# ==============================================================================
ws2 = wb.create_sheet(title='Format Grid (c.xlsx)')
ws2.views.sheetView[0].showGridLines = True
configure_a4_print(ws2, orientation='portrait', fit_w=1, fit_h=4)

ws2['A2'] = 'STOCK'
ws2['A2'].font = font_main_title
ws2['A3'] = 'Minggu 30/8/2026'
ws2['A3'].font = font_subtitle

grid_start_r = 5
for b_num in range(1, 17):
    matrix = blocks_data[b_num]["data"]
    num_saf = len(blocks_data[b_num]["headers"])
    
    # 3 Baris Kosong di Bagian Atas (Tingkat 7, Tingkat 6, Tingkat 5)
    for r_empty in range(3):
        for c_idx in range(num_saf):
            cell = ws2.cell(row=grid_start_r + r_empty, column=c_idx + 1, value=None)
            cell.border = border_all_black
            
    # 4 Baris Data Terisi di Bagian Bawah (Tingkat 4, Tingkat 3, Tingkat 2, Tingkat 1)
    for r_idx, row_data in enumerate(matrix):
        for c_idx, val in enumerate(row_data):
            cell = ws2.cell(row=grid_start_r + 3 + r_idx, column=c_idx + 1, value=val)
            cell.font = font_data_jumbo
            cell.alignment = align_center
            cell.border = border_all_black
            
    b_cell = ws2.cell(row=grid_start_r + 3, column=8, value=b_num)
    b_cell.font = font_main_title
    b_cell.alignment = align_center
    
    if b_num in [4, 8, 12]:
        ws2.row_breaks.append(Break(id=grid_start_r + 6))
        
    grid_start_r += 7

for col in range(1, 9):
    col_letter = get_column_letter(col)
    ws2.column_dimensions[col_letter].width = 12.0

# ==============================================================================
# SHEET 4: REKAPITULASI & PENCARIAN BAL (Database)
# ==============================================================================
ws3 = wb.create_sheet(title='Rekapitulasi & Cari Bal')
ws3.views.sheetView[0].showGridLines = True
configure_a4_print(ws3, orientation='landscape', fit_w=1, fit_h=0)

ws3['B2'] = 'STOCK'
ws3['B2'].font = font_main_title
ws3['B3'] = 'REKAPITULASI STOCK & DATABASE PENCARIAN BAL'
ws3['B3'].font = font_subtitle
ws3['B4'] = 'Minggu 30/8/2026'
ws3['B4'].font = font_date

# Table 1: Rekapitulasi per Blok
ws3['B6'] = 'Tabel 1: Rekapitulasi Stock per Blok (Kapasitas 7 Tingkat)'
ws3['B6'].font = font_subhdr

rekap_headers = ['No Blok', 'Nama Blok', 'Jumlah Saf', 'Kapasitas (7 Tkt)', 'Terisi (Bal)', 'Kosong (Slot)', 'T4', 'T3', 'T2', 'T1 (Dasar)']
for c_idx, h in enumerate(rekap_headers):
    col_num = 2 + c_idx
    c_cell = ws3.cell(7, col_num, h)
    c_cell.font = font_block_hdr
    c_cell.fill = fill_banner_eco
    c_cell.alignment = align_center
    c_cell.border = border_all_black

ws3.row_dimensions[7].height = 24

r_row = 8
for b_num in range(1, 17):
    b = blocks_data[b_num]
    num_saf = len(b["headers"])
    kap_blok = num_saf * 7
    terisi_blok = num_saf * 4
    kosong_blok = num_saf * 3
    
    ws3.cell(r_row, 2, b_num).alignment = align_center
    ws3.cell(r_row, 3, b["title"]).alignment = align_left
    ws3.cell(r_row, 4, num_saf).alignment = align_center
    ws3.cell(r_row, 5, kap_blok).alignment = align_center
    ws3.cell(r_row, 6, terisi_blok).alignment = align_center
    ws3.cell(r_row, 7, kosong_blok).alignment = align_center
    ws3.cell(r_row, 8, num_saf).alignment = align_center
    ws3.cell(r_row, 9, num_saf).alignment = align_center
    ws3.cell(r_row, 10, num_saf).alignment = align_center
    ws3.cell(r_row, 11, num_saf).alignment = align_center
    
    for c in range(2, 12):
        cell = ws3.cell(r_row, c)
        cell.font = font_data_regular
        cell.border = border_all_black
        if b_num % 2 == 0:
            cell.fill = fill_tingkat_eco
    ws3.row_dimensions[r_row].height = 20
    r_row += 1

# Total Row
ws3.cell(r_row, 2, 'TOTAL').alignment = align_center
ws3.cell(r_row, 2).font = font_subhdr
ws3.merge_cells(start_row=r_row, start_column=2, end_row=r_row, end_column=3)
ws3.cell(r_row, 4, f'=SUM(D8:D{r_row-1})').alignment = align_center
ws3.cell(r_row, 5, f'=SUM(E8:E{r_row-1})').alignment = align_center
ws3.cell(r_row, 6, f'=SUM(F8:F{r_row-1})').alignment = align_center
ws3.cell(r_row, 7, f'=SUM(G8:G{r_row-1})').alignment = align_center
ws3.cell(r_row, 8, f'=SUM(H8:H{r_row-1})').alignment = align_center
ws3.cell(r_row, 9, f'=SUM(I8:I{r_row-1})').alignment = align_center
ws3.cell(r_row, 10, f'=SUM(J8:J{r_row-1})').alignment = align_center
ws3.cell(r_row, 11, f'=SUM(K8:K{r_row-1})').alignment = align_center
for c in range(2, 12):
    cell = ws3.cell(r_row, c)
    cell.font = font_subhdr
    cell.fill = fill_subhdr_eco
    cell.border = border_all_black
ws3.row_dimensions[r_row].height = 22

# Table 2: Database Flat Table for Search
r_row += 3
ws3.cell(r_row, 2, 'Tabel 2: Database Seluruh Bal (Gunakan Filter Excel untuk Cari Nomor Bal)').font = font_subhdr
r_row += 1
db_start_r = r_row

db_headers = ['No', 'Nomor Bal', 'Blok', 'Tingkat', 'Posisi Tingkat', 'Saf / Kolom']
for c_idx, h in enumerate(db_headers):
    col_num = 2 + c_idx
    c_cell = ws3.cell(db_start_r, col_num, h)
    c_cell.font = font_block_hdr
    c_cell.fill = fill_banner_eco
    c_cell.alignment = align_center
    c_cell.border = border_all_black
ws3.row_dimensions[db_start_r].height = 22

item_idx = 1
for b_num in range(1, 17):
    b = blocks_data[b_num]
    tier_order = [
        (4, 'Tingkat 4'),
        (3, 'Tingkat 3'),
        (2, 'Tingkat 2'),
        (1, 'Tingkat 1 (Dasar)')
    ]
    for t_row_idx, (t_num, t_desc) in enumerate(tier_order):
        row_vals = b["data"][t_row_idx]
        for c_idx, val in enumerate(row_vals):
            r_row += 1
            saf_name = b["headers"][c_idx]
            
            ws3.cell(r_row, 2, item_idx).alignment = align_center
            ws3.cell(r_row, 3, val).alignment = align_center
            ws3.cell(r_row, 3).font = font_data_jumbo
            ws3.cell(r_row, 4, b["title"]).alignment = align_center
            ws3.cell(r_row, 5, f'T{t_num}').alignment = align_center
            ws3.cell(r_row, 6, t_desc).alignment = align_left
            ws3.cell(r_row, 7, saf_name).alignment = align_center
            
            for c in range(2, 8):
                cell = ws3.cell(r_row, c)
                cell.border = border_all_black
                if item_idx % 2 == 0:
                    cell.fill = fill_tingkat_eco
            ws3.row_dimensions[r_row].height = 20
            item_idx += 1

# Enable AutoFilter on Table 2
ws3.auto_filter.ref = f'B{db_start_r}:G{r_row}'

# Column widths for Sheet 4
ws3.column_dimensions['A'].width = 2.5
ws3.column_dimensions['B'].width = 9
ws3.column_dimensions['C'].width = 15
ws3.column_dimensions['D'].width = 14
ws3.column_dimensions['E'].width = 14
ws3.column_dimensions['F'].width = 20
ws3.column_dimensions['G'].width = 14
ws3.column_dimensions['H'].width = 10
ws3.column_dimensions['I'].width = 10
ws3.column_dimensions['J'].width = 10
ws3.column_dimensions['K'].width = 10

# Save outputs
target_path1 = r'C:\Users\xenov\Downloads\stock gudang susunan\c.xlsx'
target_path2 = r'C:\Users\xenov\Downloads\stock gudang susunan\Stock_Susunan_Bal_30_Agustus_2026.xlsx'

try:
    wb.save(target_path1)
    print(f'Successfully saved to {target_path1}')
except Exception as e:
    print(f'Warning: Could not save to {target_path1}: {e}')

try:
    wb.save(target_path2)
    print(f'Successfully saved to {target_path2}')
except Exception as e:
    print(f'Warning: Could not save to {target_path2}: {e}')

print('Clean headers without utara/selatan completed successfully!')
