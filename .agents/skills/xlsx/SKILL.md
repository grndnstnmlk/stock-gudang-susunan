---
name: xlsx
description: >-
  Comprehensive guide and best practices for creating, reading, modifying, styling,
  and configuring Excel (.xlsx) workbooks using Python and openpyxl. Covers professional
  spreadsheet design, styling tokens, A4 print setup, page break calculations, uniform
  column spacing, and robust file saving workflows.
---

# Excel & Openpyxl Automation Skill (xlsx)

This skill provides standard procedures, code snippets, and design guidelines for building and styling production-ready Excel (`.xlsx`) workbooks.

---

## 1. Core Principles for Spreadsheet Generation

1. **Clean Data Model**: Separate raw data definitions from styling and layout rendering loops.
2. **Consistent Visual Hierarchy**:
   - **Header Banner**: Dark navy fill (`#1B365D`), bold white text (`#FFFFFF`), centered or left-aligned.
   - **Table Headers**: Soft blue fill (`#D9E1F2`), bold dark text (`#1B365D`), centered.
   - **Data Cells**: 10–14pt font, bold for primary numbers, centered or aligned appropriately.
   - **Alternating Rows (Zebra)**: Very soft gray/blue (`#F2F5F9`) for readability.
   - **Borders**: Thin soft borders (`#9BB0C1` or `#B0C4DE`) applied to all data cells.
3. **Uniform Column Widths**:
   - Ensure all data columns in a grid share identical widths so tables render symmetrically.

---

## 2. A4 Print Configuration & Page Break Engineering

When configuring sheets for printing to A4 paper:

```python
import openpyxl
from openpyxl.worksheet.pagebreak import Break

def configure_a4_print(ws, orientation='portrait', fit_w=1, fit_h=0):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = orientation
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True
    ws.print_options.gridLines = True
```

### Preventing Page Cut-Offs:
- **A4 Height in Points**: A4 paper has a total height of 842 points. With 0.35" top/bottom margins, printable height is ~792 points.
- **Budgeting Heights**: Calculate the sum of row heights per page:
  $$\text{Total Height} = \sum (\text{row\_height}) < 790\text{ pt}$$
- **Explicit Page Breaks**: Insert `ws.row_breaks.append(Break(id=row_number))` at exact logical boundaries (e.g. after every 4 blocks) to prevent Excel from inserting accidental mid-table breaks.

---

## 3. Standard Styling Palettes

```python
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Fonts
font_title = Font(name='Segoe UI', size=14, bold=True, color='1B365D')
font_block_hdr = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
font_subhdr = Font(name='Segoe UI', size=10, bold=True, color='1B365D')
font_data_jumbo = Font(name='Segoe UI', size=13, bold=True, color='000000')
font_data = Font(name='Segoe UI', size=11, bold=False, color='000000')

# Fills
fill_navy = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
fill_soft_blue = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_zebra = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')

# Borders
thin_side = Side(border_style='thin', color='9BB0C1')
border_all_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# Alignments
align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
```

---

## 4. Safe File Saving Pattern

Always handle file locking (when the user has the workbook open in Microsoft Excel):

```python
def save_workbook_safely(wb, file_paths):
    for path in file_paths:
        try:
            wb.save(path)
            print(f"Successfully saved to {path}")
        except PermissionError:
            print(f"Warning: Could not save to {path} (File is open in Excel). Close Excel and retry.")
        except Exception as e:
            print(f"Error saving to {path}: {e}")
```

---

## 5. Elderly / Accessibility Formatting Checklist

- [x] Numbers: 13–14pt Bold font.
- [x] Row height: Minimum 22–26pt per row.
- [x] Fixed items per page: Exactly 4 blocks per A4 sheet.
- [x] No distracting page banners eating vertical space.
- [x] High-contrast borders and text.
