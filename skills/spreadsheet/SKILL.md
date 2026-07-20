# SPREADSHEET SKILL — XLSX Manipulation

Covers reading, writing, editing, and formatting spreadsheet files using `openpyxl`.

---

## 0. Prerequisites

```bash
pip install openpyxl lxml
```

---

## 1. Reading XLSX Files

```python
from openpyxl import load_workbook

wb = load_workbook("data.xlsx")

# List sheet names
print(f"Sheets: {wb.sheetnames}")

# Read a specific sheet
ws = wb["Sheet1"]
print(f"Dimensions: {ws.dimensions}")
print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")

# Read cell values
for row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
    print(row)

# Read by cell reference
for row in range(1, ws.max_row + 1):
    cell_value = ws.cell(row=row, column=1).value
    print(f"Row {row}, Col A: {cell_value}")
```

---

## 2. Writing XLSX Files

```python
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

wb = Workbook()
ws = wb.active
ws.title = "Data"

# Write headers
headers = ["No", "Name", "Score", "Grade"]
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal="center")
    cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    cell.font = Font(bold=True, color="FFFFFF")

# Write data
data = [
    [1, "Alice", 85, "A"],
    [2, "Bob", 72, "B"],
    [3, "Charlie", 90, "A"],
]
for row_idx, row_data in enumerate(data, 2):
    for col_idx, value in enumerate(row_data, 1):
        ws.cell(row=row_idx, column=col_idx, value=value)

# Auto-adjust column widths
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[column].width = max_length + 2

wb.save("output.xlsx")
```

---

## 3. Editing Existing Files

```python
from openpyxl import load_workbook

wb = load_workbook("existing.xlsx")
ws = wb.active

# Find and replace
for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
    for cell in row:
        if cell.value and "old text" in str(cell.value):
            cell.value = str(cell.value).replace("old text", "new text")

# Add a new row at the end
new_row = ws.max_row + 1
ws.cell(row=new_row, column=1, value="New data")
ws.cell(row=new_row, column=2, value=123)

# Insert a column
ws.insert_cols(2)  # Insert column before column 2

# Delete a row
ws.delete_rows(5)  # Delete row 5

wb.save("edited.xlsx")
```

---

## 4. Formatting

```python
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)

# Cell font
cell.font = Font(name="Calibri", size=11, bold=True, italic=True, color="FF0000")

# Cell fill
cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# Alignment
cell.alignment = Alignment(
    horizontal="center",
    vertical="center",
    wrap_text=True,
    text_rotation=0
)

# Borders
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)
cell.border = thin_border

# Number format
cell.number_format = "#,##0.00"
cell.number_format = "dd/mm/yyyy"
cell.number_format = "0.00%"

# Merge cells
ws.merge_cells("A1:C1")
ws["A1"] = "Merged Header"
ws["A1"].alignment = Alignment(horizontal="center")
```

---

## 5. Working with Formulas

```python
from openpyxl.utils import get_column_letter

# Write formulas
ws["C1"] = "Total"
ws["C2"] = "=A2+B2"         # Simple addition
ws["C3"] = "=SUM(A:A)"      # Sum entire column
ws["C4"] = "=AVERAGE(B:B)"  # Average column
ws["C5"] = "=IF(A2>80,\"Pass\",\"Fail\")"

# Formula with dynamic cell references
last_row = 10
ws.cell(row=last_row, column=4).value = f"=SUM(D1:D{last_row-1})"
```

---

## 6. Charts

```python
from openpyxl.chart import BarChart, Reference

chart = BarChart()
chart.title = "Score Distribution"
chart.x_axis.title = "Students"
chart.y_axis.title = "Scores"

data = Reference(ws, min_col=3, min_row=1, max_row=10)
categories = Reference(ws, min_col=2, min_row=2, max_row=10)

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

ws.add_chart(chart, "E2")
```

---

## 7. Data Validation & Conditional Formatting

```python
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# Data validation — dropdown list
dv = DataValidation(
    type="list",
    formula1='"Option A,Option B,Option C"',
    allow_blank=True
)
dv.error = "Please select from the list"
dv.errorTitle = "Invalid Input"
ws.add_data_validation(dv)
dv.add("A2:A100")

# Conditional formatting — highlight cells
ws.conditional_formatting.add(
    "B2:B100",
    CellIsRule(operator="greaterThan", formula=["80"], fill=PatternFill(bgColor="92D050"))
)
```

---

## 8. Validation

```python
from openpyxl import load_workbook
import zipfile
from lxml import etree

def validate_xlsx(filepath):
    print(f"=== Validating {filepath} ===")
    with zipfile.ZipFile(filepath, 'r') as z:
        bad = z.testzip()
        if bad:
            print(f"  ✗ Corrupted: {bad}")
            return False
        print(f"  ✓ ZIP valid ({len(z.namelist())} files)")

    wb = load_workbook(filepath)
    print(f"  ✓ Opens: {wb.sheetnames}")

    with zipfile.ZipFile(filepath, 'r') as z:
        for name in z.namelist():
            if name.endswith('.xml'):
                try:
                    etree.fromstring(z.read(name))
                except:
                    print(f"  ✗ Invalid XML: {name}")
                    return False
        print("  ✓ All XML valid")
    return True

validate_xlsx("output.xlsx")
```

---

## 9. Raw XML Editing

For advanced changes (pivot tables, macros, custom XML parts):

```bash
mkdir -p work && cd work
cp ../original.xlsx ../original_backup.xlsx
unzip -q ../original.xlsx -d unpacked/
# Edit xl/worksheets/sheet1.xml etc.
cd unpacked && zip -Xr ../output.xlsx . && cd ..
```

Key XML files:

| Path | Content |
|------|---------|
| `xl/workbook.xml` | Workbook structure, sheet names |
| `xl/worksheets/sheetN.xml` | Individual sheet data |
| `xl/sharedStrings.xml` | Shared string table |
| `xl/styles.xml` | Cell formatting styles |
| `xl/theme/themeN.xml` | Theme colors/fonts |
| `xl/calcChain.xml` | Formula calculation chain |

---

## 10. Common Issues

| Issue | Cause | Fix |
|-------|-------|-----|
| Formulas not calculating | Open in Excel to trigger calculation | Not an error — Excel recalculates on open |
| Formatting lost | openpyxl may not preserve all XML styles | Use raw XML edit for complex formatting |
| Shared strings broken | String table de-synchronized | Use `data_only=False` and let openpyxl manage strings |
| Images disappear | Missing relationship entries in .rels | Add image parts and relationships manually |
| Large files slow | Iterate with `read_only=True` | Use `load_workbook(filename, read_only=True)` |
