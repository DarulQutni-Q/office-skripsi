# PDF SKILL — PDF Handling

Covers extracting text/images from PDF, converting between formats, and verifying document output.

---

## 0. Prerequisites

```bash
# poppler-utils for text/image extraction
brew install poppler

# LibreOffice for DOCX ↔ PDF conversion
brew install --cask libreoffice

# Python PDF toolkits
pip install PyPDF2 pdfminer.six
```

---

## 1. Extract Text from PDF

```bash
# Basic extraction
pdftotext -layout document.pdf output.txt

# Extract specific pages (1-indexed)
pdftotext -layout -f 5 -l 10 document.pdf pages_5-10.txt

# Extract with table layout preservation
pdftotext -layout -table document.pdf table_output.txt

# Extract raw text (no layout preservation)
pdftotext document.pdf raw.txt
```

### Python approach with better control:

```python
import PyPDF2

with open("document.pdf", "rb") as f:
    reader = PyPDF2.PdfReader(f)
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        print(f"=== Page {i+1} ===")
        print(text[:500])
```

---

## 2. Convert PDF to Images

```bash
# Convert specific page to JPEG
pdftoppm -jpeg -r 150 -f 1 -l 1 document.pdf page1

# Convert all pages to PNG
pdftoppm -png -r 100 document.pdf output_prefix

# Convert range to TIFF
pdftoppm -tiff -r 200 -f 3 -l 8 document.pdf pages_3-8
```

---

## 3. Convert Between Formats (via LibreOffice)

```bash
# DOCX → PDF
soffice --headless --convert-to pdf input.docx

# PPTX → PDF
soffice --headless --convert-to pdf input.pptx

# XLSX → PDF
soffice --headless --convert-to pdf input.xlsx

# PDF → DOCX (limited — use with caution)
soffice --headless --convert-to docx input.pdf

# Batch convert all docx in folder
soffice --headless --convert-to pdf *.docx
```

**Note:** LibreOffice conversion is **not perfect**. Complex formatting (tables, images, fonts) may differ from Microsoft Office output. Always verify the converted result.

---

## 4. PDF Metadata

```bash
# Get document info
pdfinfo document.pdf

# Sample output:
# Title:          Skripsi
# Author:         Andy Pratama
# Pages:          126
# File size:      2456789 bytes
# PDF version:    1.7
```

---

## 5. Verify Document Output

Use PDF rendering as the "visual truth" to verify that DOCX edits rendered correctly:

```bash
# 1. Render the edited DOCX to PDF
soffice --headless --convert-to pdf edited.docx

# 2. Extract text
pdftotext -layout edited.pdf edited.txt

# 3. Search for expected content
grep -n "BAB 1" edited.txt
grep -n "Tabel 4.5" edited.txt

# 4. Search for known error markers
grep -n "TODO\|FIXME\|ERROR\|PLACEHOLDER" edited.txt || echo "No error markers found"

# 5. Verify page count
pdfinfo edited.pdf | grep Pages
```

### Automated verification script:

```python
import subprocess
import re

def verify_pdf_content(pdf_path, checks):
    """Run content checks on a PDF file."""
    # Extract text
    txt_path = pdf_path.replace('.pdf', '.txt')
    subprocess.run(['pdftotext', '-layout', pdf_path, txt_path], check=True)

    with open(txt_path, 'r') as f:
        text = f.read()

    results = {}
    for label, pattern in checks.items():
        results[label] = bool(re.search(pattern, text))
        status = "✓" if results[label] else "✗"
        print(f"{status} {label}: {pattern!r}")
    return results

checks = {
    "Has BAB 1": r"BAB\s*1\s+PENDAHULUAN",
    "Has BAB 2": r"BAB\s*2\s+TINJAUAN\s+",
    "Has BAB 3": r"BAB\s*3\s+METODE",
    "Has DAFTAR PUSTAKA": r"DAFTAR\s+PUSTAKA",
    "No error markers": r"TODO|FIXME|ERROR",
}
verify_pdf_content("final.pdf", checks)
```

---

## 6. PDF Comparison

To check if two PDF versions differ:

```python
import PyPDF2

def pdfs_differ(pdf1_path, pdf2_path):
    with open(pdf1_path, 'rb') as f1, open(pdf2_path, 'rb') as f2:
        r1 = PyPDF2.PdfReader(f1)
        r2 = PyPDF2.PdfReader(f2)

        if len(r1.pages) != len(r2.pages):
            return True, f"Page count: {len(r1.pages)} vs {len(r2.pages)}"

        for i in range(len(r1.pages)):
            t1 = r1.pages[i].extract_text().strip()
            t2 = r2.pages[i].extract_text().strip()
            if t1 != t2:
                return True, f"Page {i+1} differs"

        return False, "Identical"
```

---

## 7. Common Issues

| Issue | Cause | Fix |
|-------|-------|-----|
| Text missing after extraction | Image-based PDF (scanned) | Use OCR (tesseract) instead of pdftotext |
| Layout garbled | Complex tables/columns | Try `-table` flag or `pdfminer` with `laparams` |
| Conversion loses formatting | LibreOffice vs MS Office differences | Accept as reference only; verify final in MS Word |
| Fonts missing/wrong | Fonts not embedded in PDF | Embed fonts during DOCX generation |
| Page count changed | Content shift from edits | Update TOC/PAGEREF to match new page count |
