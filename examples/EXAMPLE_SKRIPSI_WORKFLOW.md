# Example: Skripsi DOCX Edit Workflow

This shows how to apply the DOCX skill to a real thesis editing session.

## Setup

```bash
mkdir -p work && cd work
cp ../Skripsi_AndyPratama_Revisi.docx ./original_backup.docx
unzip -q ../Skripsi_AndyPratama_Revisi.docx -d unpacked/
```

## Step 1: Understand the Document

```bash
python3 ../scripts/docx-tools.py original_backup.docx all
```

This lists all headings, captions, tables, citations, and SEQ fields — a complete picture before editing.

## Step 2: Render Original to PDF

```bash
soffice --headless --convert-to pdf original_backup.docx
pdftotext -layout original_backup.pdf original_backup.txt
```

Now you have the original text to compare against after edits.

## Step 3: Merge Runs

```bash
python3 ../scripts/merge-runs.py unpacked/word/document.xml
```

## Step 3.5: Automatic Heading Numbering (optional)

If headings are typed by hand (`1.`, `1.1`, `2.1.2 ...`) they are hardcoded — deleting one doesn't renumber the rest. Convert them to Word's native multilevel numbering (works only after merge-runs):

```bash
# Read-only audit: exits 1 if any hardcoded prefixes remain
python3 ../scripts/numbering.py unpacked/ analyze

# Fix: attach w:numPr + strip the literal "1.2.3" prefix (idempotent)
python3 ../scripts/numbering.py unpacked/ convert
```

Run this BEFORE making other text edits, then verify by rendering (numbers aren't a cached field, so LibreOffice is authoritative):

```bash
soffice --headless --convert-to pdf unpacked-repack.pdf  # or repack first (Step 5)
pdftotext -layout unpacked-repack.pdf unpacked-repack.txt
```

Delete a section and re-render — surviving headings must renumber themselves.

## Step 4: Make Edits

Edit the XML file (`unpacked/word/document.xml`) with targeted str_replace. Always include surrounding context to ensure uniqueness.

```python
# Example: fix a heading
old = 'Heading 1"><w:t>1. Sistem Informasi Pembayaran Sekolah</w:t>'
new = 'Heading 1"><w:t>1. Sistem Informasi Pembayaran Berbasis Web</w:t>'
content = content.replace(old, new)
```

## Step 5: Repack and Validate

```bash
cd unpacked && zip -Xr ../revised.docx . && cd ..
python3 ../scripts/validate-docx.py revised.docx
```

## Step 6: Render and Verify

```bash
soffice --headless --convert-to pdf revised.docx
pdftotext -layout revised.pdf revised.txt

# Compare with original
diff original_backup.txt revised.txt | head -50

# Or search for specific changes
grep -n "Sistem Informasi Pembayaran Berbasis Web" revised.txt
```

## Step 7: Update SEQ Fields

If you added or removed tables/figures, renumber SEQ fields:

```python
# In your edit script
def renumber_seq(content, seq_type="Tabel"):
    import re
    counter = 0
    pattern = r'(<w:fldSimple[^>]*w:instr="\s*SEQ\s+' + seq_type + r'[^"]*"[^>]*>.*?)(<w:t[^>]*>)(\d+)(</w:t>)(.*?</w:fldSimple>)'
    def replacement(match):
        nonlocal counter
        counter += 1
        return match.group(1) + match.group(2) + str(counter) + match.group(4) + match.group(5)
    content = re.sub(pattern, replacement, content, flags=re.DOTALL)
    return content, counter

content, _ = renumber_seq(content, "Tabel")
content, _ = renumber_seq(content, "Gambar")
```

Then repack, revalidate, and reverify.
