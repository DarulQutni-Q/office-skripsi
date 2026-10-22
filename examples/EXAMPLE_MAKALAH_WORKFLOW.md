# Example: Makalah DOCX Edit Workflow

This shows how to apply the DOCX skill to a course paper (makalah) editing session.
Makalah is a first-class document type here: same techniques as skripsi, but the
expected structure is PENDAHULUAN, PEMBAHASAN, PENUTUP, DAFTAR PUSTAKA.

## Setup

```bash
mkdir -p work && cd work
cp ../Makalah_Revisi.docx ./original_backup.docx
unzip -q ../Makalah_Revisi.docx -d unpacked/
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

Makalah headings are often typed by hand (`1.`, `1.1`, `2. ...`) — hardcoded, so
deleting one does not renumber the rest. Convert them to Word native multilevel
numbering (works only after merge-runs):

```bash
# Read-only audit: exits 1 if any hardcoded prefixes remain
python3 ../scripts/numbering.py unpacked/ analyze

# Fix: attach w:numPr + strip the literal "1.2" prefix (idempotent)
python3 ../scripts/numbering.py unpacked/ convert
```

Verify by rendering (numbers are not a cached field, so LibreOffice is authoritative):

```bash
# repack first (Step 5), then:
soffice --headless --convert-to pdf revised.docx
pdftotext -layout revised.pdf revised.txt
```

## Step 3.6: Italic for Foreign Terms (recommended for makalah)

Makalah pembahasan sections are full of foreign terms. Per PUEBI they must be
italic — but whole sentences often end up italic by accident:

```bash
# Read-only audit first
python3 ../scripts/italic-foreign.py audit original_backup.docx --strip-other

# Apply: writes makalah.italic.docx, original untouched
python3 ../scripts/italic-foreign.py apply original_backup.docx --strip-other
```

## Step 4: Make Edits

Edit the XML file (`unpacked/word/document.xml`) with targeted str_replace. Always include surrounding context to ensure uniqueness.

## Step 5: Repack and Validate (as makalah!)

```bash
cd unpacked && zip -Xr ../revised.docx . && cd ..
python3 ../scripts/validate-docx.py revised.docx --type makalah
```

The `--type makalah` flag checks for PENDAHULUAN, PEMBAHASAN, PENUTUP,
DAFTAR PUSTAKA instead of the skripsi chapter list. Without it, validation
would wrongly FAIL on a perfectly good makalah.

## Step 6: Render and Verify

```bash
soffice --headless --convert-to pdf revised.docx
pdftotext -layout revised.pdf revised.txt

# Compare with original
diff original_backup.txt revised.txt | head -50
```
