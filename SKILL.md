# OFFICE SKILL — AI Agent Document Manipulation Skill

Teaches AI coding agents how to manipulate Office documents (DOCX, PPTX, PDF, XLSX) programmatically using python-docx, python-pptx, openpyxl, LibreOffice, and raw XML editing.

---

## 0. Prerequisites

| Tool | Install | Purpose |
|------|---------|---------|
| `python3` + `python-docx` | `pip install python-docx` | Read/write DOCX |
| `python3` + `python-pptx` | `pip install python-pptx` | Read/write PPTX |
| `python3` + `openpyxl` | `pip install openpyxl` | Read/write XLSX |
| `LibreOffice` | `brew install --cask libreoffice` | DOCX ↔ PDF, PPTX → PDF |
| `poppler-utils` | `brew install poppler` | PDF → text/images |
| `zip`/`unzip` | Built-in on macOS/Linux | Extract/repack Office files |
| `lxml` | `pip install lxml` | XML validation |

---

## 1. Core Concept

| Format | Internals | Edit Approach |
|--------|-----------|---------------|
| DOCX | ZIP of XML files (`word/document.xml`) | python-docx API OR unzip → edit XML → rezip |
| PPTX | ZIP of XML files (`ppt/slides/slideN.xml`) | python-pptx API OR unzip → edit XML → rezip |
| PDF | Binary format (not editable) | Extract text, convert via LibreOffice, overlay |
| XLSX | ZIP of XML files (`xl/worksheets/sheetN.xml`) | openpyxl API OR unzip → edit XML → rezip |

---

## 2. Workflow

```
1. Backup      → cp original.docx original_backup.docx
2. Understand  → python-docx to inspect structure + soffice to render PDF
3. Analyze     → Identify ALL issues (headings, captions, references, numbering)
4. Edit        → python-docx API OR unzip → merge runs → str_replace XML → rezip
5. Validate    → ZIP integrity → Document opens → XML well-formed → content check
6. Verify      → Render PDF → pdftotext → grep for correctness
7. Repeat 4-6  → Until all issues resolved
```

---

## 3. Sub-skills

| Skill | File | For |
|-------|------|-----|
| DOCX | `skills/docx/SKILL.md` | Word documents — edit text, headings, TOC, captions, tables, styles |
| PPTX | `skills/pptx/SKILL.md` | PowerPoint presentations — slides, layouts, text replacement |
| PDF | `skills/pdf/SKILL.md` | PDF extraction, conversion, verification |
| Spreadsheet | `skills/spreadsheet/SKILL.md` | Excel spreadsheets — read/write, formatting |

---

## 4. Universal Rules

1. **DOCX/PPTX/XLSX are ZIP files** — always extract first if editing XML directly. Never edit the binary.
2. **Backup before every edit** — `cp original.docx original_backup.docx`
3. **One edit batch at a time** — edit → validate → edit → validate. Never chain 10 edits without checking.
4. **Render to PDF to verify** — what you see in XML != what renders. Always do `soffice --headless --convert-to pdf out.docx` then `pdftotext -layout out.pdf out.txt` and grep.
5. **Character encoding** — In Office XML, `&` is `&amp;`, `<` is `&lt;`, `>` is `&gt;`. Always use XML-escaped strings in str_replace.
6. **Merge runs first** — Word fragments text across multiple `<w:r>` elements during editing. Merge them before search/replace to avoid missed matches.
7. **Field codes (TOC, PAGEREF, SEQ) don't auto-update** — LibreOffice headless is unreliable for recalculation. Hardcode correct values in the cached field result, then set `updateFields=true` in settings.xml as a safety net.
8. **Cross-reference effects** — Inserting content shifts page numbers, caption numbers, and all text references. Always update SEQ fields AND plain-text references ("see Table 5") after insertion.

---

## 5. Validation (Run After Every Edit)

```bash
python3 scripts/validate-docx.py output.docx
```
Or check manually:
```python
from docx import Document
import zipfile
from lxml import etree

with zipfile.ZipFile('output.docx', 'r') as z:
    assert z.testzip() is None, 'ZIP corrupt'
doc = Document('output.docx')
assert len(doc.paragraphs) > 0, 'No paragraphs'
for name in z.namelist():
    if name.endswith('.xml'):
        etree.fromstring(z.read(name))
        print(f'  ✓ {name}')
print('All checks passed')
```

---

## 6. Anti-Patterns (Don't Do These)

- ❌ Edit `.docx`/`.pptx` binary directly — use unzip first
- ❌ Replace short strings without surrounding XML context (will hit wrong location)
- ❌ Skip backup before editing
- ❌ Claim edit worked without PDF verification
- ❌ Rely on LibreOffice to auto-update TOC/SEQ fields correctly
- ❌ Edit text in XML without merging runs first
- ❌ Hardcode page numbers that will shift after later edits
- ❌ Ignore XML errors — one unclosed tag corrupts the entire document
