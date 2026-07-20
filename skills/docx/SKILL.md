# DOCX SKILL — Word Document Manipulation

Covers editing, generating, and validating DOCX files using `python-docx` and raw XML editing.

---

## 0. Quick Reference

```bash
# Unzip to edit XML directly
mkdir -p work && cd work
cp ../original.docx ../original_backup.docx
unzip -q ../original.docx -d unpacked/

# After editing, repack
cd unpacked && zip -Xr ../output.docx . && cd ..

# Render to PDF to verify
soffice --headless --convert-to pdf output.docx
pdftotext -layout output.pdf output.txt
```

---

## 1. Inspecting a DOCX

```python
from docx import Document

doc = Document("skripsi.docx")

# List all headings
for i, p in enumerate(doc.paragraphs):
    if p.style and "Heading" in p.style.name:
        print(f"[{i}] {p.style.name}: {p.text.strip()[:80]}")

# List all captions
for i, p in enumerate(doc.paragraphs):
    if p.style and "Caption" in p.style.name:
        has_seq_tabel = 'SEQ Tabel' in p._element.xml
        has_seq_gambar = 'SEQ Gambar' in p._element.xml
        seq_type = "Tabel" if has_seq_tabel else ("Gambar" if has_seq_gambar else "NONE")
        print(f"[{i}] [{seq_type}] {p.text.strip()[:60]}")

# List all tables
for i, t in enumerate(doc.tables):
    print(f"Table {i}: {len(t.rows)} rows x {len(t.columns)} cols")
    for j, row in enumerate(t.rows[:3]):
        cells = [c.text.strip()[:30] for c in row.cells]
        print(f"  Row {j}: {cells}")

# Count images
image_count = sum(1 for r in doc.part.rels.values() if 'image' in r.reltype)
print(f"Images: {image_count}")
```

---

## 2. Working with XML Directly

### 2.1 Merge Runs First

Word fragments text across `<w:r>` elements. Always merge before search/replace:

```python
import re

def merge_runs_in_xml(xml_content):
    """Merge consecutive <w:r> elements that share identical <w:rPr>."""
    def merge_paragraph(match):
        para = match.group(0)
        # Find all run groups within this paragraph
        runs = list(re.finditer(
            r'<w:r\b[^>]*>'
            r'(?:<w:rPr>.*?</w:rPr>)?'
            r'<w:t[^>]*>.*?</w:t>'
            r'</w:r>',
            para, re.DOTALL
        ))
        if len(runs) < 2:
            return para

        # Check if consecutive runs can be merged (same rPr)
        merged = []
        i = 0
        while i < len(runs):
            current = runs[i]
            current_text = current.group(0)

            rpr_match = re.search(r'<w:rPr>.*?</w:rPr>', current_text, re.DOTALL)
            current_rpr = rpr_match.group(0) if rpr_match else ''

            # Try to absorb next runs
            j = i + 1
            while j < len(runs):
                next_run = runs[j]
                next_text = next_run.group(0)
                next_rpr_match = re.search(r'<w:rPr>.*?</w:rPr>', next_text, re.DOTALL)
                next_rpr = next_rpr_match.group(0) if next_rpr_match else ''

                if current_rpr == next_rpr:
                    j += 1
                else:
                    break

            if j > i + 1:
                combined_text = ''
                for k in range(i, j):
                    t_match = re.search(r'<w:t[^>]*>(.*?)</w:t>', runs[k].group(0), re.DOTALL)
                    if t_match:
                        combined_text += t_match.group(1)
                escaped = combined_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                tag_match = re.search(r'<w:t[^>]*>', current_text)
                t_tag_open = tag_match.group(0) if tag_match else '<w:t xml:space="preserve">'
                merged_run = f'<w:r>{current_rpr}{t_tag_open}{escaped}</w:t></w:r>'
                merged.append(merged_run)
            else:
                merged.append(current_text)
            i = j

        result = ''
        pos = 0
        for m in merged:
            idx = para.find(m, pos)
            if idx > pos:
                result += para[pos:idx]
            result += m
            pos += len(m)
        result += para[pos:]
        return result

    xml_content = re.sub(
        r'<w:p\b[^>]*>.*?</w:p>',
        merge_paragraph,
        xml_content,
        flags=re.DOTALL
    )
    return xml_content

with open("unpacked/word/document.xml", encoding="utf-8") as f:
    content = f.read()

content = merge_runs_in_xml(content)

with open("unpacked/word/document.xml", "w", encoding="utf-8") as f:
    f.write(content)
```

### 2.2 Search with Context

```python
with open("unpacked/word/document.xml", encoding="utf-8") as f:
    content = f.read()

target = "TEKS YANG DICARI"
idx = content.find(target)
if idx >= 0:
    print(content[idx-200:idx+200])
```

### 2.3 Find by Bookmark (More Accurate for Headings)

```python
idx = content.find('w:name="_Toc235026593"')
if idx >= 0:
    p_start = content.rfind("<w:p ", 0, idx)
    p_end = content.find("</w:p>", idx) + 6
    para_xml = content[p_start:p_end]
    print(para_xml)
```

### 2.4 Safe str_replace in XML

```python
old = '<w:t>Old Heading Text</w:t>'
new = '<w:t>New Heading Text</w:t>'
content = content.replace(old, new)
```

Include enough surrounding XML context (30-60 chars) to make the match unique.

---

## 3. Field Codes (TOC, SEQ, PAGEREF)

### 3.1 Structure

```xml
<w:fldSimple w:instr=" SEQ Tabel \* ARABIC ">
  <w:bookmarkStart w:id="5" w:name="_Toc235026567"/>
  <w:r><w:t>1</w:t></w:r>
</w:fldSimple>
```

The `<w:t>1</w:t>` is the **cached value**. Word uses this until fields are refreshed (Ctrl+A → F9).

### 3.2 Updating SEQ Numbers

```python
import re

with open("unpacked/word/document.xml", encoding="utf-8") as f:
    content = f.read()

def renumber_seq(content, seq_type="Tabel"):
    """Renumber all SEQ fields of a given type to be sequential."""
    pattern = r'(<w:fldSimple[^>]*w:instr="\s*SEQ\s+' + seq_type + r'[^"]*"[^>]*>.*?)(<w:t[^>]*>)(\d+)(</w:t>)(.*?</w:fldSimple>)'

    def replacement(match):
        nonlocal counter
        counter += 1
        return match.group(1) + match.group(2) + str(counter) + match.group(4) + match.group(5)

    counter = 0
    content = re.sub(pattern, replacement, content, flags=re.DOTALL)
    return content, counter

content, tabel_count = renumber_seq(content, "Tabel")
content, gambar_count = renumber_seq(content, "Gambar")
print(f"Renumbered {tabel_count} tables, {gambar_count} figures")
```

### 3.3 Updating TOC Entries (PAGEREF)

After final render, extract actual page numbers from PDF:

```bash
pdftotext -layout final.pdf final.txt
grep -n "BAB 1\|BAB 2\|BAB 3" final.txt
```

Then update PAGEREF cached values:
```python
# In document.xml, find PAGEREF fields and update cached page numbers
idx = 0
while True:
    idx = content.find('PAGEREF', idx)
    if idx < 0:
        break
    # ... update the <w:t> value after the separator
    idx += 1
```

### 3.4 Chain Effects of Inserting Content (Most Commonly Overlooked)

When you insert a heading, table, or figure in the middle of a document:

1. **Page numbers shift** for ALL content after the insertion point — TOC entries, Daftar Tabel, Daftar Gambar all need updating.
2. **SEQ numbers shift** — inserting Table 2.1 between existing Table 2.1 and 2.2 means old 2.2 becomes 2.3, old 2.3 becomes 2.4, etc. Update ALL SEQ cached values.
3. **Plain-text references break** — "...as shown in Table 2.1" now points to the wrong table. These are plain text, not fields — you must find and fix each one manually.
4. **Bookmark gaps** — headings added manually after initial TOC generation may lack `_Toc` bookmarks, making them invisible in the TOC. Always verify every heading has a matching bookmark.

### 3.5 Pre-Final Checklist

- [ ] Every heading (Heading1/2/3) has a unique `_Toc...` bookmark in the body
- [ ] Every bookmark has a matching hyperlink+PAGEREF entry in the TOC
- [ ] Every table/figure uses `SEQ Tabel`/`SEQ Gambar` field for numbering (not hardcoded numbers)
- [ ] All SEQ field values are sequential — no duplicate or skipped numbers
- [ ] All plain-text references ("see Table 5") match the actual SEQ numbers
- [ ] After final render, grep all TOC/Daftar Tabel/Daftar Gambar page numbers and compare with actual page locations

---

## 4. Thesis Substance Checklist (Content Consistency)

Beyond formatting, verify the actual academic content:

### 4.1 Theory vs Implementation Cross-Check

```python
import re

def check_theory_vs_implementation(filepath):
    doc = Document(filepath)
    all_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

    # Split by chapters (assumes "BAB 1", "BAB 2" etc. as separators)
    chapters = re.split(r'BAB\s+\d+', all_text)
    if len(chapters) >= 4:
        bab2 = chapters[1]  # Landasan Teori
        bab3 = chapters[2]  # Metode
        bab4 = chapters[3]  # Hasil/Pembahasan

        # Extract technical terms used in Bab 3/4
        tech_terms = set(re.findall(r'(flowchart|uml|er[dds]|use.?case|sequence.?diagram|activity.?diagram|class.?diagram|black.?box|white.?box)', bab3 + bab4, re.I))

        # Check each term has theoretical basis in Bab 2
        for term in tech_terms:
            if term.lower() in bab2.lower():
                print(f"  ✓ {term} found in Bab 2")
            else:
                print(f"  ✗ {term} used in Bab 3/4 but MISSING from Bab 2!")

    return tech_terms
```

### 4.2 Number & Enum Consistency Across Chapters

- [ ] Count of steps/phases/stages (e.g. "five stages") matches the actual list count
- [ ] Enum values/state names mentioned in narrative match source code (if available)
- [ ] States/enums mentioned in 3 places (Bab 4 narrative, UI screenshots, source code) are all consistent
- [ ] **Source code is ground truth** — if code says 9 states but text says 8, the text is wrong

### 4.3 Citation & Bibliography Cross-Check

```python
def check_citations(bib_filepath, doc_filepath):
    """Check in-text citations have matching bibliography entries."""
    doc = Document(doc_filepath)
    all_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

    # Find all in-text citations: (Author, Year) or (Author et al., Year)
    citations = re.findall(r'\(([^)]+?\d{4}[^)]*?)\)', all_text)
    print(f"In-text citations found: {len(citations)}")

    # Extract bibliography entries (simplified — assumes Daftar Pustaka section)
    in_bib = False
    bib_entries = []
    for p in doc.paragraphs:
        if 'DAFTAR PUSTAKA' in p.text.upper():
            in_bib = True
            continue
        if in_bib and p.text.strip() and 'BAB ' in p.text.upper():
            in_bib = False
        if in_bib and p.text.strip():
            bib_entries.append(p.text.strip())

    # Quick check: first author surname in each citation should appear in bibliography
    for cite in citations:
        first_author = cite.split(',')[0].split()[-1].strip('(')
        found = any(first_author.lower() in entry.lower() for entry in bib_entries)
        if not found:
            print(f"  ✗ '{cite}' — author not found in Daftar Pustaka!")
        else:
            print(f"  ✓ {cite}")

    # Reverse check: entries never cited
    for entry in bib_entries:
        first_word = entry.split()[0].strip('.').strip(',')
        if first_word.lower() not in all_text.lower():
            print(f"  ? '{first_word}' in Daftar Pustaka but never cited in text")
```

- [ ] Every in-text citation (Author, Year) has a matching entry in Daftar Pustaka
- [ ] Author name format matches exactly (if bib uses "Yudha Z.", citation must too)
- [ ] Adding new cited content? Add the bibliography entry too, in alphabetical order
- [ ] Reverse check: bibliography entries that are never cited in text (informational only)

### 4.4 Biography / Riwayat Hidup

- [ ] No duplicate education entries (same school name twice with consecutive years)
- [ ] **Do not guess** — if ambiguous (e.g. two SD entries), ask the author to confirm
- [ ] Verify timeline continuity (no gaps or overlapping years)

### 4.5 Source Code Verification (When Available)

If the author provides the actual source code:

```bash
# Example: check ConversationState enum against narrative
grep -n "ConversationState::" src/*.php | grep -o "ConversationState::[A-Z_]*" | sort -u

# Compare with states listed in thesis narrative
# Code is ground truth — update text to match code
```

- [ ] Enums/states/constants mentioned in text exist in the actual code
- [ ] When text and code disagree, code wins (it's what actually runs)
- [ ] UI screenshots that can't be edited but show wrong content → flag to author for re-screenshot

---

## 5. Writing Original Thesis Content (Anti-Plagiarism Guide)

For when you need to WRITE new content (not just edit formatting) — target similarity <20%.

### 5.1 Core Principles

1. **Never copy verbatim** from any source (books, articles, Wikipedia), even for "commonly known" definitions. Similarity checkers detect identical sentences regardless of source popularity.
2. **Paraphrase = total restructuring, not synonym substitution.** Changing 2-3 words while keeping the same sentence structure is still detectable as plagiarism ("patchwriting"). You must change:
   - Clause order (subject-predicate-object rearranged)
   - Sentence length (split long sentences, combine short ones)
   - Perspective (source says "what X is", you say "why X is used in this research")
3. **Always cite the source** (Author, Year) — correct citation reduces plagiarism risk because it's acknowledged as a derived paraphrase.
4. **Contextualize to this research** at the end of every theoretical paragraph. This automatically makes the text 0% similar to any source because it describes specifics of the author's own system.
5. **Vary sources** — don't base one sub-chapter on a single source. Synthesize 2-3 sources per concept.

### 5.2 Safe Paragraph Template for Landasan Teori

```
Sentence 1: Define the concept, paraphrased from source + citation.
Sentence 2: Elaborate components (synthesize from 2+ sources + citations).
Sentence 3: Explicit contextualization to THIS research — mention actual
            features/variables/entities from the author's system.
            (This sentence is 100% unique — no similarity checker can flag it.)
```

### 5.3 Example: Good vs Bad Paraphrase

**BAD (patchwriting — will be detected):**
> "Flowchart adalah representasi grafis yang menggambarkan urutan proses atau langkah-langkah dalam menjalankan suatu prosedur atau sistem secara sistematis."
> (Nearly identical to Wikipedia's Indonesian definition — highly detectable.)

**GOOD (restructured + contextualized):**
> "Dalam pengembangan sistem, alur kerja perlu divisualisasikan agar mudah dipahami oleh pengembang maupun pemangku kepentingan. Salah satu teknik visualisasi yang banyak dipakai adalah flowchart, yaitu diagram yang menunjukkan urutan langkah proses dari awal hingga akhir menggunakan simbol-simbol baku (Pressman & Maxim, 2020). Pada penelitian ini, flowchart digunakan untuk memetakan alur pembuatan tagihan otomatis, proses pembayaran melalui Midtrans, hingga pengiriman notifikasi WhatsApp, sehingga setiap tahapan sistem dapat diverifikasi secara visual sebelum implementasi dilakukan."

### 5.4 MUST Avoid

- Copying Indonesian Wikipedia definitions (most commonly detected — many theses cite the same source verbatim)
- Verbatim text from other theses/journals on similar topics (cross-database checkers catch this)
- Literal 1:1 translation from English sources (sentence structure remains detectable)
- Relying on a single source per sub-chapter

### 5.5 High-Risk Paragraphs

Based on experience, these are most frequently flagged:
1. General definitions (Flowchart, UML, ERD, Waterfall, Black Box Testing)
2. Technology descriptions (Laravel, MySQL, Payment Gateway)
3. Research methodology (Black Box Testing, UAT, Likert Scale)

**Solution:** Always append a contextualization sentence specific to this research at the end of every definition paragraph.

### 5.6 After Writing — Verify

If possible, run through a similarity checker (Turnitin, Quetext, Grammarly) before final submission. Target: <20-25% (confirm with university policy). If highlighted, the structure is still too close to the source — re-paraphrase, don't just swap 1-2 words.

---

## 6. Editing with python-docx (High-Level API)

```python
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document("input.docx")

# Find and replace text in paragraphs
for p in doc.paragraphs:
    if "old text" in p.text:
        for run in p.runs:
            if "old text" in run.text:
                run.text = run.text.replace("old text", "new text")

# Add a heading
doc.add_heading("New Section", level=2)

# Add a paragraph with formatting
p = doc.add_paragraph()
run = p.add_run("This is bold and italic text")
run.bold = True
run.italic = True
run.font.size = Pt(12)
run.font.name = "Times New Roman"

# Add a table
table = doc.add_table(rows=3, cols=4)
table.style = "Table Grid"
for i, row in enumerate(table.rows):
    for j, cell in enumerate(row.cells):
        cell.text = f"Cell {i},{j}"

doc.save("output.docx")
```

**Limitations of python-docx:**
- Cannot update fields (TOC, SEQ, PAGEREF) — must edit XML directly
- May lose some formatting on save (especially complex styles)
- Cannot merge runs internally — use XML approach for fine-grained control

---

## 7. Common DOCX Issues & Fixes

| Issue | How to Detect | How to Fix |
|-------|---------------|------------|
| Missing heading numbers | Check `numPr` in heading XML or missing prefix | Add `w:numPr` with correct `w:ilvl` in paragraph properties |
| Duplicate caption numbers | SEQ values not in [1,2,3,N] sequence | Renumber SEQ fields (see §3.2) |
| TOC shows wrong pages | PAGEREF cached values stale | Render PDF → extract pages → hardcode PAGEREF values |
| Caption references wrong | Text says "Table 5" but SEQ shows Table 4 | Fix the text reference manually |
| Text split across many runs | Search text fails in XML but visible in Word | Merge runs first (§2.1) |
| Corrupt document after edit | Validation fails | Check XML for unclosed tags, unescaped chars |
| Heading missing from TOC | No `_Toc` bookmark on that heading in body | Add bookmark + TOC entry manually |

---

## 8. Final Validation Checklist

- [ ] ZIP integrity passes (not corrupted)
- [ ] Document opens in python-docx without error
- [ ] All XML files well-formed
- [ ] All required sections present (PENDAHULUAN, TINJAUAN PUSTAKA, etc.)
- [ ] Headings numbered sequentially
- [ ] All `_Toc` bookmarks have matching TOC entries
- [ ] SEQ Tabel/Gambar numbers sequential (no duplicates or gaps)
- [ ] Plain-text references to table/figure numbers match SEQ values
- [ ] TOC, Daftar Tabel, Daftar Gambar page numbers match actual pages
- [ ] PDF renders without missing/broken content
- [ ] Search for known error keywords returns empty

---

## 9. Lessons Learned (Real Experience)

### 9.1 Common Problems Found in Real Theses

1. **Headings without numbers** — Sub-chapters in Landasan Teori added later but never numbered.
2. **SEQ fields out of sync** — Caption numbers not sequential because insertions were never renumbered.
3. **Duplicate biography entries** — Same school listed twice (e.g. "SD Negeri X" and "SD 001 X") due to copy-paste errors.
4. **Stale TOC** — Daftar Isi/Tabel/Gambar still shows old page numbers because fields were never updated.
5. **Plagiarism-flagged definitions** — Generic definitions (Flowchart, UML, ERD) too close to Wikipedia/other sources.

### 9.2 Solutions

1. **Headings**: Search for all `w:pStyle w:val="Heading3"` and check for `numPr` or hardcoded number prefix. Add `w:numPr` if missing.
2. **SEQ fields**: List all `SEQ Tabel`/`SEQ Gambar` and verify sequential ordering. Renumber with script (§3.2).
3. **Biography**: Ask the author — never delete without confirmation. Could be a valid timeline or a typo.
4. **TOC**: After all edits, ask user to open in Word and press Ctrl+A → F9 to refresh all fields. This is more reliable than LibreOffice recalculation.
5. **Plagiarism**: Always append a contextualization sentence specific to this research at the end of every definition paragraph. This single sentence makes the entire paragraph unique.

### 9.3 Hard Rules (Never Break These)

- ❌ Never edit `.docx` binary directly — always unzip first
- ❌ Never replace numbers without checking chain effects on all downstream content
- ❌ Never delete content without author confirmation
- ❌ Never claim "fixed" without rendering to PDF and verifying
- ❌ Never batch 10+ edits before validating — one bad XML tag corrupts the entire document
- ❌ Never ignore XML errors — fix them immediately
- ❌ Never assume LibreOffice headless will recalculate fields correctly
