# PPTX SKILL — PowerPoint Manipulation

Covers editing, generating, and validating PPTX files using `python-pptx` and raw XML editing.

---

## 0. Prerequisites

```bash
pip install python-pptx lxml
```

---

## 1. Inspecting a PPTX

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation("presentation.pptx")

print(f"Slide width: {prs.slide_width}, height: {prs.slide_height}")
print(f"Slide layouts: {len(prs.slide_layouts)}")
print(f"Slides: {len(prs.slides)}")

for i, slide in enumerate(prs.slides):
    print(f"\n=== Slide {i+1} ===")
    print(f"  Layout: {slide.slide_layout.name}")
    for shape in slide.shapes:
        print(f"  Shape: {shape.shape_type}, name='{shape.name}', "
              f"pos=({shape.left},{shape.top}), size=({shape.width},{shape.height})")
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if text:
                    print(f"    Text: {text[:80]}")
        if shape.has_table:
            table = shape.table
            print(f"    Table: {len(table.rows)}x{len(table.columns)}")
```

---

## 2. Editing Content

### 2.1 Find and Replace Text

```python
from pptx import Presentation

prs = Presentation("input.pptx")

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "old text" in run.text:
                        run.text = run.text.replace("old text", "new text")

prs.save("output.pptx")
```

### 2.2 Add a Slide

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

prs = Presentation("input.pptx")
slide_layout = prs.slide_layouts[1]  # Title and Content
slide = prs.slides.add_slide(slide_layout)

title = slide.shapes.title
title.text = "New Slide Title"

content = slide.placeholders[1]
content.text = "Bullet point 1\nBullet point 2"

prs.save("output.pptx")
```

### 2.3 Edit Slide Master / Layout

```python
for layout in prs.slide_layouts:
    layout.name = f"Custom - {layout.name}"
    for ph in layout.placeholders:
        print(f"  Placeholder idx={ph.placeholder_format.idx}, type={ph.placeholder_format.type}")
```

---

## 3. Working with Tables

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation("input.pptx")

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.text_frame.paragraphs:
                        for run in para.runs:
                            run.font.size = Pt(11)
                            run.font.name = "Calibri"

prs.save("output.pptx")
```

---

## 4. Images in Slides

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation("input.pptx")
slide = prs.slides[0]

# Add image
slide.shapes.add_picture(
    "image.png",
    Inches(1), Inches(1),  # left, top
    width=Inches(4)         # height auto-scales
)

prs.save("output.pptx")
```

---

## 5. Raw XML Editing (for Complex Changes)

PPTX is also a ZIP file with XML, similar to DOCX:

```bash
mkdir -p work && cd work
cp ../original.pptx ../original_backup.pptx
unzip -q ../original.pptx -d unpacked/
```

Key XML files:

| Path | Content |
|------|---------|
| `ppt/slides/slideN.xml` | Individual slide content |
| `ppt/slides/_rels/slideN.xml.rels` | Slide relationships (images, links) |
| `ppt/slideMasters/slideMasterN.xml` | Slide masters |
| `ppt/slideLayouts/slideLayoutN.xml` | Layouts |
| `ppt/presentation.xml` | Presentation structure |
| `ppt/presentation.xml.rels` | Relationships |

After editing XML:
```bash
cd unpacked && zip -Xr ../output.pptx . && cd ..
```

---

## 6. Validation

```python
from pptx import Presentation
import zipfile
from lxml import etree

def validate_pptx(filepath):
    print(f"=== Validating {filepath} ===")
    with zipfile.ZipFile(filepath, 'r') as z:
        bad = z.testzip()
        if bad:
            print(f"  ✗ Corrupted: {bad}")
            return False
        print(f"  ✓ ZIP valid ({len(z.namelist())} files)")

    prs = Presentation(filepath)
    print(f"  ✓ Opens: {len(prs.slides)} slides")

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

validate_pptx("output.pptx")
```

---

## 7. Common Issues

- **Lost formatting on save** — python-pptx may not preserve all XML-level formatting. For complex slides, edit XML directly.
- **Images missing after rezip** — Check `[Content_Types].xml` and `.rels` files for proper image references.
- **Slide numbers reset** — Edit `ppt/slides/slideN.xml` to add/modify slide number placeholders.
- **Font substitution** — Specify exact font names in XML to avoid cross-platform substitution.
