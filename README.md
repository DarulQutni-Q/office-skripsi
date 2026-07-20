# Office Skill — AI Agent Office Document Manipulation

A comprehensive skill for AI agents (OpenCode, Claude Code, etc.) to edit, generate, and validate Office documents programmatically — DOCX, PPTX, PDF, and XLSX — without requiring Microsoft Office.

## What This Is

This is a **skill pack** designed to be loaded by AI coding agents. It teaches the agent exactly how to:

- **DOCX**: Unzip → edit XML → rezip (no Office needed). Handle headings, TOC, captions, tables, styles, cross-references.
- **PPTX**: Edit slide content, replace text/images, manage layouts, regenerate slide masters.
- **PDF**: Extract text, convert to/from DOCX via LibreOffice, verify visual output.
- **Spreadsheets**: Read/write XLSX, manipulate cells/ranges, apply formatting.

## Quick Start

```bash
# Load the skill in your AI agent config:
# In opencode.json:
{
  "skills": ["/path/to/office-skill/SKILL.md"]
}
```

Or use sub-skills for specific formats:

```bash
skills/docx/SKILL.md    # Word documents
skills/pptx/SKILL.md    # PowerPoint
skills/pdf/SKILL.md     # PDF handling
skills/spreadsheet/SKILL.md  # Spreadsheets
```

## Prerequisites

| Tool | Purpose |
|------|---------|
| `python3` + `python-docx` | DOCX read/write |
| `python3` + `python-pptx` | PPTX read/write |
| `python3` + `openpyxl` | XLSX read/write |
| `LibreOffice` (`soffice`) | DOCX ↔ PDF conversion |
| `poppler-utils` (`pdftotext`, `pdftoppm`) | PDF text/image extraction |
| `zip`/`unzip` | DOCX/PPTX internals (they're ZIP files) |
| `lxml` | XML validation |

## Project Structure

```
office-skill/
├── SKILL.md                    # Main skill (master orchestrator)
├── package.json                # npx compatibility
├── index.js                    # CLI entry point
├── .gitignore                  # Ignore personal files
├── README.md                   # This file
├── skills/
│   ├── docx/SKILL.md           # DOCX sub-skill
│   ├── pptx/SKILL.md           # PPTX sub-skill
│   ├── pdf/SKILL.md            # PDF sub-skill
│   └── spreadsheet/SKILL.md    # Spreadsheet sub-skill
├── scripts/
│   ├── validate-docx.py        # DOCX validation tool
│   ├── docx-tools.py           # DOCX analysis toolkit
│   └── merge-runs.py           # Merge split runs in DOCX XML
├── templates/                  # Reusable document templates
└── examples/                   # Example usage patterns
```

## Key Principles

1. **DOCX/PPTX are ZIP files** — Always unzip first, edit XML, rezip. Never edit the binary `.docx`/`.pptx` directly.
2. **Verify before claiming** — Render to PDF, extract text, grep for correctness. Don't assume edits worked.
3. **Backup before every edit** — Keep `original_backup.docx` alongside the working copy.
4. **One edit at a time** — Validate after each batch. Finding one bad XML change is easy; finding one among twenty is not.
5. **Source code > screenshots > text** — When content conflicts, the code (if available) is the ground truth.

## License

MIT
