#!/usr/bin/env python3
"""
DOCX Analysis Toolkit — inspect, analyze, and check thesis documents.

Usage:
    python docx-tools.py path/to/document.docx [command]

Commands:
    headings    List all headings with style names
    captions    List all captions with SEQ type
    tables      List all tables and their content
    citations   Extract in-text citations
    seq         List all SEQ fields with cached numbers
    all         Run all checks
"""
import sys
import re
from docx import Document


def cmd_headings(doc):
    print("=== HEADINGS ===")
    for i, p in enumerate(doc.paragraphs):
        if p.style and "Heading" in p.style.name:
            print(f"[{i:4d}] {p.style.name:15s} {p.text.strip()[:80]}")


def cmd_captions(doc):
    print("=== CAPTIONS ===")
    for i, p in enumerate(doc.paragraphs):
        if p.style and "Caption" in p.style.name:
            has_seq_tabel = 'SEQ Tabel' in p._element.xml
            has_seq_gambar = 'SEQ Gambar' in p._element.xml
            seq_type = "TABEL" if has_seq_tabel else ("GAMBAR" if has_seq_gambar else "NONE")
            print(f"[{i:4d}] [{seq_type:6s}] {p.text.strip()[:60]}")


def cmd_tables(doc):
    print("=== TABLES ===")
    for i, t in enumerate(doc.tables):
        print(f"Table {i}: {len(t.rows)} rows x {len(t.columns)} cols")
        for j, row in enumerate(t.rows[:5]):
            if j >= 5:
                print(f"  ... ({len(t.rows) - 5} more rows)")
                break
            cells = [c.text.strip()[:40] for c in row.cells]
            print(f"  Row {j}: {cells}")


def cmd_citations(doc):
    print("=== CITATIONS ===")
    all_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    citations = re.findall(r'\(([^)]+\d{4}[^)]*)\)', all_text)
    for c in citations:
        print(f"  {c.strip()}")
    print(f"\nTotal: {len(citations)} citations found")


def cmd_seq(doc):
    print("=== SEQ FIELDS ===")
    all_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    seqs = re.findall(r'(Tabel|Gambar)\s+(\d+(?:\.\d+)?)\s*[—–\-]?\s*([^\n]*)', all_text)
    for seq_type, num, title in seqs:
        print(f"  {seq_type} {num}: {title.strip()[:60]}")
    print(f"\nTotal: {len(seqs)} SEQ references")


def cmd_all(doc):
    cmd_headings(doc)
    print()
    cmd_captions(doc)
    print()
    cmd_tables(doc)
    print()
    cmd_citations(doc)
    print()
    cmd_seq(doc)


COMMANDS = {
    'headings': cmd_headings,
    'captions': cmd_captions,
    'tables': cmd_tables,
    'citations': cmd_citations,
    'seq': cmd_seq,
    'all': cmd_all,
}


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    filepath = sys.argv[1]
    command = sys.argv[2] if len(sys.argv) > 2 else 'all'

    if command not in COMMANDS:
        print(f"Unknown command: {command}")
        print(f"Available: {', '.join(COMMANDS.keys())}")
        sys.exit(1)

    doc = Document(filepath)
    COMMANDS[command](doc)


if __name__ == '__main__':
    main()
