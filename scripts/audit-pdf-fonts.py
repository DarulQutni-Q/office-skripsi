#!/usr/bin/env python3
"""Audit which fonts actually render text in a PDF.

Usage:
    pdftohtml -xml -i doc.pdf /tmp/doc.xml
    python3 audit-pdf-fonts.py /tmp/doc.xml [--target "Times"]
    python3 audit-pdf-fonts.py /tmp/doc.xml --target "Times" --pages 30-70

Reports every font (excluding the target) that carries visible text,
grouped by page. Zero "Calibri" in the DOCX package is NOT proof of
compliance — this is the check that catches theme-font inheritance
(styles without rFonts inheriting docDefaults asciiTheme=minorHAnsi).
"""
import argparse
import re


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xml", help="output of `pdftohtml -xml -i doc.pdf out.xml`")
    ap.add_argument("--target", default="Times",
                    help="acceptable font family substring (default: Times)")
    ap.add_argument("--pages", default=None,
                    help="restrict to page range, e.g. 30-70")
    args = ap.parse_args()

    x = open(args.xml, encoding="utf-8", errors="ignore").read()
    fam = {}
    for m in re.finditer(r'<fontspec id="(\d+)"[^>]*family="([^"]*)"', x):
        fam[m.group(1)] = m.group(2)

    page = None
    hits = {}
    for line in x.splitlines():
        m = re.match(r'<page number="(\d+)"', line)
        if m:
            page = m.group(1)
            continue
        m = re.search(r'<text[^>]*font="(\d+)"[^>]*>([^<]{1,60})</text>', line)
        if m and m.group(2).strip():
            f = fam.get(m.group(1), "?")
            if args.target not in f:
                hits.setdefault((page, f), []).append(m.group(2).strip())

    if args.pages:
        a, b = (int(p) for p in args.pages.split("-"))
        hits = {(p, f): t for (p, f), t in hits.items() if a <= int(p) <= b}

    if not hits:
        print("OK: all visible text renders in", args.target)
        return
    for (pg, f), ts in sorted(hits.items(), key=lambda k: (int(k[0][0]), k[0][1])):
        print("page %s %s -> %s" % (pg, f, [t for t in ts][:8]))


if __name__ == "__main__":
    main()
