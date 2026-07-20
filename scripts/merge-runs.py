#!/usr/bin/env python3
"""
Merge split runs in DOCX XML — combines consecutive <w:r> elements
that share identical formatting (<w:rPr>) into single runs.

Usage:
    python merge-runs.py path/to/document.xml
    # Edits the file in-place
"""
import re
import sys


def merge_runs_in_xml(xml_content):
    def merge_paragraph(match):
        para = match.group(0)
        runs = list(re.finditer(
            r'<w:r\b[^>]*>'
            r'(?:<w:rPr>.*?</w:rPr>)?'
            r'<w:t[^>]*>.*?</w:t>'
            r'</w:r>',
            para, re.DOTALL
        ))
        if len(runs) < 2:
            return para

        merged = []
        i = 0
        while i < len(runs):
            current = runs[i]
            current_text = current.group(0)
            rpr_match = re.search(r'<w:rPr>.*?</w:rPr>', current_text, re.DOTALL)
            current_rpr = rpr_match.group(0) if rpr_match else ''

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

    return re.sub(
        r'<w:p\b[^>]*>.*?</w:p>',
        merge_paragraph,
        xml_content,
        flags=re.DOTALL
    )


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    filepath = sys.argv[1]
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    original_len = len(content)
    content = merge_runs_in_xml(content)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    reduction = original_len - len(content)
    print(f"Merged runs in {filepath}")
    print(f"Size: {original_len} → {len(content)} bytes ({reduction} saved)")


if __name__ == '__main__':
    main()
