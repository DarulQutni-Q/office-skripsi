#!/usr/bin/env python3
"""
DOCX Validation Tool — check integrity, structure, and content of a .docx file.

Usage:
    python validate-docx.py path/to/document.docx
"""
import sys
import zipfile
from docx import Document
from lxml import etree


def validate_docx(filepath):
    print(f"=== VALIDATING {filepath} ===\n")

    passed = 0
    failed = 0

    # 1. ZIP integrity
    print("1. ZIP INTEGRITY")
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            bad = z.testzip()
            if bad:
                print(f"  FAIL: Corrupted entry: {bad}")
                failed += 1
            else:
                print(f"  PASS: ZIP valid ({len(z.namelist())} files)")
                passed += 1
    except Exception as e:
        print(f"  FAIL: ZIP error: {e}")
        failed += 1
        return False

    # 2. Document opens
    print("\n2. DOCUMENT OPENS")
    try:
        doc = Document(filepath)
        print(f"  PASS: Opens successfully")
        print(f"    Paragraphs: {len(doc.paragraphs)}")
        print(f"    Tables: {len(doc.tables)}")
        image_count = sum(1 for r in doc.part.rels.values() if 'image' in r.reltype)
        print(f"    Images: {image_count}")
        passed += 1
    except Exception as e:
        print(f"  FAIL: Error: {e}")
        failed += 1
        return False

    # 3. XML validity
    print("\n3. XML VALIDITY")
    xml_errors = []
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            for name in z.namelist():
                if name.endswith('.xml') and 'word/' in name:
                    try:
                        etree.fromstring(z.read(name))
                    except Exception as e:
                        xml_errors.append((name, str(e)))
                        print(f"  FAIL: {name}")
        if xml_errors:
            for name, err in xml_errors:
                print(f"    {name}: {err}")
            failed += 1
        else:
            print(f"  PASS: All XML well-formed")
            passed += 1
    except Exception as e:
        print(f"  FAIL: XML check error: {e}")
        failed += 1

    # 4. Content preservation
    print("\n4. CONTENT PRESERVATION")
    required_sections = [
        "PENDAHULUAN", "TINJAUAN PUSTAKA", "METODE PENELITIAN",
        "HASIL DAN PEMBAHASAN", "KESIMPULAN", "DAFTAR PUSTAKA"
    ]
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            content = z.read('word/document.xml').decode('utf-8')
            missing = []
            for section in required_sections:
                if section not in content:
                    missing.append(section)
                    print(f"  FAIL: {section} MISSING!")
            if not missing:
                print(f"  PASS: All required sections present")
                passed += 1
            else:
                failed += 1
    except Exception as e:
        print(f"  FAIL: Content check error: {e}")
        failed += 1

    # 5. Summary
    print(f"\n{'='*40}")
    print(f"RESULTS: {passed} passed, {failed} failed")
    if failed == 0:
        print("ALL CHECKS PASSED")
        return True
    else:
        print("SOME CHECKS FAILED")
        return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    success = validate_docx(sys.argv[1])
    sys.exit(0 if success else 1)
