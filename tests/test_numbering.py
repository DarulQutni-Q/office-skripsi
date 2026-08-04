#!/usr/bin/env python3
"""Test harness + fixture builder for the multilevel-list automation feature.

Phase 0 (this file): proves the baseline BUG the feature will fix.

The dev's own pipeline (SKILL.md §3) is:
    unpack -> merge-runs.py (unpacked XML) -> edit XML -> rezip -> validate-docx.py

So this harness builds a docx with *hardcoded* numeric prefixes written as plain
text (e.g. "1.1", "2.1.2", "1)"), exactly the anti-pattern the skill currently
teaches. We then assert:

  1. The fixture really contains hardcoded prefixes and NO automatic numbering
     (w:numPr) — i.e. deleting one heading would NOT renumber the rest.
  2. The dev's validate-docx.py passes on that fixture (it is a valid skripsi
     skeleton with the required sections present).

Usage:
    python3 -m unittest tests.test_numbering -v
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile

from docx import Document

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(REPO_ROOT, "scripts")

sys.path.insert(0, SCRIPTS)
import numbering  # noqa: E402

# merge-runs.py has a hyphen, so it loads via file path rather than import.
import importlib.util  # noqa: E402
_MR_SPEC = importlib.util.spec_from_file_location(
    "merge_runs", os.path.join(SCRIPTS, "merge-runs.py"))
merge_runs = importlib.util.module_from_spec(_MR_SPEC)
_MR_SPEC.loader.exec_module(merge_runs)

# Paragraphs laid out as (style_name, text). Headings carry LITERAL numbers.
# The top-level sections match exactly what validate-docx.py requires, so the
# whole fixture is a coherent skripsi skeleton AND validates cleanly.
FIXTURE_PARAGRAPHS = [
    ("Heading 1", "1. PENDAHULUAN"),
    ("Heading 2", "1.1 Latar Belakang"),
    ("Heading 2", "1.2 Rumusan Masalah"),
    ("Heading 1", "2. TINJAUAN PUSTAKA"),
    ("Heading 2", "2.1 Landasan Teori"),
    ("Heading 3", "2.1.1 Definisi Sistem"),
    ("Heading 3", "2.1.2 Pengertian Skripsi"),
    ("Heading 2", "2.2 Kerangka Penelitian"),
    ("Heading 1", "3. METODE PENELITIAN"),
    ("Heading 1", "4. HASIL DAN PEMBAHASAN"),
    ("Heading 1", "5. KESIMPULAN"),
    ("Heading 1", "DAFTAR PUSTAKA"),
    # Numbered BODY list with hardcoded markers (ordered-list anti-pattern).
    ("Normal", "1) Langkah pertama"),
    ("Normal", "2) Langkah kedua"),
    ("Normal", "3) Langkah ketiga"),
]

# Sibling pair used later (Phase 4) to prove cascade renumbering:
# removing "2.1.1 Definisi Sistem" must turn AUTO-numbered "2.1.2" into "2.1.1".
CASCADE_DELETE_TEXT = "2.1.1 Definisi Sistem"
CASCADE_SURVIVOR_TEXT = "2.1.2 Pengertian Skripsi"


def build_fixture(path, paragraphs=None):
    """Create a docx from (style, text) pairs.

    Returns path. The numeric prefixes in the text are plain text — this is the
    hardcoded-numbering bug, not automatic Word multilevel numbering.
    """
    doc = Document()
    for style, text in (paragraphs if paragraphs is not None else FIXTURE_PARAGRAPHS):
        p = doc.add_paragraph(style=style)
        p.add_run(text)
    doc.save(path)
    return path


def has_auto_numbering(path):
    """True if ANY paragraph in the docx already carries w:numPr.

    A doc with w:numPr uses Word's automatic multilevel numbering (paragraph
    order + w:ilvl governs the number). Hardcoded-prefix docs have none.
    """
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return "<w:numPr" in xml


def unpack(path, dest):
    """Unzip a docx into a directory (mirrors `unzip -q x.docx -d unpacked/`)."""
    with zipfile.ZipFile(path) as z:
        z.extractall(dest)
    return dest


def repack(src_dir, path):
    """Zip a directory into a docx with no leading path prefix."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(src_dir):
            for name in files:
                full = os.path.join(root, name)
                z.write(full, os.path.relpath(full, src_dir))
    return path


def unpacked_fixture(tmp, paragraphs=None):
    """Build + unpack a fixture into tmp; returns (docx_path, unpacked_dir)."""
    docx = build_fixture(os.path.join(tmp, "fixture.docx"), paragraphs)
    unpacked = os.path.join(tmp, "unpacked")
    unpack(docx, unpacked)
    return docx, unpacked


def numbering_xml(unpacked):
    with open(os.path.join(unpacked, "word", "numbering.xml"),
              encoding="utf-8") as f:
        return f.read()


def count_multilevel(xml):
    """Number of multilevel definitions in a numbering.xml string."""
    return len(re.findall(r'<w:multiLevelType w:val="multilevel"', xml))


def count_abstracts(xml):
    return len(re.findall(r'<w:abstractNum w:abstractNumId="(\d+)"', xml))


def count_nums(xml):
    return len(re.findall(r'<w:num w:numId="(\d+)"', xml))


def remove_numbering_part(unpacked):
    """Delete word/numbering.xml and undo its package wiring (rels + types).

    Simulates a docx that never had a numbering part, to exercise the
    create-part branch of define.
    """
    num = os.path.join(unpacked, "word", "numbering.xml")
    if os.path.exists(num):
        os.remove(num)
    ct = os.path.join(unpacked, "[Content_Types].xml")
    if os.path.exists(ct):
        with open(ct, encoding="utf-8") as f:
            xml = f.read()
        xml = re.sub(r'<Override[^>]*PartName="/word/numbering.xml"[^>]*/>',
                     "", xml)
        with open(ct, "w", encoding="utf-8") as f:
            f.write(xml)
    rels = os.path.join(unpacked, "word", "_rels", "document.xml.rels")
    if os.path.exists(rels):
        with open(rels, encoding="utf-8") as f:
            xml = f.read()
        xml = re.sub(r'<Relationship[^>]*Target="numbering.xml"[^>]*/>',
                     "", xml)
        with open(rels, "w", encoding="utf-8") as f:
            f.write(xml)


def render_to_pdf(docx_path, out_dir):
    """Render a docx to PDF via LibreOffice; returns the PDF path.

    Skips the caller's test if soffice/libreoffice is not installed.
    """
    so = shutil.which("soffice") or shutil.which("libreoffice")
    if not so:
        raise unittest.SkipTest("soffice/libreoffice not available")
    subprocess.run([so, "--headless", "--convert-to", "pdf",
                    "--outdir", out_dir, docx_path],
                   capture_output=True, text=True, check=True, timeout=120)
    return os.path.join(out_dir,
                        os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")


def pdf_text(pdf_path):
    """Extract laid-out text from a PDF (skips if pdftotext is absent)."""
    pt = shutil.which("pdftotext")
    if not pt:
        raise unittest.SkipTest("pdftotext not available")
    proc = subprocess.run([pt, "-layout", pdf_path, "-"],
                          capture_output=True, text=True)
    return proc.stdout


def remove_paragraph_by_text(unpacked, prefix):
    """Delete the first paragraph whose visible text starts with prefix."""
    path = os.path.join(unpacked, "word", "document.xml")
    with open(path, encoding="utf-8") as f:
        xml = f.read()
    out = []
    last = 0
    removed = False
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.DOTALL):
        out.append(xml[last:m.start()])
        p = m.group(0)
        if not removed and numbering.text_of_paragraph(p).startswith(prefix):
            removed = True
        else:
            out.append(p)
        last = m.end()
    out.append(xml[last:])
    if not removed:
        raise AssertionError(f"paragraph with text prefix {prefix!r} not found")
    with open(path, "w", encoding="utf-8") as f:
        f.write("".join(out))
    return removed


class TestNumberingPhase0(unittest.TestCase):
    """Baseline: the hardcoded-numbering bug is present in the fixture."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="numbering_phase0_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.docx = build_fixture(os.path.join(self.tmp, "fixture.docx"))

    def test_hardcoded_prefixes_stored_as_plain_text(self):
        doc = Document(self.docx)
        texts = [p.text for p in doc.paragraphs]
        self.assertIn("2.1.2 Pengertian Skripsi", texts)
        self.assertIn("2.1.1 Definisi Sistem", texts)
        self.assertIn("1) Langkah pertama", texts)

    def test_bug_is_present_no_automatic_numbering(self):
        # No w:numPr anywhere => numbers are frozen plain text; deleting a
        # heading will NOT renumber the rest. This is the bug the feature fixes.
        self.assertFalse(has_auto_numbering(self.docx),
                         "Fixture unexpectedly uses automatic numbering.")

    def test_cascade_sibling_pair_present(self):
        texts = [p.text for p in Document(self.docx).paragraphs]
        self.assertIn(CASCADE_DELETE_TEXT, texts)
        self.assertIn(CASCADE_SURVIVOR_TEXT, texts)


class TestFixtureIntegrity(unittest.TestCase):
    """Deterministic, LibreOffice-free checks shared by later phases."""

    def test_validate_docx_passes_on_fixture(self):
        """The dev's real validate-docx.py must exit 0 on the fixture."""
        with tempfile.TemporaryDirectory() as tmp:
            path = build_fixture(os.path.join(tmp, "f.docx"))
            proc = subprocess.run(
                [sys.executable, os.path.join(SCRIPTS, "validate-docx.py"), path],
                capture_output=True, text=True)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_validate_docx_rejects_corrupt_doc(self):
        """Sanity: validate-docx must fail on a truncated (non-docx) file."""
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bad.docx")
            with open(path, "w", encoding="utf-8") as f:
                f.write("not a docx")
            proc = subprocess.run(
                [sys.executable, os.path.join(SCRIPTS, "validate-docx.py"), path],
                capture_output=True, text=True)
            self.assertNotEqual(proc.returncode, 0)


class TestAnalyzePhase1(unittest.TestCase):
    """analyze subcommand: read-only audit, correct level detection, CLI gate."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="numbering_phase1_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        _, self.unpacked = unpacked_fixture(self.tmp)

    def _records(self):
        return numbering.analyze_document(
            self.unpacked, numbering.DEFAULT_STYLE_LEVELS)

    def _by_text(self, records, text):
        return [r for r in records if r["text"].startswith(text)]

    def test_literal_prefix_levels(self):
        cases = {
            "1. PENDAHULUAN": 0,
            "1.1 Latar Belakang": 1,
            "2.1.2 Pengertian Skripsi": 2,
            "1) Langkah pertama": 0,
        }
        records = self._records()
        for text, level in cases.items():
            hit = self._by_text(records, text)
            self.assertEqual(len(hit), 1, f"no record for {text!r}")
            self.assertEqual(hit[0]["level"], level, text)

    def test_hardcoded_statuses_and_warn(self):
        statuses = {r["text"]: r["status"] for r in self._records()}
        self.assertEqual(statuses["1. PENDAHULUAN"], numbering._STATUS_HARDCODED)
        self.assertEqual(statuses["2.1.1 Definisi Sistem"],
                         numbering._STATUS_HARDCODED)
        self.assertEqual(statuses["1) Langkah pertama"], numbering._STATUS_HARDCODED)
        # Heading with no number at all -> warn (informational), not hardcoded.
        self.assertEqual(statuses["DAFTAR PUSTAKA"],
                         numbering._STATUS_WARN_NO_NUMBER)

    def test_auto_detection(self):
        p_xml = ('<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:numPr>'
                 '<w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr>'
                 '<w:r><w:t>Latar Belakang</w:t></w:r></w:p>')
        info = numbering.paragraph_info(p_xml)
        status, level = numbering.classify(info, numbering.DEFAULT_STYLE_LEVELS)
        self.assertEqual(status, numbering._STATUS_AUTO)
        self.assertEqual(level, 1)
        self.assertEqual(info["numid"], 7)

    def test_paragraph_text_decode(self):
        p_xml = ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
                 '<w:r><w:t>2.1.2 Biaya &amp; Jadwal</w:t></w:r></w:p>')
        info = numbering.paragraph_info(p_xml)
        self.assertEqual(info["text"], "2.1.2 Biaya & Jadwal")
        self.assertEqual(info["style"], "Heading1")

    def test_cli_exit_1_when_hardcoded(self):
        proc = subprocess.run(
            [sys.executable, os.path.join(SCRIPTS, "numbering.py"),
             self.unpacked, "analyze"],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HARDCODED", proc.stdout)
        self.assertIn("FAIL:", proc.stdout)
        self.assertIn("RESULT:", proc.stdout)

    def test_cli_exit_0_when_clean(self):
        clean = [
            ("Heading 1", "PENDAHULUAN"),
            ("Heading 2", "Latar Belakang"),
            ("Normal", "Teks biasa tanpa nomor"),
        ]
        _, unpacked = unpacked_fixture(self.tmp, clean)
        proc = subprocess.run(
            [sys.executable, os.path.join(SCRIPTS, "numbering.py"),
             unpacked, "analyze"],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("PASS:", proc.stdout)


class TestDefinePhase2(unittest.TestCase):
    """define subcommand: multilevel definition, idempotency, merge safety."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="numbering_phase2_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        _, self.unpacked = unpacked_fixture(self.tmp)

    def test_define_creates_one_multilevel_definition(self):
        res = numbering.ensure_numbering(self.unpacked)
        self.assertTrue(res["created_abstract"])
        self.assertTrue(res["created_num"])
        self.assertFalse(res["created_numbering_part"])  # template has the part
        self.assertIsNotNone(res["abstract_num_id"])
        self.assertIsNotNone(res["num_id"])
        xml = numbering_xml(self.unpacked)
        self.assertEqual(count_multilevel(xml), 1)

    def test_define_default_level_texts_and_format(self):
        numbering.ensure_numbering(self.unpacked)
        xml = numbering_xml(self.unpacked)
        aid = numbering.abstract_num_ids(xml)
        block = re.search(rf'<w:abstractNum w:abstractNumId="{aid[-1]}">(.*?)'
                          r'</w:abstractNum>', xml, re.DOTALL).group(1)
        texts = re.findall(r'<w:lvlText w:val="([^"]*)"', block)
        fmts = re.findall(r'<w:numFmt w:val="([^"]*)"', block)
        starts = re.findall(r'<w:start w:val="([^"]*)"', block)
        self.assertEqual(texts, ["%1", "%1.%2", "%1.%2.%3"])
        self.assertEqual(fmts, ["decimal", "decimal", "decimal"])
        self.assertEqual(starts, ["1", "1", "1"])
        self.assertEqual(len(re.findall(r'<w:lvl ', block)), 3)

    def test_define_is_idempotent(self):
        first = numbering.ensure_numbering(self.unpacked)
        second = numbering.ensure_numbering(self.unpacked)
        self.assertEqual(first["num_id"], second["num_id"])
        self.assertEqual(first["abstract_num_id"], second["abstract_num_id"])
        self.assertFalse(second["created_abstract"])
        self.assertFalse(second["created_num"])
        self.assertEqual(count_multilevel(numbering_xml(self.unpacked)), 1)
        # 9 stock singleLevel abstractNums + our new one; 9 stock nums + one.
        self.assertEqual(count_abstracts(numbering_xml(self.unpacked)), 10)
        self.assertEqual(count_nums(numbering_xml(self.unpacked)), 10)

    def test_define_does_not_clobber_stock_singlelevel(self):
        before = numbering_xml(self.unpacked)
        before_abstracts = numbering.abstract_num_ids(before)
        before_nums = numbering.num_ids(before)
        res = numbering.ensure_numbering(self.unpacked)
        xml = numbering_xml(self.unpacked)
        # Stock template defs are untouched; the new ids sit above them.
        self.assertIn(res["abstract_num_id"], numbering.abstract_num_ids(xml))
        self.assertIn(res["num_id"], numbering.num_ids(xml))
        self.assertNotIn(res["abstract_num_id"], before_abstracts)
        self.assertNotIn(res["num_id"], before_nums)
        self.assertGreater(res["abstract_num_id"], max(before_abstracts))
        self.assertGreater(res["num_id"], max(before_nums))
        # The stock singleLevel defs are all still singleLevel.
        self.assertEqual(count_multilevel(xml), 1)

    def test_define_custom_levels_and_text(self):
        # Default 3-level first, then a distinct 2-level def: both must coexist.
        numbering.ensure_numbering(self.unpacked)
        res = numbering.ensure_numbering(
            self.unpacked, levels=2,
            lvl_texts=["(1)", "(1.1)"], num_fmt="lowerRoman")
        xml = numbering_xml(self.unpacked)
        block = re.search(rf'<w:abstractNum w:abstractNumId="'
                          rf'{res["abstract_num_id"]}">(.*?)</w:abstractNum>',
                          xml, re.DOTALL).group(1)
        texts = re.findall(r'<w:lvlText w:val="([^"]*)"', block)
        fmts = re.findall(r'<w:numFmt w:val="([^"]*)"', block)
        self.assertEqual(texts, ["(1)", "(1.1)"])
        self.assertEqual(fmts, ["lowerRoman", "lowerRoman"])
        self.assertEqual(count_multilevel(xml), 2)  # 3-level + 2-level coexist
        # Idempotent across a different-params call too (params drive reuse).
        again = numbering.ensure_numbering(
            self.unpacked, levels=2,
            lvl_texts=["(1)", "(1.1)"], num_fmt="lowerRoman")
        self.assertFalse(again["created_abstract"])
        self.assertFalse(again["created_num"])
        self.assertEqual(count_multilevel(numbering_xml(self.unpacked)), 2)

    def test_define_creates_part_and_wiring_when_missing(self):
        remove_numbering_part(self.unpacked)
        res = numbering.ensure_numbering(self.unpacked)
        self.assertTrue(res["created_numbering_part"])
        self.assertTrue(res["wired_rels"])
        self.assertTrue(res["wired_content_types"])
        self.assertEqual(res["abstract_num_id"], 0)
        self.assertEqual(res["num_id"], 1)
        xml = numbering_xml(self.unpacked)
        self.assertEqual(count_abstracts(xml), 1)
        self.assertEqual(count_nums(xml), 1)
        # Re-run: part exists now, still idempotent.
        again = numbering.ensure_numbering(self.unpacked)
        self.assertEqual(again["num_id"], 1)
        self.assertEqual(count_abstracts(xml), 1)

    def test_define_cli_exit_0(self):
        proc = subprocess.run(
            [sys.executable, os.path.join(SCRIPTS, "numbering.py"),
             self.unpacked, "define"],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("PASS:", proc.stdout)
        self.assertIn("numId=", proc.stdout)

    def test_define_cli_repack_validate_passes(self):
        """define then repack must still pass the dev's validate-docx.py."""
        subprocess.run(
            [sys.executable, os.path.join(SCRIPTS, "numbering.py"),
             self.unpacked, "define"],
            capture_output=True, text=True, check=True)
        out_docx = os.path.join(self.tmp, "defined.docx")
        repack(self.unpacked, out_docx)
        proc = subprocess.run(
            [sys.executable, os.path.join(SCRIPTS, "validate-docx.py"), out_docx],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


def document_xml(unpacked):
    with open(os.path.join(unpacked, "word", "document.xml"),
              encoding="utf-8") as f:
        return f.read()


class TestConvertPhase3(unittest.TestCase):
    """convert subcommand: strip hardcoded prefixes + link w:numPr."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="numbering_phase3_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        _, self.unpacked = unpacked_fixture(self.tmp)
        self.result = numbering.convert_document(
            self.unpacked, numbering.DEFAULT_STYLE_LEVELS)

    def _records(self):
        return numbering.analyze_document(
            self.unpacked, numbering.DEFAULT_STYLE_LEVELS)

    def _by_text(self, records, text):
        return [r for r in records if r["text"].startswith(text)]

    def test_convert_returns_num_id_and_changed_count(self):
        self.assertIsNotNone(self.result["num_id"])
        # 11 numbered headings + 3 numbered body paragraphs = 14; the one
        # unnumbered heading (DAFTAR PUSTAKA) is left alone.
        self.assertEqual(self.result["changed"], 14)

    def test_all_headings_now_auto(self):
        statuses = {r["text"]: r["status"] for r in self._records()}
        self.assertEqual(statuses["PENDAHULUAN"], numbering._STATUS_AUTO)
        self.assertEqual(statuses["Latar Belakang"], numbering._STATUS_AUTO)
        self.assertEqual(statuses["Pengertian Skripsi"], numbering._STATUS_AUTO)
        # Unnumbered heading stays WARN (it has no number at all).
        self.assertEqual(statuses["DAFTAR PUSTAKA"],
                         numbering._STATUS_WARN_NO_NUMBER)

    def test_prefixes_stripped_from_text(self):
        texts = [r["text"] for r in self._records()]
        self.assertIn("PENDAHULUAN", texts)
        self.assertIn("Latar Belakang", texts)
        self.assertIn("Pengertian Skripsi", texts)
        self.assertIn("Langkah pertama", texts)
        self.assertNotIn("2.1.1 Definisi Sistem", texts)
        self.assertNotIn("1. PENDAHULUAN", texts)
        self.assertNotIn("1) Langkah pertama", texts)

    def test_body_list_linked_at_level_zero(self):
        hits = self._by_text(self._records(), "Langkah")
        self.assertEqual(len(hits), 3)
        for r in hits:
            self.assertEqual(r["status"], numbering._STATUS_AUTO)
            self.assertEqual(r["level"], 0)

    def test_levels_match_styles(self):
        cases = {"PENDAHULUAN": 0, "Latar Belakang": 1,
                 "Pengertian Skripsi": 2}
        records = self._records()
        for text, level in cases.items():
            hit = self._by_text(records, text)
            self.assertEqual(len(hit), 1, text)
            self.assertEqual(hit[0]["level"], level, text)
            self.assertEqual(hit[0]["ilvl"], level, text)

    def test_numpr_placed_after_pstyle(self):
        xml = document_xml(self.unpacked)
        m = re.search(r"<w:pPr><w:pStyle[^>]*Heading1[^>]*/>.*?</w:pPr>",
                      xml, re.DOTALL)
        self.assertIsNotNone(m)
        self.assertLess(m.group(0).index("<w:numPr"),
                        len(m.group(0)))
        self.assertGreater(m.group(0).index("<w:numPr"),
                           m.group(0).index("<w:pStyle"))

    def test_convert_is_idempotent(self):
        first = self.result
        second = numbering.convert_document(
            self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        self.assertEqual(second["changed"], 0)
        self.assertEqual(second["num_id"], first["num_id"])
        self.assertEqual(document_xml(self.unpacked),
                         numbering.read_part(self.unpacked,
                                             numbering.DOCUMENT_PART))

    def test_convert_reuses_define(self):
        # ensure_numbering ran inside convert; a separate define must reuse.
        first = self.result
        res = numbering.ensure_numbering(self.unpacked)
        self.assertEqual(res["num_id"], first["num_id"])
        self.assertEqual(self._count_multilevel(), 1)

    def _read_numbering(self):
        with open(os.path.join(self.unpacked, "word", "numbering.xml"),
                  encoding="utf-8") as f:
            return f.read()

    def _count_multilevel(self):
        return len(re.findall(r'<w:multiLevelType w:val="multilevel"',
                              self._read_numbering()))

    def test_convert_cli_exit_0_and_analyze_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, unpacked = unpacked_fixture(tmp)
            proc = subprocess.run(
                [sys.executable, os.path.join(SCRIPTS, "numbering.py"),
                 unpacked, "convert"],
                capture_output=True, text=True)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertIn("PASS:", proc.stdout)
            self.assertIn("numId=", proc.stdout)
            # Re-audit: no hardcoded prefixes remain.
            audit = subprocess.run(
                [sys.executable, os.path.join(SCRIPTS, "numbering.py"),
                 unpacked, "analyze"],
                capture_output=True, text=True)
            self.assertEqual(audit.returncode, 0, audit.stdout + audit.stderr)

    def test_convert_cli_repack_validate_passes(self):
        out_docx = os.path.join(self.tmp, "converted.docx")
        repack(self.unpacked, out_docx)
        proc = subprocess.run(
            [sys.executable, os.path.join(SCRIPTS, "validate-docx.py"),
             out_docx],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class TestRenderPhase4(unittest.TestCase):
    """Render-verified cascade, restart & idempotency (needs LibreOffice)."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="numbering_phase4_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.render_dir = os.path.join(self.tmp, "render")
        os.makedirs(self.render_dir)
        _, self.unpacked = unpacked_fixture(self.tmp)

    def _repack(self, name="conv.docx"):
        path = os.path.join(self.tmp, name)
        repack(self.unpacked, path)
        return path

    def _render_text(self):
        return pdf_text(render_to_pdf(self._repack(), self.render_dir))

    def test_headings_render_multilevel(self):
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        txt = self._render_text()
        self.assertIsNotNone(re.search(r"1\s+PENDAHULUAN", txt))
        self.assertIsNotNone(re.search(r"1\.1\s+Latar Belakang", txt))
        self.assertIsNotNone(re.search(r"2\.1\.2\s+Pengertian Skripsi", txt))
        self.assertIsNotNone(re.search(r"3\s+METODE PENELITIAN", txt))

    def test_heading_prefix_is_stripped(self):
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        txt = self._render_text().replace("\t", " ")
        self.assertNotIn("1. 1 PENDAHULUAN", txt)
        self.assertNotIn("2.1.1 Definisi", txt)

    def test_body_list_restarts_per_section(self):
        # Body list sits in its own section (after DAFTAR PUSTAKA) -> 1,2,3.
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        txt = self._render_text()
        self.assertIsNotNone(re.search(r"1\s+Langkah pertama", txt))
        self.assertIsNotNone(re.search(r"2\s+Langkah kedua", txt))
        self.assertIsNotNone(re.search(r"3\s+Langkah ketiga", txt))
        self.assertIsNone(re.search(r"6\s+Langkah", txt))

    def test_delete_section_renumbers_survivor(self):
        """CASCADE: delete 'Definisi Sistem'; 'Pengertian Skripsi' -> 2.1.1."""
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        remove_paragraph_by_text(self.unpacked, "Definisi Sistem")
        txt = self._render_text()
        self.assertIsNone(re.search(r"2\.1\.2\s+Pengertian Skripsi", txt))
        self.assertIsNotNone(re.search(r"2\.1\.1\s+Pengertian Skripsi", txt))

    def test_convert_twice_is_rendered_idempotent(self):
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        txt = self._render_text()
        self.assertIsNotNone(re.search(r"1\.1\s+Latar Belakang", txt))
        # no "1.1 1.1" concatenation and no leftover literal prefix
        self.assertIsNone(re.search(r"1\.1\s+1\.1", txt))
        self.assertNotIn("1. 1.1 Latar", txt.replace("\t", " "))

    def test_two_body_lists_restart_per_section_in_render(self):
        """Two chapters, each with a body list: BOTH must render 1,2 (never
        4,5,6). LibreOffice continues the counter across nums sharing one
        abstract, so each list needs its own w:startOverride instance."""
        paras = [
            ("Heading 1", "1. PENDAHULUAN"),
            ("Normal", "1) a satu"),
            ("Normal", "2) a dua"),
            ("Heading 1", "2. TINJAUAN PUSTAKA"),
            ("Normal", "1) b satu"),
            ("Normal", "2) b dua"),
            ("Normal", "3) b tiga"),
        ]
        _, unpacked = unpacked_fixture(self.tmp, paras)
        numbering.convert_document(unpacked, numbering.DEFAULT_STYLE_LEVELS)
        txt = pdf_text(render_to_pdf(repack(unpacked, os.path.join(
            self.tmp, "multi.docx")), self.render_dir))
        self.assertIsNotNone(re.search(r"1\s+a satu", txt))
        self.assertIsNotNone(re.search(r"2\s+a dua", txt))
        self.assertIsNotNone(re.search(r"1\s+b satu", txt))
        self.assertIsNotNone(re.search(r"3\s+b tiga", txt))
        self.assertIsNone(re.search(r"4\s+b", txt),
                          "second chapter's list must restart at 1, got continuation")


class TestBodyRestartUnit(unittest.TestCase):
    """LibreOffice-free checks of the restart-per-section layout."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="numbering_phase4u_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        _, self.unpacked = unpacked_fixture(self.tmp)

    def _nid_for_text(self, prefix):
        xml = document_xml(self.unpacked)
        for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.DOTALL):
            if numbering.text_of_paragraph(m.group(0)).startswith(prefix):
                info = numbering.paragraph_info(m.group(0))
                return info["numid"], info["ilvl"]
        self.fail(f"paragraph {prefix!r} not found")

    def test_body_list_uses_distinct_num_from_headings(self):
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        head_nid, _ = self._nid_for_text("PENDAHULUAN")
        body_nid, ilvl = self._nid_for_text("Langkah")
        self.assertNotEqual(body_nid, head_nid)
        self.assertEqual(ilvl, 0)

    def test_restart_per_section_allocates_one_num_per_section(self):
        paras = [
            ("Heading 1", "A. Satu"),
            ("Normal", "1) a1"),
            ("Normal", "2) a2"),
            ("Heading 1", "B. Dua"),
            ("Normal", "1) b1"),
        ]
        _, unpacked = unpacked_fixture(self.tmp, paras)
        res = numbering.convert_document(unpacked, numbering.DEFAULT_STYLE_LEVELS)
        self.assertEqual(res["body_sections"], 2)
        xml = document_xml(unpacked)
        body_nids = set()
        for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.DOTALL):
            info = numbering.paragraph_info(m.group(0))
            if info["numid"] is not None and \
               info["style"] not in numbering.DEFAULT_STYLE_LEVELS:
                body_nids.add(info["numid"])
        self.assertEqual(len(body_nids), 2)   # one body counter per section

    def test_body_nums_carry_start_override(self):
        """Each per-section body list instance must force restart at 1 via
        w:lvlOverride/w:startOverride so LibreOffice (and Word) renumber it."""
        paras = [
            ("Heading 1", "1. SATU"),
            ("Normal", "1) a1"),
            ("Heading 1", "2. DUA"),
            ("Normal", "1) b1"),
        ]
        _, unpacked = unpacked_fixture(self.tmp, paras)
        numbering.convert_document(unpacked, numbering.DEFAULT_STYLE_LEVELS)
        xml = document_xml(unpacked)
        body_nids = [
            numbering.paragraph_info(m.group(0))["numid"]
            for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.DOTALL)
            if numbering.paragraph_info(m.group(0))["style"]
            not in numbering.DEFAULT_STYLE_LEVELS
            and numbering.paragraph_info(m.group(0))["numid"] is not None
        ]
        nb = numbering_xml(self.unpacked)
        for nid in sorted(set(body_nids)):
            block = re.search(rf'<w:num w:numId="{nid}">(.*?)</w:num>',
                              nb, re.DOTALL).group(1)
            self.assertIn('<w:startOverride w:val="1"', block,
                          f"body num {nid} lacks restart override")
        # The heading (multilevel) num must NOT be force-restarted per section.
        head_nid = numbering.paragraph_info(
            next(m.group(0) for m in re.finditer(
                r"<w:p\b[^>]*>.*?</w:p>", xml, re.DOTALL)
                if numbering.text_of_paragraph(m.group(0)) == "SATU"))["numid"]
        head_block = re.search(rf'<w:num w:numId="{head_nid}">(.*?)</w:num>',
                               nb, re.DOTALL).group(1)
        self.assertNotIn("startOverride", head_block)


class TestWordSchemaOrder(unittest.TestCase):
    """Regression tests for the two Word-strict OOXML order bugs.

    LibreOffice and validate-docx.py (well-formed XML only) tolerate both; real
    Word flags "unreadable content, repair?" for either, so they MUST be pinned.
    """

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="numbering_schema_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        _, self.unpacked = unpacked_fixture(self.tmp)

    # --- Bug 2: CT_Numbering order (all <w:num> after the last <w:abstractNum>)

    @staticmethod
    def abstract_positions(xml):
        """List of (start,end) for every <w:abstractNum> block."""
        return [(m.start(), m.end())
                for m in re.finditer(r'</w:abstractNum>', xml)]

    @staticmethod
    def num_starts(xml):
        return [m.start() for m in re.finditer(r'<w:num w:numId=', xml)]

    def assert_abstracts_before_nums(self, xml, label):
        nums = self.num_starts(xml)
        if not nums:
            return
        last_abs_end = max(e for _, e in self.abstract_positions(xml))
        # No abstract may start after the first <w:num> begins; adjacency (==)
        # is the ideal "all abstracts, then all nums" layout.
        self.assertLessEqual(last_abs_end, nums[0],
                             f"{label}: abstractNum block appears after <w:num> "
                             f"(abs_end={last_abs_end}, first num={nums[0]})")

    def test_insert_helpers_keep_order(self):
        xml = numbering.read_part(self.unpacked, numbering.NUMBERING_PART)
        # Build the realistic convert sequence of in-memory inserts:
        #   heading abstract -> heading num -> body abstract -> body num
        xml, h_numid, _ = numbering.build_numbering(
            xml, 3, numbering.default_lvl_texts(3), ["decimal"] * 3, [0] * 3)
        xml, aid2 = numbering.ensure_single_abstract(xml, "decimal", "%1")
        xml, b_numid = numbering.alloc_num(xml, aid2)
        self.assert_abstracts_before_nums(xml, "insert-helper order")
        # Both allocated nums must be valid ids (reuse or fresh both fine).
        self.assertIn(h_numid, numbering.num_ids(xml))
        self.assertIn(b_numid, numbering.num_ids(xml))

    def test_convert_output_numbering_order(self):
        numbering.convert_document(self.unpacked, numbering.DEFAULT_STYLE_LEVELS)
        xml = numbering_xml(self.unpacked)
        self.assert_abstracts_before_nums(xml, "convert(): numbering.xml")

    def test_define_output_numbering_order(self):
        numbering.ensure_numbering(self.unpacked)
        xml = numbering_xml(self.unpacked)
        self.assert_abstracts_before_nums(xml, "define(): numbering.xml")

    # --- Bug 1: CT_PPrBase order (numPr after pStyle/keepNext/etc., before spacing)

    NUM_PRE = ["<w:pStyle", "<w:keepNext", "<w:keepLines",
               "<w:pageBreakBefore", "<w:framePr", "<w:widowControl"]
    NUM_POST = ["<w:suppressLineNumbers", "<w:pBdr", "<w:shd", "<w:tabs",
                "<w:spacing", "<w:ind", "<w:jc", "<w:outlineLvl", "<w:rPr"]

    def assert_numpr_after(self, ppr, markers):
        idx = ppr.index("<w:numPr")
        for m in markers:
            self.assertNotEqual(ppr.index(m), -1, f"{m} missing in {ppr}")
            self.assertLess(ppr.index(m), idx,
                            f"{m} must come BEFORE numPr (got {ppr})")

    def assert_numpr_before(self, ppr, markers):
        idx = ppr.index("<w:numPr")
        for m in markers:
            self.assertNotEqual(ppr.index(m), -1, f"{m} missing in {ppr}")
            self.assertGreater(ppr.index(m), idx,
                               f"{m} must come AFTER numPr (got {ppr})")

    def test_numpr_after_keepnext_and_widow_control(self):
        p_xml = ('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/>'
                 '<w:widowControl w:val="0"/><w:spacing w:after="120"/>'
                 '<w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>')
        out = numbering.add_numpr_to_paragraph(p_xml, 0, 5)
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', out, re.DOTALL).group(0)
        self.assert_numpr_after(ppr, self.NUM_PRE[:2])          # pStyle, keepNext
        self.assert_numpr_before(ppr, ["<w:spacing", "<w:outlineLvl"])

    def test_numpr_after_keepnext_preceding_pstyle(self):
        # keepNext appears BEFORE pStyle in the source pPr; numPr still after both.
        p_xml = ('<w:p><w:pPr><w:keepNext/><w:pStyle w:val="Heading1"/>'
                 '<w:spacing w:before="240"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>')
        out = numbering.add_numpr_to_paragraph(p_xml, 0, 5)
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', out, re.DOTALL).group(0)
        self.assert_numpr_after(ppr, self.NUM_PRE[:2])          # keepNext, pStyle
        self.assert_numpr_before(ppr, ["<w:spacing"])

    def test_numpr_first_when_no_predecessors(self):
        p_xml = ('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                 '<w:r><w:t>x</w:t></w:r></w:p>')
        out = numbering.add_numpr_to_paragraph(p_xml, 0, 5)
        ppr = re.search(r'<w:pPr>.*?</w:pPr>', out, re.DOTALL).group(0)
        self.assert_numpr_before(ppr, ["<w:jc"])

    def test_self_closing_ppr_becomes_single_ppr_with_numpr(self):
        """python-docx/Word emit <w:pPr/> for empty paragraph props; converting
        must rewrite it in place, never add a second <w:pPr> (CT_P violation)."""
        p_xml = '<w:p><w:pPr w:rsidR="00AB12CD"/><w:r><w:t>1) item</w:t></w:r></w:p>'
        out = numbering.add_numpr_to_paragraph(p_xml, 0, 5)
        self.assertEqual(out.count("<w:pPr"), 1,
                         f"expected one <w:pPr>, got: {out}")
        # attrs preserved on the rewritten open tag, numPr inside, no leftover.
        self.assertIn('<w:pPr w:rsidR="00AB12CD"><w:numPr>', out)
        self.assertIn("</w:numPr></w:pPr>", out)
        self.assertNotIn("<w:pPr/>", out)

    def test_self_closing_ppr_convert_full_pipeline(self):
        """End-to-end on a fixture paragraph that starts as <w:pPr/>: after
        convert exactly one pPr exists and it holds the numPr."""
        paras = [("Normal", "1) kata"), ("Normal", "2) kunci")]
        _, unpacked = unpacked_fixture(self.tmp, paras)
        # python-docx emits <w:pPr/> for style="Normal" paragraphs; the fixture
        # is already in the Word-realistic form this bug was about.
        numbering.convert_document(unpacked, numbering.DEFAULT_STYLE_LEVELS)
        out = document_xml(unpacked)
        for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", out, re.DOTALL):
            p = m.group(0)
            if numbering.text_of_paragraph(p).startswith("kata"):
                self.assertEqual(p.count("<w:pPr"), 1, p)
                self.assertIn("<w:numPr>", p)
                self.assertNotIn("<w:pPr/>", p)

    def test_realistic_heading_pPr_after_convert_order(self):
        # A Word-style heading with paragraph-level keepNext + spacing/outlineLvl
        # is the exact case LibreOffice/validate missed; assert full order.
        paras = [
            ("Heading 1", "1. PENDAHULUAN"),
            ("Heading 2", "1.1 Latar Belakang"),
        ]
        _, unpacked = unpacked_fixture(self.tmp, paras)
        # Inject Word-style block props into the Heading1 paragraph's pPr.
        doc = document_xml(unpacked)
        new = ('<w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/>'
               '<w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/>'
               '</w:pPr>')
        patched = re.sub(r'<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>', new, doc,
                         count=1)
        self.assertNotEqual(patched, doc)
        with open(os.path.join(unpacked, "word", "document.xml"), "w",
                  encoding="utf-8") as f:
            f.write(patched)
        numbering.convert_document(unpacked, numbering.DEFAULT_STYLE_LEVELS)
        out = document_xml(unpacked)
        ppr = re.search(r'<w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/>'
                        r'.*?</w:pPr>', out, re.DOTALL).group(0)
        order = [ppr.index("<w:pStyle"), ppr.index("<w:keepNext"),
                 ppr.index("<w:numPr"), ppr.index("<w:spacing"),
                 ppr.index("<w:outlineLvl")]
        self.assertEqual(order, sorted(order),
                         f"pPr child order violates CT_PPrBase: {ppr}")


class TestMergeRunsRegressions(unittest.TestCase):
    """Regressions for two merge-runs.py bugs found during deep real-world
    testing on a Word-emission-style stress docx.

    Bug A (catastrophic): the paragraph rebuild loop re-located merged runs via
    ``para.find(constructed_string, pos)``. For split runs that string never
    exists verbatim, ``find`` returned -1, ``pos`` then jumped too far, and
    ``result += para[pos:]`` swallowed the rest of the paragraph — deleting up
    to several whole paragraphs and corrupting the document.

    Bug B: the run regex ``<w:t[^>]*>`` also matched ``<w:tab/>`` (prefix
    match), swallowing content across run boundaries; and merging could run
    across a non-text run (e.g. ``<w:tab/>``), dropping it.
    """

    @staticmethod
    def merge(xml):
        return merge_runs.merge_runs_in_xml(xml)

    def test_multi_run_paragraphs_are_preserved(self):
        xml = ('<w:document xmlns:w="http://schemas.openxmlformats.org/'
               'wordprocessingml/2006/main"><w:body>'
               '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>'
               '<w:r><w:t>1.1</w:t></w:r><w:r><w:t> Latar Belakang</w:t></w:r>'
               '</w:p>'
               '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>'
               '<w:r><w:t>2.1</w:t></w:r><w:r><w:t>.</w:t></w:r>'
               '<w:r><w:t>1 Definisi Sistem</w:t></w:r></w:p>'
               '</w:body></w:document>')
        out = self.merge(xml)
        paras = re.findall(r'<w:p\b[^>]*>.*?</w:p>', out, re.DOTALL)
        self.assertEqual(len(paras), 2, "merge-runs deleted whole paragraphs")
        texts = [re.sub(r'<[^>]+>', '', p) for p in paras]
        self.assertEqual(texts, ["1.1 Latar Belakang", "2.1.1 Definisi Sistem"])

    def test_tab_run_not_matched_and_not_crossed(self):
        xml = ('<w:document xmlns:w="http://schemas.openxmlformats.org/'
               'wordprocessingml/2006/main"><w:body>'
               '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>'
               '<w:r><w:t>1.2</w:t></w:r><w:r><w:tab/></w:r>'
               '<w:r><w:t xml:space="preserve"> Rumusan Masalah</w:t></w:r>'
               '</w:p></w:body></w:document>')
        out = self.merge(xml)
        # The two text runs are NOT consecutive (a <w:tab/> sits between them),
        # so they must NOT be merged — the tab must survive untouched.
        self.assertIn('<w:r><w:tab/></w:r>', out,
                      "merge-runs dropped a <w:tab/> run")
        self.assertEqual(out.count('<w:r>'), 3, out)
        text = "".join(re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', out, re.DOTALL))
        self.assertEqual(text, "1.2 Rumusan Masalah")

    def test_merge_only_combines_identical_rpr(self):
        xml = ('<w:document xmlns:w="http://schemas.openxmlformats.org/'
               'wordprocessingml/2006/main"><w:body>'
               '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>'
               '<w:r><w:t>2.1</w:t></w:r>'
               '<w:r><w:rPr><w:b/></w:rPr><w:t>.</w:t></w:r>'
               '<w:r><w:rPr><w:b/></w:rPr><w:t>1 Definisi</w:t></w:r>'
               '</w:p></w:body></w:document>')
        out = self.merge(xml)
        text = "".join(re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', out, re.DOTALL))
        self.assertEqual(text, "2.1.1 Definisi")
        # exactly two runs left: [plain "2.1"] + [bold ".1 Definisi"]
        self.assertEqual(out.count('<w:r>'), 2, out)


if __name__ == "__main__":
    unittest.main(verbosity=2)