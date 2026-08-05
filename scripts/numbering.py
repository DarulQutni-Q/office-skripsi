#!/usr/bin/env python3
"""DOCX Multilevel Numbering Tool.

Converts hardcoded "1.2.3" numbering (typed as plain text) into Word's
automatic multilevel numbering (w:numPr + word/numbering.xml), so the numbers
are derived from paragraph order + w:ilvl and renumber themselves on insert /
delete / level change.

Pipeline placement (SKILL.md §3): runs in the unpack phase, right AFTER
    scripts/merge-runs.py unpacked/word/document.xml

Usage:
    python3 numbering.py <unpacked_dir> analyze  [--style-level H1=0,H2=1,...]
    python3 numbering.py <unpacked_dir> define   [--levels N] [--lvl-text ...]
                                                 [--num-fmt FMT] [--indent ...]
    python3 numbering.py <unpacked_dir> convert  [...]   (see help)

analyze — read-only audit. Lists every heading / numbered-list paragraph with
its intended level, its literal prefix and its numbering state, and exits 1 if
any hardcoded prefixes (the "deleting one doesn't renumber the rest" bug) are
found. Safe to run at any time.

define — ensure a multilevel numbering definition exists in word/numbering.xml.
Merge-safe and idempotent: reuses a matching definition and picks the next free
ids, so existing numbering is never clobbered. Prints the resulting numId for
use by convert.

convert — define + link + strip. Ensures the definition, attaches w:numPr
(ilvl + numId) to every heading / numbered-list paragraph, and deletes the
literal "1.2.3" prefix from the text, so the numbers become automatic and
renumber on insert / delete / level change. Idempotent; re-audits afterwards
and exits 1 if any hardcoded prefixes remain.
"""
import argparse
import re
import sys
from collections import Counter
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DOCUMENT_PART = Path("word/document.xml")
NUMBERING_PART = Path("word/numbering.xml")

# Default style -> multilevel level (0-based, like w:ilvl / outlineLvl).
DEFAULT_STYLE_LEVELS = {
    "Heading1": 0,
    "Heading2": 1,
    "Heading3": 2,
    "Heading4": 3,
    "Heading5": 4,
    "Heading6": 5,
    "Heading7": 6,
    "Heading8": 7,
    "Heading9": 8,
}

# Literal numeric prefixes an agent might have typed by hand:
#   "1. ", "1.1 ", "2.1.2 "   (dotted decimal, any depth)
#   "1) ", "2) ", "1. "       (single-level ordered lists)
# A trailing "." or ")" separator is optional when a space / end follows, so
# "1 PENDAHULUAN" (dot-less headings) is caught too.
NUM_PREFIX_RE = re.compile(r"^(\d+(?:\.\d+)*)[.)]?(?:\s+|$)")

_PARAGRAPH_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.DOTALL)

_STATUS_HARDCODED = "HARDCODED"
_STATUS_AUTO = "AUTO"
_STATUS_WARN_NO_NUMBER = "WARN_NO_NUMBER"
_STATUS_OK = "OK"

# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------


def read_part(unpacked_dir, part):
    """Return text content of a part (e.g. Path('word/document.xml'))."""
    return (Path(unpacked_dir) / part).read_text(encoding="utf-8")


def write_part(unpacked_dir, part, content):
    (Path(unpacked_dir) / part).write_text(content, encoding="utf-8")


def split_paragraphs(xml):
    """Yield each <w:p>…</w:p> block. Safe: <w:p> cannot nest."""
    return _PARAGRAPH_RE.findall(xml)


def _decode(xml_text):
    """Decode XML entities back to plain characters inside <w:t>."""
    return (xml_text.replace("&amp;", "&")
                    .replace("&lt;", "<")
                    .replace("&gt;", ">")
                    .replace("&quot;", '"')
                    .replace("&apos;", "'"))


def text_of_paragraph(p_xml):
    """Concatenated visible text of a paragraph (all <w:t> contents)."""
    chunks = re.findall(r"<w:t\b[^>]*>(.*?)</w:t>", p_xml, re.DOTALL)
    return _decode("".join(chunks))


def paragraph_info(p_xml):
    """Extract (style, numPr/ilvl/numId, text) from one <w:p> block."""
    ppr = re.search(r"<w:pPr\b[^>]*>.*?</w:pPr>", p_xml, re.DOTALL)
    ppr_xml = ppr.group(0) if ppr else ""

    style = None
    m = re.search(r"<w:pStyle[^>]*w:val=\"([^\"]*)\"", ppr_xml)
    if m:
        style = m.group(1)

    numpr = ilvl = numid = None
    m = re.search(r"<w:numPr>(.*?)</w:numPr>", ppr_xml, re.DOTALL)
    if m:
        numpr = m.group(0)
        m2 = re.search(r"<w:ilvl[^>]*w:val=\"(\d+)\"", numpr)
        if m2:
            ilvl = int(m2.group(1))
        m3 = re.search(r"<w:numId[^>]*w:val=\"(\d+)\"", numpr)
        if m3:
            numid = int(m3.group(1))

    return {
        "style": style,
        "numpr": numpr,
        "ilvl": ilvl,
        "numid": numid,
        "text": text_of_paragraph(p_xml),
    }


# ---------------------------------------------------------------------------
# Numbering analysis
# ---------------------------------------------------------------------------


def literal_prefix(text):
    """Detect a hardcoded numeric prefix at the start of a paragraph text.

    Returns (prefix, level): the exact matched prefix string (including the
    trailing separator) and the 0-based multilevel level it implies
    ("1." -> 0, "1.1" -> 1, "2.1.2" -> 2, "1)" -> 0). (None, None) if none.
    """
    m = NUM_PREFIX_RE.match(text)
    if m:
        digits = m.group(1)
        return m.group(0), digits.count(".")
    return None, None


def _is_plausible_list_prefix(text):
    """True when a literal numeric prefix is a REAL typed list marker.

    NUM_PREFIX_RE alone also matches prose that merely starts with a number
    ("2026 merupakan…", "10 metode yang…", "2.5 kg beras…") — converting those
    would STRIP the number from the text and misnumber the paragraph as a list
    item. Body lists are typed with an explicit separator (".", ")") or a
    dotted multi-level number ("1.1", "2.1.2") whose first word starts
    uppercase (a list title), which prose never satisfies.
    """
    m = NUM_PREFIX_RE.match(text)
    if not m:
        return False
    digits = m.group(1)
    if len(digits) == 4 and digits.isdigit() and 1900 <= int(digits) <= 2099:
        return False                     # "2024 …" is a year, not a list item
    rest = text[len(digits):]
    if rest.startswith((".", ")")):
        return True                      # "1.", "1)", "2.1.2)", "10."
    if "." in digits:
        after_space = rest.lstrip(" \t")
        return bool(after_space) and after_space[0].isupper()
    return False                         # "10 metode…", "2.5 kg…" are prose


def classify(info, style_levels):
    """Return (status, level) for one paragraph.

    AUTO              -> already uses w:numPr (automatic numbering)
    HARDCODED         -> literal numeric prefix typed as text (the bug)
    WARN_NO_NUMBER    -> heading style but neither number nor w:numPr
    OK                -> nothing numbering-related
    """
    prefix, lit_level = literal_prefix(info["text"])
    style_level = style_levels.get(info["style"]) if info["style"] else None

    if info["numpr"]:
        return _STATUS_AUTO, info["ilvl"]
    if style_level is not None:
        if prefix:
            return _STATUS_HARDCODED, style_level
        return _STATUS_WARN_NO_NUMBER, style_level
    if prefix and _is_plausible_list_prefix(info["text"]):
        return _STATUS_HARDCODED, lit_level
    return _STATUS_OK, None


def analyze_document(unpacked_dir, style_levels):
    """Audit every paragraph of word/document.xml. Returns list of records.

    Each record: {idx, style, level, status, ilvl, numid, prefix, text}.
    """
    xml = read_part(unpacked_dir, DOCUMENT_PART)
    records = []
    for idx, p_xml in enumerate(split_paragraphs(xml)):
        info = paragraph_info(p_xml)
        status, level = classify(info, style_levels)
        prefix, _ = literal_prefix(info["text"])
        records.append({
            "idx": idx,
            "style": info["style"],
            "level": level,
            "status": status,
            "ilvl": info["ilvl"],
            "numid": info["numid"],
            "prefix": prefix,
            "text": info["text"],
        })
    return records


def format_records(records):
    """Render analyze/convert records into aligned display lines."""
    lines = []
    for r in records:
        level = "" if r["level"] is None else f"L{r['level']}"
        lines.append(f"  [{r['idx']:>4}] {str(r['style'] or '-'):<11} {level:<3} "
                     f"{r['status']:<16} {r['text'][:60]}")
    return lines


def report_analyze(records, unpacked_dir, style_levels):
    """Print the audit report; returns exit code (1 if hardcoded found)."""
    print("=== NUMBERING ANALYZE ===")
    print(f"part: {Path(unpacked_dir) / DOCUMENT_PART}")
    map_str = ", ".join(f"{k}={v}" for k, v in
                        sorted(style_levels.items(), key=lambda kv: kv[1]))
    print(f"style map: {map_str}")
    print()
    for line in format_records(records):
        print(line)
    counts = Counter(r["status"] for r in records)
    print()
    print(f"RESULT: {counts.get(_STATUS_HARDCODED, 0)} hardcoded, "
          f"{counts.get(_STATUS_AUTO, 0)} auto-numbered, "
          f"{counts.get(_STATUS_WARN_NO_NUMBER, 0)} heading without number, "
          f"{counts.get(_STATUS_OK, 0)} ok")
    if counts.get(_STATUS_HARDCODED, 0):
        print("FAIL: hardcoded numeric prefixes found (they will NOT renumber "
              "when one is deleted) — run: numbering.py <unpacked> convert")
        return 1
    print("PASS: no hardcoded prefixes")
    return 0


# ---------------------------------------------------------------------------
# Numbering definition (word/numbering.xml)
# ---------------------------------------------------------------------------

NUMBERING_HEADER = ('<w:numbering xmlns:w='
                    '"http://schemas.openxmlformats.org/wordprocessingml/2006/main">')
NUMBERING_CONTENT_TYPE = ("application/vnd.openxmlformats-officedocument."
                          "wordprocessingml.numbering+xml")
NUMBERING_REL_TYPE = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                      "relationships/numbering")

_ABSTRACT_RE = re.compile(r'<w:abstractNum w:abstractNumId="(\d+)">(.*?)'
                          r'</w:abstractNum>', re.DOTALL)
_NUM_RE = re.compile(r'<w:num w:numId="(\d+)">(.*?)</w:num>', re.DOTALL)


def _escape_xml(value):
    return (value.replace("&", "&amp;").replace("<", "&lt;")
                 .replace(">", "&gt;"))


def default_lvl_texts(levels):
    """Per-level number text: level 0 "%1", level 1 "%1.%2", level 2 "%1.%2.%3"."""
    return [".".join(f"%{j + 1}" for j in range(i + 1)) for i in range(levels)]


def build_lvl(ilvl, num_fmt, lvl_text, indent):
    """One <w:lvl> element for the given multilevel level."""
    indent = max(0, int(indent))
    if indent:
        ppr = (f'<w:pPr><w:tabs><w:tab w:val="num" w:pos="{indent}"/></w:tabs>'
               f'<w:ind w:left="{indent}" w:hanging="{indent}"/></w:pPr>')
    else:
        ppr = '<w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr>'
    return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/>'
            f'<w:numFmt w:val="{num_fmt}"/>'
            f'<w:lvlText w:val="{_escape_xml(lvl_text)}"/>'
            f'<w:lvlJc w:val="left"/>{ppr}</w:lvl>')


def build_abstract_num(abstract_num_id, levels, lvl_texts, num_fmts, indents):
    """A complete <w:abstractNum> block (multilevel, decimal per level)."""
    parts = [f'<w:abstractNum w:abstractNumId="{abstract_num_id}">',
             '<w:multiLevelType w:val="multilevel"/>']
    for i in range(levels):
        parts.append(build_lvl(i, num_fmts[i], lvl_texts[i], indents[i]))
    parts.append('</w:abstractNum>')
    return "".join(parts)


def build_num(num_id, abstract_num_id, force_restart=False):
    """One <w:num> referencing an abstract.

    force_restart=True adds a w:lvlOverride / w:startOverride (ilvl 0 -> start
    1) so the instance restarts its counter at 1 deterministically. Word starts
    each num instance at the abstract's start anyway, but LibreOffice continues
    the counter across every num sharing one abstract, so body-list restarts
    MUST carry an explicit override to render identically in both.
    """
    override = ""
    if force_restart:
        override = ('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/>'
                    '</w:lvlOverride>')
    return (f'<w:num w:numId="{num_id}">'
            f'<w:abstractNumId w:val="{abstract_num_id}"/>{override}</w:num>')


def build_numpr(ilvl, num_id):
    """The w:numPr fragment to attach a paragraph to a multilevel level."""
    return (f'<w:numPr><w:ilvl w:val="{ilvl}"/>'
            f'<w:numId w:val="{num_id}"/></w:numPr>')


def abstract_num_ids(xml):
    return sorted(int(m.group(1)) for m in _ABSTRACT_RE.finditer(xml))


def num_ids(xml):
    return sorted(int(m.group(1)) for m in _NUM_RE.finditer(xml))


def num_for_abstract(xml, abstract_num_id):
    """Return the first numId whose w:num references abstract_num_id, or None."""
    for m in _NUM_RE.finditer(xml):
        inner = m.group(2)
        ref = re.search(r'<w:abstractNumId[^>]*w:val="(\d+)"', inner)
        if ref and int(ref.group(1)) == abstract_num_id:
            return int(m.group(1))
    return None


def find_compatible_abstract_id(xml, levels, lvl_texts, num_fmts):
    """Reuse an existing abstractNum identical to ours (idempotency).

    A candidate must be multilevel and match our per-level numFmt and lvlText
    sequences exactly; the stock singleLevel definitions never match.
    """
    for m in _ABSTRACT_RE.finditer(xml):
        block = m.group(2)
        mt = re.search(r'<w:multiLevelType w:val="([^"]*)"', block)
        if not mt or mt.group(1) != "multilevel":
            continue
        texts = re.findall(r'<w:lvlText w:val="([^"]*)"', block)
        fmts = re.findall(r'<w:numFmt w:val="([^"]*)"', block)
        if texts == lvl_texts and fmts == num_fmts:
            return int(m.group(1))
    return None


def _insert_after_last_abstract(xml, block):
    """Insert a <w:abstractNum> block after the last existing one.

    CT_Numbering requires every <w:num> to come after the LAST <w:abstractNum>,
    so new abstract definitions must go after the tail of the abstract section
    (i.e. before the first <w:num>), never after an existing <w:num>.
    """
    idx = -1
    for m in _ABSTRACT_RE.finditer(xml):
        idx = m.end()
    if idx < 0:
        # No abstractNum yet: place right after <w:numbering ...> open tag.
        open_end = xml.find(">") + 1
        if open_end <= 0 or not xml.startswith("<w:numbering"):
            raise ValueError("malformed numbering.xml: missing <w:numbering>")
        return xml[:open_end] + block + xml[open_end:]
    return xml[:idx] + block + xml[idx:]


def _insert_before_numbering_close(xml, block):
    """Insert a <w:num> block right before </w:numbering>.

    Kept for <w:num> insertion only: nums legally follow the whole abstract
    section, so placing them just before the closing tag is schema-valid
    (provided abstract inserts go through _insert_after_last_abstract).
    """
    idx = xml.rfind("</w:numbering>")
    if idx < 0:
        raise ValueError("malformed numbering.xml: missing </w:numbering>")
    return xml[:idx] + block + xml[idx:]


def _ensure_rels_wire(unpacked_dir, result):
    """Add the numbering relationship to document.xml.rels if absent."""
    rels_path = Path(unpacked_dir) / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        rels_path.parent.mkdir(parents=True, exist_ok=True)
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<Relationships xmlns="http://schemas.openxmlformats.org/'
               'package/2006/relationships"></Relationships>')
    else:
        xml = rels_path.read_text(encoding="utf-8")
    if NUMBERING_REL_TYPE in xml:
        result["wired_rels"] = False
        return
    used = [int(r) for r in re.findall(r'Id="rId(\d+)"', xml)]
    new_id = (max(used) + 1) if used else 1
    entry = (f'<Relationship Id="rId{new_id}" Type="{NUMBERING_REL_TYPE}" '
             f'Target="numbering.xml"/>')
    xml = xml.replace("</Relationships>", entry + "</Relationships>")
    rels_path.write_text(xml, encoding="utf-8")
    result["wired_rels"] = True


def _ensure_content_type(unpacked_dir, result):
    """Add the numbering Override to [Content_Types].xml if absent."""
    ct_path = Path(unpacked_dir) / "[Content_Types].xml"
    if not ct_path.exists():
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
               'content-types"></Types>')
    else:
        xml = ct_path.read_text(encoding="utf-8")
    if "numbering" in xml:
        result["wired_content_types"] = False
        return
    override = (f'<Override PartName="/word/numbering.xml" '
                f'ContentType="{NUMBERING_CONTENT_TYPE}"/>')
    xml = xml.replace("</Types>", override + "</Types>")
    ct_path.write_text(xml, encoding="utf-8")
    result["wired_content_types"] = True


def build_numbering(num_xml, levels, lvl_texts, num_fmts, indents):
    """Ensure a multilevel abstract + a num exist in num_xml (in memory).

    Reuses a compatible definition / num when present (idempotent), otherwise
    appends one using the next free ids. Returns (num_xml, abstract_id, num_id).
    """
    aid = find_compatible_abstract_id(num_xml, levels, lvl_texts, num_fmts)
    if aid is None:
        ids = abstract_num_ids(num_xml)
        aid = (max(ids) + 1) if ids else 0
        num_xml = _insert_after_last_abstract(
            num_xml, build_abstract_num(aid, levels, lvl_texts, num_fmts, indents))
    num_id = num_for_abstract(num_xml, aid)
    if num_id is None:
        num_xml, num_id = alloc_num(num_xml, aid)
    return num_xml, aid, num_id


def find_compatible_single_abstract(xml, num_fmt, lvl_text):
    """Reuse an existing singleLevel abstractNum matching our body list format."""
    for m in _ABSTRACT_RE.finditer(xml):
        block = m.group(2)
        mt = re.search(r'<w:multiLevelType w:val="([^"]*)"', block)
        if not mt or mt.group(1) != "singleLevel":
            continue
        texts = re.findall(r'<w:lvlText w:val="([^"]*)"', block)
        fmts = re.findall(r'<w:numFmt w:val="([^"]*)"', block)
        if texts == [lvl_text] and fmts == [num_fmt]:
            return int(m.group(1))
    return None


def build_single_abstract(abstract_num_id, num_fmt, lvl_text):
    """A complete single-level <w:abstractNum> (for body lists)."""
    return (f'<w:abstractNum w:abstractNumId="{abstract_num_id}">'
            f'<w:multiLevelType w:val="singleLevel"/>'
            f'{build_lvl(0, num_fmt, lvl_text, 0)}'
            f'</w:abstractNum>')


def ensure_single_abstract(num_xml, num_fmt, lvl_text):
    """Ensure a single-level abstract for body lists exists; returns (xml, aid)."""
    aid = find_compatible_single_abstract(num_xml, num_fmt, lvl_text)
    if aid is None:
        ids = abstract_num_ids(num_xml)
        aid = (max(ids) + 1) if ids else 0
        num_xml = _insert_after_last_abstract(
            num_xml, build_single_abstract(aid, num_fmt, lvl_text))
    return num_xml, aid


def alloc_num(num_xml, abstract_num_id, force_restart=False):
    """Append a fresh <w:num> referencing abstract_num_id; returns (xml, num_id)."""
    nids = num_ids(num_xml)
    num_id = (max(nids) + 1) if nids else 1
    num_xml = _insert_before_numbering_close(
        num_xml, build_num(num_id, abstract_num_id, force_restart=force_restart))
    return num_xml, num_id


def ensure_numbering(unpacked_dir, levels=3, lvl_texts=None,
                     num_fmt="decimal", indents=None):
    """Make sure a compatible multilevel definition exists in numbering.xml.

    Merge-safe: never clobbers existing numbering; picks the next free
    abstractNumId / numId, and reuses a matching definition (idempotent).

    Returns dict with num_id, abstract_num_id and what was created/wired.
    """
    if not 1 <= levels <= 9:
        raise ValueError("levels must be in 1..9")
    lvl_texts = lvl_texts or default_lvl_texts(levels)
    num_fmts = [num_fmt] * levels
    indents = indents or [0] * levels
    if len(lvl_texts) != levels or len(indents) != levels:
        raise ValueError("lvl_texts / indents must have exactly 'levels' entries")

    result = {"num_id": None, "abstract_num_id": None,
              "created_numbering_part": False, "created_abstract": False,
              "created_num": False, "wired_rels": False,
              "wired_content_types": False}
    num_path = Path(unpacked_dir) / NUMBERING_PART

    if not num_path.exists():
        aid, num_id = 0, 1
        num_xml = (NUMBERING_HEADER
                   + build_abstract_num(aid, levels, lvl_texts, num_fmts, indents)
                   + build_num(num_id, aid)
                   + "</w:numbering>")
        num_path.write_text(num_xml, encoding="utf-8")
        result.update(created_numbering_part=True, created_abstract=True,
                      created_num=True, abstract_num_id=aid, num_id=num_id)
        _ensure_rels_wire(unpacked_dir, result)
        _ensure_content_type(unpacked_dir, result)
        return result

    num_xml = num_path.read_text(encoding="utf-8")
    # Snapshot so we can report what was added.
    matched = find_compatible_abstract_id(num_xml, levels, lvl_texts, num_fmts)
    num_before = num_for_abstract(num_xml, matched) if matched is not None else None
    num_xml, aid, num_id = build_numbering(num_xml, levels, lvl_texts,
                                           num_fmts, indents)
    result.update(abstract_num_id=aid, num_id=num_id,
                  created_abstract=(matched is None), created_num=(num_before is None))
    num_path.write_text(num_xml, encoding="utf-8")
    return result


def report_define(result, unpacked_dir, levels, lvl_texts, indents):
    """Print the define report; returns exit code (always 0)."""
    print("=== NUMBERING DEFINE ===")
    print(f"part: {Path(unpacked_dir) / NUMBERING_PART}")
    print(f"levels: {levels}")
    print(f"lvl-text: {' | '.join(lvl_texts)}")
    print(f"num-fmt: decimal  indent: {', '.join(str(i) for i in indents)}")
    print()
    print(f"abstractNum: {result['abstract_num_id']} "
          f"({'created' if result['created_abstract'] else 'reused'})")
    print(f"num: {result['num_id']} "
          f"({'created' if result['created_num'] else 'reused'})")
    print(f"numbering part: "
          f"{'created' if result['created_numbering_part'] else 'already present'}")
    print(f"content types: "
          f"{'wired' if result['wired_content_types'] else 'already ok'}")
    print(f"relationships: "
          f"{'wired' if result['wired_rels'] else 'already ok'}")
    print()
    print(f"PASS: multilevel numbering defined (numId={result['num_id']})")
    return 0


# ---------------------------------------------------------------------------
# Numbering conversion (analyze -> strip -> link)
# ---------------------------------------------------------------------------


def strip_prefix_text(p_xml, prefix):
    """Remove `prefix` from the start of a paragraph's visible text.

    The prefix may span multiple <w:t> runs; the first len(prefix) characters
    of the concatenated text are removed by editing whichever runs contain
    them. Edits raw XML, which is safe because the stripped region is plain
    ASCII (digits / dots / parens / whitespace) and never contains entities.
    """
    remaining = len(prefix)
    if remaining <= 0:
        return p_xml
    t_re = re.compile(r"(<w:t\b[^>]*>)(.*?)(</w:t>)", re.DOTALL)
    out = []
    last = 0
    for m in t_re.finditer(p_xml):
        open_tag, content, close_tag = m.group(1), m.group(2), m.group(3)
        out.append(p_xml[last:m.start()])
        if remaining > 0:
            cut = min(remaining, len(content))
            content = content[cut:]
            remaining -= cut
        out.append(open_tag + content + close_tag)
        last = m.end()
    out.append(p_xml[last:])
    if remaining:
        raise ValueError(f"prefix {prefix!r} not found in paragraph text")
    return "".join(out)


def add_numpr_to_paragraph(p_xml, level, num_id):
    """Return p_xml with a w:numPr (ilvl=level, numId=num_id) in its pPr.

    pPr is created if absent; numPr is inserted at the schema-correct position
    for CT_PPrBase — after the last of the siblings that must PRECEDE numPr
    (pStyle, keepNext, keepLines, pageBreakBefore, framePr, widowControl) and
    before everything else (spacing, ind, jc, outlineLvl, rPr, ...). Removes an
    existing numPr first so a re-convert never stacks duplicate numPr.
    """
    numpr = build_numpr(level, num_id)
    ppr = re.search(r"<w:pPr\b[^>]*>.*?</w:pPr>", p_xml, re.DOTALL)
    if not ppr:
        # Self-closing <w:pPr/> (Word & python-docx emit these for empty
        # paragraph properties). Rewrite it in place — inserting a second pPr
        # would violate CT_P (one pPr per paragraph) and make Word offer to
        # repair the file. Preserve any attributes (e.g. w:rsidR).
        sc = re.search(r"<w:pPr\b([^>]*)/>", p_xml)
        if sc:
            attrs = sc.group(1)
            new = f"<w:pPr{attrs}>{numpr}</w:pPr>"
            return p_xml[:sc.start()] + new + p_xml[sc.end():]
        m = re.match(r"(^\s*<w:p\b[^>]*>)", p_xml)
        if not m:
            raise ValueError("malformed paragraph: no <w:p> open tag")
        numpr_xml = f"<w:pPr>{numpr}</w:pPr>"
        return p_xml[:m.end()] + numpr_xml + p_xml[m.end():]

    body = ppr.group(0)
    gt = body.index(">")                     # end of "<w:pPr …>"
    # Remove any existing numPr so re-converting stays idempotent.
    sew = re.sub(r"<w:numPr>.*?</w:numPr>", "", body, flags=re.DOTALL)
    if "<w:numPr" in sew:
        raise ValueError("unexpected existing w:numPr in paragraph")
    # Find how many schema-precedessors of numPr follow <w:pPr>.
    pre = re.compile(
        r'^\s*(?:<w:pStyle\b[^>]*/>|'
        r'<w:keepNext\s*/>|<w:keepLines\s*/>|<w:pageBreakBefore\s*/>|'
        r'<w:framePr\b[^>]*/>|<w:framePr\b[^>]*>.*?</w:framePr>|'
        r'<w:widowControl\b[^>]*/>)')
    content = sew[gt + 1:]
    ins = 0
    while True:
        m = pre.match(content[ins:])
        if not m:
            break
        ins += m.end()
    new_body = sew[:gt + 1 + ins] + numpr + sew[gt + 1 + ins:]
    return p_xml[:ppr.start()] + new_body + p_xml[ppr.end():]


def convert_document(unpacked_dir, style_levels, levels=3, lvl_texts=None,
                     num_fmt="decimal", indents=None,
                     list_num_fmt="decimal", list_lvl_text="%1"):
    """Strip hardcoded prefixes and attach w:numPr to every numbered element.

    define (ensure a multilevel def + get numId) -> link (w:numPr) -> strip
    (delete the literal prefix).

    Headings get the shared multilevel numId at their style level. Body
    numbered lists (Non-heading paragraphs with a literal prefix) get their
    own single-level instance, and each top-level section restarts it at 1
    with a fresh numId carrying a w:startOverride (independent counters that
    render identically in Word and LibreOffice). Idempotent: already-AUTO
    paragraphs are left alone.

    Returns dict {num_id, abstract_num_id, changed, body_sections}.
    """
    if not 1 <= levels <= 9:
        raise ValueError("levels must be in 1..9")
    lvl_texts = lvl_texts or default_lvl_texts(levels)
    num_fmts = [num_fmt] * levels
    indents = indents or [0] * levels
    if len(lvl_texts) != levels or len(indents) != levels:
        raise ValueError("lvl_texts / indents must have exactly 'levels' entries")

    num_path = Path(unpacked_dir) / NUMBERING_PART
    if num_path.exists():
        num_xml = num_path.read_text(encoding="utf-8")
        part_created = False
    else:
        num_xml = NUMBERING_HEADER + "</w:numbering>"
        part_created = True

    num_xml, heading_aid, heading_numid = build_numbering(
        num_xml, levels, lvl_texts, num_fmts, indents)
    num_xml, body_aid = ensure_single_abstract(
        num_xml, num_fmt=list_num_fmt, lvl_text=list_lvl_text)

    xml = read_part(unpacked_dir, DOCUMENT_PART)
    out = []
    last = 0
    changed = 0
    section = 0
    section_numid = {}
    for m in _PARAGRAPH_RE.finditer(xml):
        out.append(xml[last:m.start()])
        p_xml = m.group(0)
        info = paragraph_info(p_xml)
        style = info["style"]
        if style in style_levels and style_levels[style] == 0:
            section += 1                     # top-level heading -> new section
        status, level = classify(info, style_levels)
        if status == _STATUS_HARDCODED and level is not None:
            prefix, _ = literal_prefix(info["text"])
            if prefix:
                if style in style_levels:
                    target_numid, target_level = heading_numid, level
                else:
                    # body list item: per-section counter (restarts at 1).
                    target_numid = section_numid.get(section)
                    if target_numid is None:
                        num_xml, target_numid = alloc_num(
                            num_xml, body_aid, force_restart=True)
                        section_numid[section] = target_numid
                    target_level = 0
                p_xml = strip_prefix_text(p_xml, prefix)
                p_xml = add_numpr_to_paragraph(p_xml, target_level, target_numid)
                changed += 1
        out.append(p_xml)
        last = m.end()
    out.append(xml[last:])

    write_part(unpacked_dir, NUMBERING_PART, num_xml)
    write_part(unpacked_dir, DOCUMENT_PART, "".join(out))
    if part_created:
        _ensure_rels_wire(unpacked_dir, {})
        _ensure_content_type(unpacked_dir, {})

    return {"num_id": heading_numid, "abstract_num_id": heading_aid,
            "changed": changed, "body_sections": len(section_numid)}


def report_convert(result, unpacked_dir, style_levels):
    """Print the convert report; re-audits and exits 1 if any hardcoded remain."""
    print("=== NUMBERING CONVERT ===")
    print(f"numId: {result['num_id']}  abstractNum: {result['abstract_num_id']}")
    print(f"body list instances: {result.get('body_sections', 0)} "
          f"(one per top-level section, restarting at 1)")
    print(f"paragraphs converted: {result['changed']}")
    print()
    records = analyze_document(unpacked_dir, style_levels)
    for line in format_records(records):
        print(line)
    counts = Counter(r["status"] for r in records)
    print()
    print(f"RESULT: {counts.get(_STATUS_HARDCODED, 0)} hardcoded, "
          f"{counts.get(_STATUS_AUTO, 0)} auto-numbered, "
          f"{counts.get(_STATUS_WARN_NO_NUMBER, 0)} heading without number, "
          f"{counts.get(_STATUS_OK, 0)} ok")
    if counts.get(_STATUS_HARDCODED, 0):
        print("FAIL: leftover hardcoded prefixes remain")
        return 1
    print(f"PASS: no hardcoded prefixes; numbering is now automatic "
          f"(numId={result['num_id']})")
    return 0


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_style_levels(raw):
    """Parse --style-level "Heading1=0,Heading2=1" into a dict."""
    levels = {}
    if not raw:
        return levels
    for token in raw.split(","):
        token = token.strip()
        if not token:
            continue
        key, _, value = token.partition("=")
        key, value = key.strip(), value.strip()
        if not key or not value.isdigit():
            raise argparse.ArgumentTypeError(f"bad style-level: {token!r}")
        levels[key] = int(value)
    return levels


def parse_comma_list(raw):
    """Split a comma-separated option into stripped tokens."""
    return [token.strip() for token in raw.split(",") if token.strip()]


def parse_comma_ints(raw):
    """Split a comma-separated option into ints."""
    values = []
    for token in parse_comma_list(raw):
        if not token.isdigit():
            raise argparse.ArgumentTypeError(f"expected integer, got {token!r}")
        values.append(int(token))
    return values


def build_parser():
    ap = argparse.ArgumentParser(
        prog="numbering.py",
        description="Automate multilevel numbering in a DOCX (analyze / "
                    "define / convert).",
    )
    ap.add_argument("unpacked_dir", help="unpacked docx dir (has word/)")
    ap.add_argument("command", choices=["analyze", "define", "convert"],
                    help="subcommand")
    ap.add_argument("--style-level", default=None, metavar="Style1=0,Style2=1",
                    help="override default style->level map")
    ap.add_argument("--levels", type=int, default=3, metavar="N",
                    help="number of multilevel levels (default: 3)")
    ap.add_argument("--lvl-text", default=None, metavar="T1,T2,T3",
                    help="per-level lvlText, e.g. %%1.,%%1.%%2.,%%1.%%2.%%3. "
                         "(default: %%1,%%1.%%2,%%1.%%2.%%3)")
    ap.add_argument("--num-fmt", default="decimal", metavar="FMT",
                    help="number format for all levels (default: decimal)")
    ap.add_argument("--indent", default=None, metavar="L1,L2,L3",
                    help="per-level indent in twips (default: 0,0,0)")
    ap.add_argument("--list-num-fmt", default="decimal", metavar="FMT",
                    help="number format for body lists (default: decimal)")
    ap.add_argument("--list-lvl-text", default="%1", metavar="TEXT",
                    help="rendered text per body-list item, e.g. %%1), %%1. "
                         "(default: %%1)")
    return ap


def run_analyze(args):
    style_levels = dict(DEFAULT_STYLE_LEVELS)
    if args.style_level:
        style_levels.update(parse_style_levels(args.style_level))
    records = analyze_document(args.unpacked_dir, style_levels)
    return report_analyze(records, args.unpacked_dir, style_levels)


def run_define(args):
    if not 1 <= args.levels <= 9:
        raise SystemExit("error: --levels must be in 1..9")
    lvl_texts = (parse_comma_list(args.lvl_text)
                 if args.lvl_text else default_lvl_texts(args.levels))
    indents = (parse_comma_ints(args.indent)
               if args.indent else [0] * args.levels)
    if len(lvl_texts) != args.levels:
        raise SystemExit("error: --lvl-text must have exactly 'levels' entries "
                         f"({len(lvl_texts)} given for {args.levels})")
    if len(indents) != args.levels:
        raise SystemExit("error: --indent must have exactly 'levels' entries "
                         f"({len(indents)} given for {args.levels})")
    result = ensure_numbering(args.unpacked_dir, levels=args.levels,
                              lvl_texts=lvl_texts, num_fmt=args.num_fmt,
                              indents=indents)
    return report_define(result, args.unpacked_dir, args.levels,
                         lvl_texts, indents)


def run_convert(args):
    style_levels = dict(DEFAULT_STYLE_LEVELS)
    if args.style_level:
        style_levels.update(parse_style_levels(args.style_level))
    if not 1 <= args.levels <= 9:
        raise SystemExit("error: --levels must be in 1..9")
    lvl_texts = (parse_comma_list(args.lvl_text)
                 if args.lvl_text else default_lvl_texts(args.levels))
    indents = (parse_comma_ints(args.indent)
               if args.indent else [0] * args.levels)
    if len(lvl_texts) != args.levels:
        raise SystemExit("error: --lvl-text must have exactly 'levels' entries "
                         f"({len(lvl_texts)} given for {args.levels})")
    if len(indents) != args.levels:
        raise SystemExit("error: --indent must have exactly 'levels' entries "
                         f"({len(indents)} given for {args.levels})")
    result = convert_document(args.unpacked_dir, style_levels,
                              levels=args.levels, lvl_texts=lvl_texts,
                              num_fmt=args.num_fmt, indents=indents,
                              list_num_fmt=args.list_num_fmt,
                              list_lvl_text=args.list_lvl_text)
    return report_convert(result, args.unpacked_dir, style_levels)


def main(argv=None):
    args = build_parser().parse_args(argv)
    if args.command == "analyze":
        return run_analyze(args)
    if args.command == "define":
        return run_define(args)
    if args.command == "convert":
        return run_convert(args)
    return 2


if __name__ == "__main__":
    sys.exit(main())
