#!/usr/bin/env python3
"""
italic-foreign — miringkan (italic) HANYA istilah bahasa asing di skripsi .docx.

Aturan penulisan (PUEBI): kata/istilah asing dicetak miring; teks Indonesia
biasa, nama produk/merek (Laravel, MySQL), akronim (API, UML), caption,
Daftar Isi, dan judul bab TIDAK dimiringkan. Script ini menerapkan aturan itu
secara aman di level XML:

  * mencocokkan istilah di level PARAGRAF, jadi "machine " + "learning" yang
    terpecah ke beberapa <w:r> tetap terdeteksi (tidak wajib merge-runs dulu);
  * memecah <w:r> persis di batas istilah, format lain (bold, font, ukuran,
    warna, rsid) tetap dipertahankan;
  * mengisi <w:i> + <w:iCs> di urutan yang benar di dalam <w:rPr>;
  * melewati caption/SEQ, TOC/PAGEREF, heading, Daftar Pustaka, dan run
    ber-font monospace (kode);
  * idempotent: dijalankan dua kali hasilnya sama.

Usage:
    python italic-foreign.py audit PATH            # read-only, exit 1 kalau ada yang perlu diperbaiki
    python italic-foreign.py apply PATH [-o OUT]   # terapkan perubahan

PATH boleh berupa:
    * file .docx           -> hasil ditulis ke OUT (default: <nama>.italic.docx), asli tidak disentuh
    * folder hasil unzip   -> word/document.xml (+footnotes/endnotes) diedit di tempat
    * file document.xml    -> diedit di tempat (styles.xml dibaca dari folder yang sama bila ada)

Opsi penting:
    --strip-other        hapus italic dari teks NON-asing (perbaiki "seluruh kalimat miring")
    --terms FILE         tambah/kurangi istilah (satu per baris, '#' komentar,
                         awalan '!' = jangan miringkan istilah bawaan itu)
    --term "x y"         tambah satu istilah dari command line (boleh diulang)
    --no-builtin         pakai hanya istilah dari --terms/--term
    --include-headings   ikut proses heading (default dilewati, lihat catatan di bawah)
    --include-bibliography  ikut proses Daftar Pustaka (default dilewati)
    --dry-run            (apply) laporkan saja, jangan tulis apa pun

Catatan heading: Word menyalin format karakter heading ke entri Daftar Isi saat
field di-update. Italic di heading akan "bocor" ke TOC yang seharusnya non-italic,
karena itu heading dilewati secara default.
"""
import argparse
import copy
import io
import os
import re
import shutil
import sys
import zipfile
from collections import Counter

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"

# Pengganti karakter untuk isi non-teks (tab, br, field, run kode). Karakter
# private-use: bukan whitespace dan bukan \w, jadi istilah tidak bisa "melompati"
# tab/field, tetapi tetap dianggap batas kata.
SEP = "\ue000"


def q(tag):
    return "{%s}%s" % (W_NS, tag)


P, R, T, RPR = q("p"), q("r"), q("t"), q("rPr")

# Urutan anak <w:rPr> menurut skema (CT_RPr). Word bisa menolak/menulis ulang
# bila urutan salah, jadi elemen baru disisipkan di posisi yang benar.
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
]

# --------------------------------------------------------------------------
# Daftar istilah bawaan. Ini TITIK AWAL, bukan kebenaran mutlak: tiap kampus
# beda sikap soal kata serapan. Tambah lewat --terms, kurangi dengan '!istilah'.
# Sengaja TIDAK memuat nama produk/bahasa pemrograman/library (Laravel, React,
# MySQL, Python) dan akronim (API, REST, CRUD, UML) karena tidak dimiringkan.
# --------------------------------------------------------------------------
BUILTIN_TERMS = """
# --- rekayasa perangkat lunak / web ---
backend
back-end
back end
frontend
front-end
front end
full-stack
full stack
framework
middleware
endpoint
payload
webhook
callback
query
queries
cache
caching
load balancer
deployment
deploy
rollback
hotfix
commit
pull request
merge request
repository
staging
container
microservice
monolith
monolithic
headless
boilerplate
refactoring
refactor
bug
debugging
debug
pipeline
workflow
dashboard
login
logout
sign in
sign up
upload
download
plugin
dropdown
checkbox
sidebar
navbar
wireframe
mockup
flowchart
user interface
user experience
user story
sprint
backlog
stakeholder
requirement
requirements
real-time
real time
realtime
scheduler
cron job
background job
job queue
queue
multi-tenant
multitenant
single sign-on
role-based access control
access token
refresh token
bearer token
token
hashing
hash
open source
open-source
best practice
best practices
trade-off
tradeoff
state of the art
proof of concept
use case
class diagram
sequence diagram
activity diagram
entity relationship diagram
unit test
integration test
end-to-end
black box
white box
black-box
white-box
test case
feedback
online
offline
website
browser
software
hardware
# --- data / AI ---
machine learning
deep learning
big data
dataset
data set
overfitting
underfitting
fine-tuning
fine tuning
prompt engineering
prompt
embedding
embeddings
chatbot
natural language processing
computer vision
feature extraction
cross validation
cross-validation
confusion matrix
epoch
learning rate
hyperparameter
hyperparameters
preprocessing
pre-processing
data cleaning
data mining
text mining
web scraping
scraping
crawling
crawler
web crawler
outlier
missing value
missing values
one-hot encoding
label encoding
stopword
stopwords
stemming
tokenization
lemmatization
# --- Latin / ungkapan asing umum di tulisan ilmiah ---
et al.
et al
in vitro
in vivo
in situ
a priori
a posteriori
de facto
de jure
ad hoc
per se
vice versa
ibid.
op. cit.
loc. cit.
e.g.
i.e.
""".strip().splitlines()

MONO_FONT_RE = re.compile(
    r"courier|consolas|menlo|monaco|lucida console|source code|fira (?:code|mono)"
    r"|dejavu sans mono|cascadia|monospace|jetbrains mono",
    re.I,
)
TOC_STYLE_RE = re.compile(r"^toc|\btoc\b|tableoffigures|table of figures|tableofauthorities", re.I)
CAPTION_STYLE_RE = re.compile(r"caption", re.I)
HEADING_STYLE_RE = re.compile(r"^heading|^judul|^title$", re.I)
BIB_START_RE = re.compile(
    r"^\s*(DAFTAR\s+PUSTAKA|DAFTAR\s+REFERENSI|REFERENCES|BIBLIOGRAPHY)\s*$", re.I
)
BIB_END_RE = re.compile(r"^\s*(LAMPIRAN|RIWAYAT\s+HIDUP|BIODATA|BAB\s+[IVXL\d]+)\b", re.I)
FIELD_TOC_RE = re.compile(r"(?:^|\s)(?:TOC|PAGEREF)\b")
FIELD_SEQ_RE = re.compile(r"(?:^|\s)SEQ\b")


# --------------------------------------------------------------------------
# Istilah
# --------------------------------------------------------------------------
def _norm(term):
    return " ".join(term.strip().lower().split())


def load_terms(builtin=True, files=(), extra=()):
    """Gabungkan istilah bawaan + file + --term. Baris '!x' membuang x."""
    terms = set()
    excluded = set()
    lines = []
    if builtin:
        lines.extend(BUILTIN_TERMS)
    for path in files:
        with open(path, "r", encoding="utf-8") as fh:
            lines.extend(fh.read().splitlines())
    lines.extend(extra)
    for raw in lines:
        line = raw.split("#", 1)[0].strip()  # '#' = komentar (awal atau akhir baris)
        if not line:
            continue
        if line.startswith("!"):
            excluded.add(_norm(line[1:]))
        else:
            terms.add(_norm(line))
    return terms - excluded


def build_regex(terms):
    if not terms:
        return None
    alts = []
    for term in sorted(terms, key=lambda t: (-len(t), t)):
        body = r"\s+".join(re.escape(w) for w in term.split())
        if term[-1].isalpha():
            body += r"(?:e?s)?"  # bentuk jamak Inggris: requirements, branches
        alts.append(body)
    return re.compile(r"(?<!\w)(?:" + "|".join(alts) + r")(?!\w)", re.IGNORECASE)


# --------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------
class StyleIndex:
    """Index ringan styles.xml: nama, basedOn, italic, outline level."""

    def __init__(self, styles_root=None):
        self.styles = {}
        if styles_root is None:
            return
        for st in styles_root.iter(q("style")):
            sid = st.get(q("styleId"))
            if not sid:
                continue
            name = st.find(q("name"))
            based = st.find(q("basedOn"))
            outline = st.find("%s/%s" % (q("pPr"), q("outlineLvl")))
            self.styles[sid] = {
                "name": name.get(q("val")) if name is not None else "",
                "based": based.get(q("val")) if based is not None else None,
                "italic": _flag_state(st.find(q("rPr")), "i"),
                "outline": int(outline.get(q("val"))) if outline is not None
                and (outline.get(q("val")) or "").isdigit() else None,
            }

    def name(self, sid):
        return self.styles.get(sid, {}).get("name", "")

    def _resolve(self, sid, key):
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            val = self.styles[sid][key]
            if val is not None:
                return val
            sid = self.styles[sid]["based"]
        return None

    def italic_of(self, sid):
        return bool(self._resolve(sid, "italic"))

    def outline_of(self, sid):
        return self._resolve(sid, "outline")


def _flag_state(rpr, tag):
    """None = tidak diset; True/False = nilai eksplisit (val 0/false/off = False)."""
    if rpr is None:
        return None
    el = rpr.find(q(tag))
    if el is None:
        return None
    return el.get(q("val")) not in ("0", "false", "off")


def baseline_italic(run, para_style_italic, styles):
    """Italic dari style saja (tanpa format langsung). Paragraph style dan
    character style adalah toggle property: keduanya aktif = saling membatalkan."""
    rpr = run.find(RPR)
    char_italic = False
    if rpr is not None:
        rs = rpr.find(q("rStyle"))
        if rs is not None:
            char_italic = styles.italic_of(rs.get(q("val")))
    return bool(para_style_italic) != bool(char_italic)


def effective_italic(run, baseline):
    direct = _flag_state(run.find(RPR), "i")
    return baseline if direct is None else direct


# --------------------------------------------------------------------------
# Manipulasi rPr
# --------------------------------------------------------------------------
def _local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else None


def _set_flag(rpr, tag, on):
    for old in rpr.findall(q(tag)):
        rpr.remove(old)
    el = etree.SubElement(rpr, q(tag))
    if not on:
        el.set(q("val"), "0")
    rpr.remove(el)
    idx = RPR_ORDER.index(tag)
    pos = len(rpr)
    for i, child in enumerate(rpr):
        name = _local(child)
        order = RPR_ORDER.index(name) if name in RPR_ORDER else len(RPR_ORDER)
        if order > idx:
            pos = i
            break
    rpr.insert(pos, el)


def set_italic(run, want, baseline):
    """Atur italic eksplisit pada satu run."""
    rpr = run.find(RPR)
    if want:
        if rpr is None:
            rpr = etree.SubElement(run, RPR)
            run.remove(rpr)
            run.insert(0, rpr)
        _set_flag(rpr, "i", True)
        _set_flag(rpr, "iCs", True)
        return
    if baseline:
        # style-nya italic: harus dimatikan eksplisit, bukan sekadar dihapus
        if rpr is None:
            rpr = etree.SubElement(run, RPR)
            run.remove(rpr)
            run.insert(0, rpr)
        _set_flag(rpr, "i", False)
        _set_flag(rpr, "iCs", False)
    elif rpr is not None:
        for tag in ("i", "iCs"):
            for old in rpr.findall(q(tag)):
                rpr.remove(old)
        if len(rpr) == 0:
            run.remove(rpr)


def is_mono(run):
    rpr = run.find(RPR)
    if rpr is None:
        return False
    fonts = rpr.find(q("rFonts"))
    if fonts is None:
        return False
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        val = fonts.get(q(attr))
        if val and MONO_FONT_RE.search(val):
            return True
    return False


# --------------------------------------------------------------------------
# Paragraf
# --------------------------------------------------------------------------
def nearest_para(el):
    a = el.getparent()
    while a is not None and a.tag != P:
        a = a.getparent()
    return a


def para_runs(p):
    """Run milik paragraf ini saja (bukan paragraf bersarang di textbox, bukan math)."""
    out = []
    for r in p.iter(R):
        if nearest_para(r) is not p:
            continue
        if any(a.tag.startswith("{%s}" % M_NS) for a in r.iterancestors()):
            continue
        out.append(r)
    return out


def para_text(p):
    return "".join((t.text or "") for t in p.iter(T) if nearest_para(t) is p)


def classify(p, styles):
    """'toc' | 'caption' | 'heading' | 'body'."""
    instr = []
    for it in p.iter(q("instrText")):
        if nearest_para(it) is p:
            instr.append(it.text or "")
    for fs in p.iter(q("fldSimple")):
        if nearest_para(fs) is p:
            instr.append(fs.get(q("instr")) or "")
    joined = " ".join(instr)

    ppr = p.find(q("pPr"))
    sid = ""
    outline = None
    if ppr is not None:
        ps = ppr.find(q("pStyle"))
        if ps is not None:
            sid = ps.get(q("val")) or ""
        ol = ppr.find(q("outlineLvl"))
        if ol is not None and (ol.get(q("val")) or "").isdigit():
            outline = int(ol.get(q("val")))
    label = "%s %s" % (sid, styles.name(sid))

    if TOC_STYLE_RE.search(sid) or TOC_STYLE_RE.search(styles.name(sid)) or FIELD_TOC_RE.search(joined):
        return "toc"
    if CAPTION_STYLE_RE.search(label) or FIELD_SEQ_RE.search(joined):
        return "caption"
    if outline is None:
        outline = styles.outline_of(sid)
    if HEADING_STYLE_RE.search(sid) or HEADING_STYLE_RE.search(styles.name(sid)) \
            or (outline is not None and 0 <= outline <= 5):
        return "heading"
    return "body"


# --------------------------------------------------------------------------
# Laporan
# --------------------------------------------------------------------------
class Report:
    def __init__(self):
        self.to_italicize = []   # (para_no, matched, context)
        self.already = 0         # istilah yang sudah miring
        self.to_strip = []       # (para_no, teks miring non-asing)
        self.structural = []     # (para_no, kind, teks) italic di TOC/caption
        self.skipped = Counter()
        self.terms = Counter()
        self.changed_runs = 0

    @property
    def pending(self):
        return len(self.to_italicize) + len(self.to_strip)

    def merge(self, other):
        self.to_italicize += other.to_italicize
        self.already += other.already
        self.to_strip += other.to_strip
        self.structural += other.structural
        self.skipped.update(other.skipped)
        self.terms.update(other.terms)
        self.changed_runs += other.changed_runs


# --------------------------------------------------------------------------
# Inti: proses satu paragraf
# --------------------------------------------------------------------------
def _split_flags(text, flags):
    segs, i = [], 0
    while i < len(text):
        f, j = flags[i], i
        while j < len(text) and flags[j] == f:
            j += 1
        segs.append((bool(f), text[i:j]))
        i = j
    return segs


def _make_t(orig_t, piece):
    t = copy.deepcopy(orig_t)
    t.text = piece
    t.tail = None
    if piece != piece.strip() or "  " in piece:
        t.set("{%s}space" % XML_NS, "preserve")
    return t


def _rebuild_run(run, units, baseline):
    """units: [(kind, payload, target)] kind 't' (payload=(orig_t, piece)) / 'o' (elem).
    Mengembalikan True bila ada perubahan."""
    groups = []
    for u in units:
        if groups and groups[-1][0] == u[2]:
            groups[-1][1].append(u)
        else:
            groups.append([u[2], [u]])
    if len(groups) == 1:
        target = groups[0][0]
        if target is None:
            return False
        set_italic(run, target, baseline)  # seluruh run sama -> ubah di tempat
        return True

    parent = run.getparent()
    pos = parent.index(run)
    tail = run.tail
    new_runs = []
    for target, us in groups:
        nr = copy.deepcopy(run)
        nr.tail = None
        for ch in list(nr):
            if ch.tag != RPR:
                nr.remove(ch)
        for kind, payload, _ in us:
            if kind == "t":
                nr.append(_make_t(*payload))
            else:
                nr.append(copy.deepcopy(payload))
        if target is not None:
            set_italic(nr, target, baseline)
        new_runs.append(nr)
    parent.remove(run)
    for i, nr in enumerate(new_runs):
        parent.insert(pos + i, nr)
    new_runs[-1].tail = tail
    return True


def process_paragraph(p, para_no, ctx, report):
    runs = para_runs(p)
    if not runs:
        return
    ppr = p.find(q("pPr"))
    ps = ppr.find(q("pStyle")) if ppr is not None else None
    para_style_italic = ctx.styles.italic_of(ps.get(q("val")) if ps is not None else None)

    text_parts = []
    entries = []   # (run, [(t_elem, offset)], mono, baseline)
    italic_at = []  # per karakter: miring efektif?
    offset = 0
    for run in runs:
        base = baseline_italic(run, para_style_italic, ctx.styles)
        eff = effective_italic(run, base)
        mono = is_mono(run)
        ts = []
        for child in run:
            if child.tag == RPR:
                continue
            if child.tag == T:
                txt = child.text or ""
                text_parts.append(SEP * len(txt) if mono else txt)
                italic_at.extend([eff] * len(txt))
                ts.append((child, offset))
                offset += len(txt)
            else:
                text_parts.append(SEP)
                italic_at.append(False)
                offset += 1
        entries.append((run, ts, mono, base, eff))
    text = "".join(text_parts)

    flags = bytearray(len(text))
    if ctx.regex is not None:
        for m in ctx.regex.finditer(text):
            s, e = m.span()
            for i in range(s, e):
                flags[i] = 1
            report.terms[_norm(m.group(0))] += 1
            if all(italic_at[s:e]):
                report.already += 1
            else:
                lo, hi = max(0, s - 25), min(len(text), e + 25)
                report.to_italicize.append(
                    (para_no, m.group(0), text[lo:hi].replace(SEP, " ").strip()))

    if ctx.strip_other:
        i, n = 0, len(text)
        while i < n:
            if italic_at[i] and not flags[i] and text[i] != SEP:
                j = i
                while j < n and italic_at[j] and not flags[j] and text[j] != SEP:
                    j += 1
                snippet = text[i:j].strip()
                if snippet:
                    report.to_strip.append((para_no, snippet))
                i = j
            else:
                i += 1

    if not ctx.apply:
        return

    for run, ts, mono, base, eff in entries:
        if mono or not ts:
            continue
        units = []
        for child in run:
            if child.tag == RPR:
                continue
            if child.tag != T:
                units.append(("o", child, None))
                continue
            off = next(o for (t, o) in ts if t is child)
            txt = child.text or ""
            for is_foreign, piece in _split_flags(txt, flags[off:off + len(txt)]):
                if is_foreign:
                    target = True if not eff else None
                elif ctx.strip_other and eff:
                    target = False
                else:
                    target = None
                units.append(("t", (child, piece), target))
        if _rebuild_run(run, units, base):
            report.changed_runs += 1


class Ctx:
    def __init__(self, regex, styles, apply, strip_other, include_headings, include_bib):
        self.regex = regex
        self.styles = styles
        self.apply = apply
        self.strip_other = strip_other
        self.include_headings = include_headings
        self.include_bib = include_bib


def process_tree(root, ctx):
    report = Report()
    in_bib = False
    for no, p in enumerate(list(root.iter(P)), 1):
        kind = classify(p, ctx.styles)
        plain = para_text(p).strip()
        if kind != "toc" and BIB_START_RE.match(plain):
            in_bib = True
        elif in_bib and len(plain) <= 80 and BIB_END_RE.match(plain) \
                and (kind == "heading" or plain.upper() == plain):
            in_bib = False

        reason = None
        if kind in ("toc", "caption"):
            reason = kind
        elif kind == "heading" and not ctx.include_headings:
            reason = "heading"
        elif in_bib and not ctx.include_bib:
            reason = "bibliography"

        if reason:
            report.skipped[reason] += 1
            if kind in ("toc", "caption"):
                ppr = p.find(q("pPr"))
                ps = ppr.find(q("pStyle")) if ppr is not None else None
                psi = ctx.styles.italic_of(ps.get(q("val")) if ps is not None else None)
                for r in para_runs(p):
                    txt = "".join(t.text or "" for t in r.findall(T)).strip()
                    if txt and effective_italic(r, baseline_italic(r, psi, ctx.styles)):
                        report.structural.append((no, kind, txt))
            continue
        process_paragraph(p, no, ctx, report)
    return report


# --------------------------------------------------------------------------
# I/O
# --------------------------------------------------------------------------
PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")


def _parse(data):
    parser = etree.XMLParser(resolve_entities=False, huge_tree=True, remove_blank_text=False)
    return etree.fromstring(data, parser)


def _serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _styles_from_bytes(data):
    return StyleIndex(_parse(data)) if data else StyleIndex()


def run_on_docx(path, out, ctx_factory, write):
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        styles = _styles_from_bytes(zin.read("word/styles.xml") if "word/styles.xml" in names else None)
        ctx = ctx_factory(styles)
        total = Report()
        new_parts = {}
        for part in PARTS:
            if part not in names:
                continue
            root = _parse(zin.read(part))
            total.merge(process_tree(root, ctx))
            if ctx.apply:
                new_parts[part] = _serialize(root)
        if write and ctx.apply:
            tmp = out + ".tmp"
            with zipfile.ZipFile(tmp, "w") as zout:
                for info in zin.infolist():
                    data = new_parts.get(info.filename)
                    if data is None:
                        data = zin.read(info.filename)
                    zi = zipfile.ZipInfo(info.filename, info.date_time)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zi.external_attr = info.external_attr
                    zout.writestr(zi, data)
            os.replace(tmp, out)
    return total


def run_on_dir(folder, ctx_factory, write):
    spath = os.path.join(folder, "word", "styles.xml")
    styles = StyleIndex()
    if os.path.exists(spath):
        with open(spath, "rb") as fh:
            styles = _styles_from_bytes(fh.read())
    ctx = ctx_factory(styles)
    total = Report()
    for part in PARTS:
        fp = os.path.join(folder, *part.split("/"))
        if not os.path.exists(fp):
            continue
        with open(fp, "rb") as fh:
            root = _parse(fh.read())
        total.merge(process_tree(root, ctx))
        if write and ctx.apply:
            with open(fp, "wb") as fh:
                fh.write(_serialize(root))
    return total


def run_on_xml(path, ctx_factory, write):
    sibling = os.path.join(os.path.dirname(os.path.abspath(path)), "styles.xml")
    styles = StyleIndex()
    if os.path.exists(sibling):
        with open(sibling, "rb") as fh:
            styles = _styles_from_bytes(fh.read())
    ctx = ctx_factory(styles)
    with open(path, "rb") as fh:
        root = _parse(fh.read())
    report = process_tree(root, ctx)
    if write and ctx.apply:
        with open(path, "wb") as fh:
            fh.write(_serialize(root))
    return report


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def print_report(rep, mode, limit):
    def clip(s, n=90):
        return s if len(s) <= n else s[: n - 1] + "…"

    print("=== LAPORAN ITALIC ISTILAH ASING ===")
    print("Istilah asing belum miring : %d" % len(rep.to_italicize))
    print("Istilah asing sudah miring : %d" % rep.already)
    if rep.to_strip or mode.get("strip_other"):
        print("Teks miring non-asing      : %d" % len(rep.to_strip))
    print("Paragraf dilewati          : %s" % (
        ", ".join("%s=%d" % kv for kv in sorted(rep.skipped.items())) or "-"))
    if rep.to_italicize:
        print("\nAkan dimiringkan:")
        for no, term, ctx_txt in rep.to_italicize[:limit]:
            print("  ¶%-5d [%s] …%s…" % (no, term, clip(ctx_txt)))
        if len(rep.to_italicize) > limit:
            print("  ... %d lagi (--report-limit untuk menambah)" % (len(rep.to_italicize) - limit))
    if rep.to_strip:
        print("\nTeks miring yang BUKAN istilah asing (akan di-non-italic-kan bila --strip-other):")
        print("  Kalau ada yang sebenarnya istilah asing, tambahkan ke --terms agar tetap miring.")
        for no, snippet in rep.to_strip[:limit]:
            print("  ¶%-5d %s" % (no, clip(snippet)))
        if len(rep.to_strip) > limit:
            print("  ... %d lagi" % (len(rep.to_strip) - limit))
    if rep.structural:
        print("\nPERINGATAN: italic di Daftar Isi/caption (aturan repo: non-italic). Tidak diubah otomatis:")
        for no, kind, txt in rep.structural[:limit]:
            print("  ¶%-5d [%s] %s" % (no, kind, clip(txt)))
    if rep.terms:
        print("\nIstilah terdeteksi (untuk ditinjau; kurangi dengan '!istilah' di --terms):")
        print("  " + ", ".join("%s×%d" % kv for kv in rep.terms.most_common(25)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Miringkan hanya istilah asing di skripsi .docx",
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["audit", "apply"])
    ap.add_argument("path", help=".docx, folder hasil unzip, atau word/document.xml")
    ap.add_argument("-o", "--output", help="file keluaran untuk input .docx")
    ap.add_argument("--terms", action="append", default=[], help="file istilah tambahan/pengecualian")
    ap.add_argument("--term", action="append", default=[], help="istilah tambahan (boleh diulang)")
    ap.add_argument("--no-builtin", action="store_true", help="jangan pakai daftar bawaan")
    ap.add_argument("--strip-other", action="store_true", help="hapus italic dari teks non-asing")
    ap.add_argument("--include-headings", action="store_true")
    ap.add_argument("--include-bibliography", action="store_true")
    ap.add_argument("--dry-run", action="store_true", help="apply tanpa menulis apa pun")
    ap.add_argument("--report-limit", type=int, default=40)
    args = ap.parse_args(argv)

    # konsol Windows lama (cp437/cp1252) tidak bisa mencetak semua simbol laporan
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")

    if not os.path.exists(args.path):
        print("Tidak ditemukan: %s" % args.path, file=sys.stderr)
        return 2
    terms = load_terms(builtin=not args.no_builtin, files=args.terms, extra=args.term)
    regex = build_regex(terms)
    if regex is None and not args.strip_other:
        print("Daftar istilah kosong.", file=sys.stderr)
        return 2

    do_apply = args.command == "apply"
    write = do_apply and not args.dry_run

    def factory(styles):
        return Ctx(regex, styles, do_apply, args.strip_other,
                   args.include_headings, args.include_bibliography)

    out = None
    if os.path.isdir(args.path):
        rep = run_on_dir(args.path, factory, write)
    elif args.path.lower().endswith(".docx"):
        out = args.output or re.sub(r"\.docx$", "", args.path, flags=re.I) + ".italic.docx"
        if os.path.abspath(out) == os.path.abspath(args.path):
            print("Output tidak boleh sama dengan input (.docx asli tidak ditimpa).", file=sys.stderr)
            return 2
        rep = run_on_docx(args.path, out, factory, write)
    else:
        rep = run_on_xml(args.path, factory, write)

    print_report(rep, {"strip_other": args.strip_other}, args.report_limit)
    if not do_apply:
        print("\nAUDIT: %s" % ("ada %d perubahan yang diperlukan" % rep.pending if rep.pending
                               else "tidak ada yang perlu diubah"))
        return 1 if rep.pending else 0
    if args.dry_run:
        print("\nDRY-RUN: tidak ada yang ditulis (%d run akan berubah)." % rep.changed_runs)
        return 0
    print("\nSELESAI: %d run diubah%s." % (rep.changed_runs, (" -> %s" % out) if out else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
