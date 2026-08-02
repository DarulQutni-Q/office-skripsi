---
name: skripsi
description: Edit skripsi/thesis DOCX — content, formatting, TOC, captions, citations, consistency checks, and plagiarism-safe paraphrasing. Use when the user mentions skripsi, thesis, Word document revision, DOCX editing, fixing headings/captions/TOC, or formatting skripsi.
license: MIT
compatibility: opencode
metadata:
  audience: students
  format: docx
---

# SKILL: Skripsi/Thesis DOCX Editor & Consistency Auditor

Mengajarkan agent coding cara merevisi dokumen skripsi (.docx) secara langsung — edit isi, perbaiki Daftar Isi/Tabel/Gambar, cross-check sitasi, validasi format, dan uji konsistensi lintas bab — dengan target similarity <20%.

> **Auto-load**: Repo ini sudah punya `opencode.json`. Clone → `cd office-skripsi` → skill siap pakai.

---

## 0. Dokumen Skripsi — yang Perlu Dipahami

Skripsi adalah **dokumen kompleks** dengan lapisan masalah:

| Lapisan | Contoh Masalah |
|---------|----------------|
| **Format** | Heading tanpa nomor, caption tidak urut, TOC basi, daftar pustaka tidak sinkron |
| **Konten** | Paragraf definisi terdeteksi plagiarisme, ketidaksesuaian angka/istilah antar bab |
| **Konsistensi** | Metode disebut di Bab 3 tapi tidak ada teorinya di Bab 2, state di kode 9 tapi di teks 8 |
| **Referensi** | Sitasi in-text tidak punya pasangan di daftar pustaka, nomor tabel di teks beda dengan realita |

Skill ini fokus ke **DOCX** (99% skripsi pakai Word). PPTX/PDF/XLSX adalah pendukung — lihat sub-skill masing-masing.

---

## 1. Instalasi

```bash
# Clone repo
git clone https://github.com/andypratama3/office-skripsi.git
cd office-skripsi

# Install dependensi Python
pip install python-docx lxml
```

**Dependensi sistem:**

| OS | LibreOffice | Poppler | zip/unzip |
|----|-------------|---------|-----------|
| **macOS** | `brew install --cask libreoffice` | `brew install poppler` | Bawaan |
| **Linux** | `sudo apt install libreoffice` | `sudo apt install poppler-utils` | `sudo apt install zip unzip` |
| **Windows** | Download dari [libreoffice.org](https://libreoffice.org) | [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases) | Bawaan (PowerShell `Expand-Archive`) |

Update: `cd office-skripsi && git pull`

---

## 2. Prasyarat

```bash
pip install python-docx lxml
```

Pastikan LibreOffice (`soffice`) dan Poppler (`pdftotext`) sudah terinstall dan masuk PATH.

---

## 3. Alur Kerja Wajib (10 Langkah — Jangan Lewatkan!)

```
 1. Unzip docx       → mkdir work && unzip -q original.docx -d unpacked/
 2. Backup           → cp original.docx original_backup.docx
 3. Merge runs       → python3 scripts/merge-runs.py unpacked/word/document.xml
 4. Baca & pahami    → Render PDF + pdftotext + python-docx inspect
 5. Analisis masalah → Identifikasi SEMUA masalah (heading, SEQ, sitasi, konsistensi)
 6. Edit XML         → str_replace bertarget dengan konteks 30-60 karakter
 7. Zip ulang        → cd unpacked && zip -Xr ../revised.docx .
 8. VALIDASI         → ZIP integrity + doc opens + XML well-formed + konten
 9. Render & cek     → soffice → PDF → pdftotext → grep ulang
10. Ulangi 6-9       → Sampai semua konsisten
```

> **Windows users**: Ganti `unzip` dengan `Expand-Archive`, `zip` dengan `Compress-Archive`, `cp` dengan `Copy-Item`. Gunakan `python` bukan `python3`. Jalankan LibreOffice dari path lengkap: `& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf file.docx`

---

## 4. Aturan Universal (Prioritas Tertinggi)

1. **DOCX adalah ZIP berisi XML** — jangan pernah edit .docx langsung. `unzip` dulu, baru edit `word/document.xml`.
2. **Backup SEBELUM edit** — `cp file.docx file_backup.docx`. Simpan di folder `work/`.
3. **Merge runs DAHULU** — Word memecah satu kalimat jadi banyak `<w:r>` kecil. Kalau tidak di-merge, str_replace akan gagal.
4. **Satu batch edit → validasi** — jangan edit 10 tempat lalu baru validasi. Susah lacak errornya.
5. **Render PDF & grep untuk verifikasi** — jangan percaya XML doang. `soffice` lalu `pdftotext` lalu `grep`.
6. **Field (TOC/SEQ/PAGEREF) TIDAK auto-update** — LibreOffice headless tidak bisa diandalkan. **Hitung manual, hardcode nilai cache**, baru set `updateFields=true` sebagai jaring pengaman.
7. **Setiap sisipan → efek berantai** — nambah 1 tabel di tengah menggeser nomor SEMUA tabel setelahnya + referensi teks di seluruh dokumen.
8. **Source code > screenshot > teks** — kalau kode punya 9 state tapi teks bilang 8, teks yang salah.

---

## 4A. DAFTAR GAMBAR/TABEL OTOMATIS (TOF) — Resep Terbukti

**Konteks**: mengubah DAFTAR GAMBAR/TABEL manual (ketik tangan, gampang basi) jadi field Word otomatis yang mengumpulkan caption. Resep ini **teruji di Microsoft Word for Mac 16.100.4** dan tidak tergantung LibreOffice.

### Fakta lapangan (jangan dilawan, pakai ini)

1. **Word for Mac 16.100.4 MENGABAIKAN switch `\t`, `\f` (TC fields), `\c`** pada field TOC — semuanya menghasilkan "No table of contents entries found." **Hanya `\o` dan `\u` yang bekerja.**
2. **outlineLvl bersifat 0-based**: `<w:outlineLvl w:val="6"/>` = TOC level 7, `w:val="7"` = TOC level 8.
3. **Pesan error mengidentifikasi tipe field**: "No table of **figures** entries found." = HANYA muncul dari field `TOC \c` (mis. daftar disisipkan ulang via ribbon References → Insert Table of Figures, atau file lewat Google Docs/WPS yang menulis ulang field). "No table of **contents** entries found." = field `\o`/`\u`. Jadi kalau user melapor "No table of figures" → field di file-nya sudah diganti, bukan file hasil kerja kita.
4. **`\c` butuh SEQ yang TERLIHAT**: SEQ dengan `\h` (hidden result) atau `w:vanish` TIDAK dihitung oleh `\c`; dan hasil SEQ yang vanished akan di-render ulang Word sebagai teks biasa (terlihat di caption). Artinya: tidak mungkin menyembunyikan nomor caption sambil memakai `\c`. Nomor caption manual ("Gambar 2.1") tidak cocok dengan pendekatan `\c`.

### Resep yang terbukti (outlineLvl + \o)

1. Style caption custom (mis. `GambarCaption`, `TabelCaption`), **non-italic**, beri `outlineLvl` di `styles.xml`: Gambar `val="6"`, Tabel `val="7"`.
2. Body caption dipaksa pakai style tsb (tanpa mengubah teks "Gambar 2.1 …" yang sudah ada).
3. Ganti isi DAFTAR GAMBAR/TABEL dengan field live:
   - Gambar: `TOC \h \z \o "7-7"`
   - Tabel: `TOC \h \z \o "8-8"`
   - Main TOC: `TOC \o "1-3" \h \z \u` — tidak terpengaruh.
4. Biarkan hasil field TIDAK ter-update di XML (atau update lewat Word dulu), lalu set `<w:updateFields w:val="true"/>` di `word/settings.xml` sebagai jaring pengaman.
5. Pastikan entry TOC di daftar **non-italic** di 4 lapis: runs, style TOC-level, style Hyperlink, dan `rPr` paragraph-mark.
6. Verifikasi WAJIB: buka di Word Mac → update semua field 2× → cek jumlah entri per field (harus = jumlah caption) → render PDF → cocokkan nomor halaman dengan referensi.

### Perintah untuk user (jangan lupa disampaikan)

- Update daftar: klik kanan daftar → **Update Field**, atau `Cmd+A` → `F9`.
- **JANGAN** hapus lalu sisipkan ulang via ribbon (References → Insert Table of Figures) — itu membuat field `\c` yang butuh SEQ → muncul "No table of figures entries found." dan daftar jadi kosong.
- **JANGAN** lewatkan file lewat Google Docs / WPS untuk "preview" — konversi menulis ulang field TOC custom dan menghancurkannya.

---

## 4B. FORMAT PEDOMAN (Margin/Font/Heading) — Resep Terbukti

**Konteks**: menyesuaikan format dokumen (font, margin, caption, footer) dengan pedoman kampus. Teruji di skripsi 126 halaman (replace Calibri → Times New Roman di seluruh dokumen).

### Spesifikasi pedoman TA SI (Ver 26.01) — sumber lokal

Pedoman PDF biasanya ada di `~/Downloads/Pedoman-Penulisan-*.pdf`; ekstrak dengan `pdftotext -layout` lalu grep section margin/font.

| Aturan | Nilai | Twips (1 cm = 567) |
|--------|-------|---------------------|
| Margin kiri/kanan | 3,5 cm | `w:left/right="1985"` |
| Margin atas | 3 cm | `w:top="1701"` |
| Margin bawah | 4 cm | `w:bottom="2268"` |
| Jarak nomor halaman footer | 1,5 cm dari tepi bawah | `w:footer="851"` (708 = 1,25 cm, SALAH) |
| Isi | TNR 12, spasi 1 (`w:line="240" w:lineRule="auto"`) | |
| BAB & judul bab | TNR 12, KAPITAL, tebal | |
| Caption gambar/tabel | TNR 10, bold, center, **tanpa titik akhir** | |

### Fakta lapangan (jangan dilawan)

1. **Word meng-STRIP edit style-level `rFonts` saat save** — patch `styles.xml` (Heading1–5 → TNR) tampak benar sampai Word di-save; style kembali ke font tema. **Yang bertahan: direct formatting run-level** (`<w:rPr><w:rFonts .../></w:rPr>` di tiap `<w:r>`) + `w:sz` + footer margin.
2. **Trap inheritansi tema**: style TANPA `rFonts` (mis. `NoSpacing`, dipakai sel tabel) mewarisi `docDefaults` `w:asciiTheme="minorHAnsi"` → di Word/PDF render **Calibri** meski XML tidak mengandung string "Calibri" sama sekali. Gejala: `grep Calibri` = 0 tapi PDF ber-Calibri → cek `docDefaults` + style tanpa rFonts.
3. **Style `TOC1`–`TOC9`** juga sering membawa `w:asciiTheme="minorHAnsi"` — patch juga saat konversi font dokumen.
4. **Word menulis lock file `~$*.docx`** — hapus sebelum menyalin deliverable.
5. **`pdftotext` glyph artifacts**: teks seperti "Descrip8on" / "9dak" = stand-in pdftotext untuk glyph ligature subset font, BUKAN teks rusak. "Arial" berteks kosong di pdftohtml = artifact hyperlink, aman.

### Alur fix font yang terbukti

1. Backup + unzip (`cp file.docx format_backup.docx`; `unzip -q fmt.docx -d fmt_work/`).
2. **Patch RUN-level** (bukan style-level) untuk semua run target; tambahkan rFonts TNR juga di style-nya (belt & suspenders untuk renderer non-Word seperti LibreOffice).
3. Rezip: `(cd fmt_work && zip -q -r -X ../fmt5.docx .)` — WAJIB dari dalam folder agar tidak ada prefix folder di dalam zip.
4. Validasi XML: `python3 -c "from lxml import etree; etree.parse('word/document.xml')"`.
5. **Verifikasi via Word, bukan LibreOffice**: AppleScript `update fields ×2 → save as format PDF → close saving no` → PDF dihasilkan dari docx yang BELUM di-save = bukti edit XML utuh. Pola `/tmp/update_nosave_export.applescript`.
6. **Font audit berbasis PDF**: `pdffonts out.pdf` (daftar font subset tersemat) + `pdftohtml -xml -i out.pdf` + script mapping `fontspec → teks` → cari font non-target yang punya glyph nyata (lihat `scripts/audit-pdf-fonts.py`).
7. Setelah ganti font, **nomor halaman berubah** (127→126, TOC 106→105) — selalu re-export & cek ulang halaman di DAFTAR ISI.
8. Copy ke deliverable, hapus `~$`, lapor ke user.

---

## 4C. WAJIB TANYA KE PENULIS — Ketentuan Penulisan Sebelum Edit Format

**JANGAN PERNAH mengasumsikan pedoman** (margin/font/spasi beda per kampus & per prodi, dan bisa berubah per angkatan). Tanya ke penulis dulu, lalu simpan jawabannya ke `work/pedoman.md` sebagai rujukan audit.

### Pertanyaan wajib (bisa dijadikan satu message/pertanyaan bertingkat)

**A. Pedoman resmi**
1. Punya file pedoman resmi dari prodi/kampus? Versi & tahun berapa? (minta PDF-nya — sumber utama, bukan ingatan penulis)
2. Ada template `.docx` resmi prodi? (lebih baik daripada bikin dari nol)

**B. Format halaman & font**
3. Ukuran kertas (A4 / kuarto 21×33 cm / lainnya)?
4. Margin: kiri / kanan / atas / bawah (cm)?
5. Font isi & ukuran? Font judul bab / heading sub-bab?
6. Spasi baris isi (1 / 1,5 / 2)? Spasi setelah paragraf?
7. Font caption gambar/tabel (nama, ukuran, bold/tidak)? Posisi caption (di atas/bawah)? Penomoran caption ("Gambar 2.1" / "Gambar 2.1.1" / berurutan)? Akhiran titik pada caption (boleh/tidak)?

**C. Penomoran halaman**
8. Nomor halaman: posisi (atas/bawah, kiri/tengah/kanan)? Format halaman depan (romawi kecil) vs isi (angka)? Mulai nomor dari halaman mana?

**D. Daftar isi / daftar gambar / daftar tabel**
9. DAFTAR ISI / GAMBAR / TABEL: manual (ketik) atau otomatis (field)? Jumlah tingkat heading yang masuk (sampai 1.1.1? 1.1.1.1?)

**E. Sitasi & daftar pustaka**
10. Gaya sitasi (IEEE / APA / Vancouver / Numeric / MLA / lainnya)? Format daftar pustaka (spasi, indentasi gantung)?
11. Aturan sitasi tabel/gambar yang dikutip dari sumber (wajib dicantumkan "Sumber: ..." atau tidak)?

**F. Struktur dokumen**
12. Struktur wajib tiap bab (ada template per-bab dari prodi)? Lampiran wajib?
13. Penomoran bab (BAB I, BAB II / Bab 1, Bab 2) & penomoran sub-bab (1.1, 1.1.1)?
14. Judul bab: KAPITAL penuh / Title Case? Posisi (kiri/tengah)? Garis bawah?
15. Bahasa & ejaan (EYD/PUEBI)? Ada kata/istilah khusus yang wajib dipakai?
16. Batas similarity/plagiarisme (mis. Turnitin <20% atau <30%)?

### Aturan pakai

- Kalau penulis jawab "ikuti pedoman" tanpa memberi file → minta file pedoman-nya; kalau tidak ada, jawab default yang sudah diketahui (mis. TNR 12, spasi 1) tetapi **tandai sebagai asumsi** dan minta konfirmasi.
- Jawaban disimpan ke `work/pedoman.md` + dipakai sebagai checklist audit sebelum klaim "sesuai pedoman".
- Kalau ada konflik antara jawaban penulis dan isi dokumen → yang menang pedoman resmi; konfirmasi ke penulis kalau ragu.
- Minta nama file pedoman + letakkan salinannya di `work/` agar bisa di-audit ulang (contoh: `~/Downloads/Pedoman-Penulisan-Tugas-Akhir-SI-Maret 2026-ttd.pdf` → `pdftotext -layout` → grep section margin/font).

---

## 5. Teknik Detail (DOCX Sub-Skill)

Semua teknik detail ada di `skills/docx/SKILL.md`:

| Butuh | Buka |
|-------|------|
| Merge runs & edit XML | `skills/docx/SKILL.md` §2 |
| SEQ / TOC / PAGEREF | `skills/docx/SKILL.md` §3 |
| Checklist isi skripsi (teori vs implementasi, sitasi, riwayat hidup, kode) | `skills/docx/SKILL.md` §4 |
| Panduan menulis orisinal (anti-plagiarisme) | `skills/docx/SKILL.md` §5 |
| python-docx high-level API | `skills/docx/SKILL.md` §6 |
| Tabel masalah umum & solusi | `skills/docx/SKILL.md` §7 |
| Validasi akhir & checklist | `skills/docx/SKILL.md` §8 |
| Lessons learned (pengalaman nyata) | `skills/docx/SKILL.md` §9 |

**Sub-skill pendukung:**

| Skill | File | Untuk |
|-------|------|-------|
| PPTX | `skills/pptx/SKILL.md` | Edit slide presentasi sidang |
| PDF | `skills/pdf/SKILL.md` | Ekstraksi teks, konversi, verifikasi |
| Spreadsheet | `skills/spreadsheet/SKILL.md` | Olah data XLSX |

---

## 6. Utility Scripts

Semua script Python — **cross-platform** (macOS, Windows, Linux).

```bash
# Validasi dokumen setelah edit
python3 scripts/validate-docx.py revised.docx

# Merge runs di XML (wajib sebelum str_replace)
python3 scripts/merge-runs.py unpacked/word/document.xml

# Analisis isi: headings, captions, tables, citations
python3 scripts/docx-tools.py skripsi.docx all
python3 scripts/docx-tools.py skripsi.docx headings
python3 scripts/docx-tools.py skripsi.docx citations

# Audit font PDF (deteksi sisa Calibri/theme-font setelah konversi font)
pdftohtml -xml -i fmt.pdf /tmp/fmt.xml && python3 scripts/audit-pdf-fonts.py /tmp/fmt.xml --target Times
```

> Di Windows ganti `python3` dengan `python`. Pastikan Python terdaftar di PATH.

---

## 7. Anti-Patterns

| ❌ Jangan | ✅ Lakukan |
|-----------|-----------|
| Edit .docx langsung | `unzip` dulu, edit XML |
| Replace string pendek tanpa konteks | Ambil 30-60 karakter sekitar target |
| Skip merge runs | Jalankan `merge-runs.py` dulu |
| Skip backup | `cp file.docx file_backup.docx` |
| Langsung klaim "beres" tanpa render | `soffice → PDF → pdftotext → grep` |
| Andalkan auto-update TOC/SEQ | Hardcode manual, set updateFields safety |
| Edit 10 tempat lalu validasi 1x | Satu batch → validasi → ulangi |
| Tebak data ambigu (riwayat hidup, dll) | Tanya penulis — jangan asal hapus/ganti |
| Abaikan error XML | Satu tag tidak tertutup → dokumen corrupt |
