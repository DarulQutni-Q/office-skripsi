# SKILL: Skripsi/Thesis DOCX Editor & Consistency Auditor

Mengajarkan agent coding cara merevisi dokumen skripsi (.docx) secara langsung — edit isi, perbaiki Daftar Isi/Tabel/Gambar, cross-check sitasi, validasi format, dan uji konsistensi lintas bab — dengan target similarity <20%.

> **Auto-load**: Repo ini sudah punya `opencode.json` → `./SKILL.md`. Clone → `cd office-skripsi` → langsung ter-load otomatis di OpenCode.

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
