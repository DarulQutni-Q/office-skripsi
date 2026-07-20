# Skripsi Skill — AI Agent untuk Revisi Skripsi/Thesis DOCX

Skill pack untuk AI coding agents (OpenCode, Claude Code, dll) yang mengajarkan cara merevisi dokumen skripsi (.docx) secara programatik — edit konten, perbaiki format, sinkronisasi Daftar Isi/Tabel/Gambar, cross-check sitasi, validasi konsistensi lintas bab, dan menjaga orisinalitas tulisan.

## Untuk Siapa

- **Mahasiswa** yang ingin merevisi skripsinya dengan bantuan AI agent
- **AI coding agents** yang perlu tahu cara edit .docx tanpa Microsoft Office
- Siapapun yang bekerja dengan dokumen skripsi berbahasa Indonesia

## Instalasi

### Opsi 1: Clone (Rekomendasi)

```bash
git clone https://github.com/andypratama3/office-skill.git
cd office-skill
pip install python-docx lxml
brew install --cask libreoffice poppler
```

### Opsi 2: npx dari GitHub (tanpa clone manual)

```bash
npx github:andypratama3/office-skill
```

Tapi untuk benar-benar pakai skill-nya, tetap perlu clone — npx cuma nampilin info.

### Opsi 3: Download ZIP

Download dari https://github.com/andypratama3/office-skill, extract, lalu jalankan:

```bash
cd office-skill-main
pip install python-docx lxml
```

## Cara Pakai

### Sebagai AI Agent Skill

Di konfigurasi OpenCode (`opencode.json` atau `~/.config/opencode/opencode.json`):

```json
{
  "skills": ["/path/ke/office-skill/SKILL.md"]
}
```

Untuk Claude Code / agent lain yang support system instructions:

```bash
cat SKILL.md  # lalu copy outputnya ke system prompt agent
```

### Sebagai Script Tools Langsung

```bash
# Analisis dokumen skripsi
python3 scripts/docx-tools.py skripsi.docx all

# Validasi setelah edit
python3 scripts/validate-docx.py revised.docx

# Merge runs sebelum edit XML
python3 scripts/merge-runs.py work/unpacked/word/document.xml
```

### Alur Kerja Singkat

```bash
# 1. Setup
mkdir -p work && cd work
cp ../skripsi.docx ../skripsi_backup.docx
unzip -q ../skripsi.docx -d unpacked/

# 2. Merge runs (wajib!)
python3 ../scripts/merge-runs.py unpacked/word/document.xml

# 3. Edit XML (str_replace)
# Buka unpacked/word/document.xml, cari teks target, ganti

# 4. Repack
cd unpacked && zip -Xr ../revised.docx . && cd ..

# 5. Validasi
python3 ../scripts/validate-docx.py revised.docx

# 6. Render & verifikasi
soffice --headless --convert-to pdf revised.docx
pdftotext -layout revised.pdf revised.txt
grep -n "BAB 1\|Tabel 4" revised.txt
```

## Prasyarat

| Tool | Untuk |
|------|-------|
| `python3` + `python-docx` | Baca/tulis DOCX |
| `LibreOffice` (`soffice`) | DOCX → PDF (verifikasi visual) |
| `poppler-utils` (`pdftotext`) | PDF → teks (grep untuk ngecek) |
| `zip`/`unzip` | Ekstrak/repack DOCX (karena DOCX = ZIP) |
| `lxml` | Validasi XML |

## Struktur Project

```
office-skill/
├── SKILL.md                        # Skill utama — panduan skripsi lengkap
├── package.json                    # npx @andypratama3/office-skill
├── index.js
├── .gitignore                      # Otomatis ignore file .docx/.pdf pribadi
├── README.md
├── skills/
│   ├── docx/SKILL.md               # Semua teknik detail edit DOCX
│   ├── pptx/SKILL.md               # PPTX (sidang)
│   ├── pdf/SKILL.md                # PDF handling
│   └── spreadsheet/SKILL.md        # XLSX (data)
├── scripts/
│   ├── validate-docx.py            # Validasi DOCX setelah edit
│   ├── docx-tools.py               # Analisis isi dokumen
│   └── merge-runs.py               # Gabung fragmentasi Word di XML
└── examples/
    └── EXAMPLE_SKRIPSI_WORKFLOW.md
```

## Apa yang Bisa Dilakukan

| Masalah Skripsi | Solusi di Skill Ini |
|----------------|---------------------|
| Heading tidak bernomor | Edit XML: tambah `w:numPr` |
| Caption tabel/gambar tidak urut | Renumber SEQ field |
| Daftar Isi halamannya salah | Hardcode PAGEREF dari hasil render PDF |
| Sitasi in-text tidak punya pasangan | Cross-check otomatis |
| Definisi terdeteksi plagiarism | Template paragraf anti-plagiarism |
| State di kode 9 tapi di teks 8 | Verifikasi langsung ke source code |
| Riwayat hidup duplikat | Deteksi & minta konfirmasi penulis |
| TOC hilang sub-bab baru | Cek bookmark `_Toc` |

## Prinsip Kerja

1. **DOCX = ZIP** — unzip dulu, edit XML, zip ulang. Jangan edit binary langsung.
2. **Verifikasi dengan PDF** — render, `pdftotext`, `grep`. Jangan percaya edit "beres" sebelum lihat hasil render.
3. **Backup sebelum edit** — simpan file asli terpisah.
4. **Satu batch → validasi** — jangan timbun 10 edit tanpa cek.
5. **Source code > screenshot > teks** — kalau bertentangan, kode yang benar.

## License

MIT
