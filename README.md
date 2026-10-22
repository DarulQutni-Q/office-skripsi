# Skripsi & Makalah Skill — AI Agent untuk Revisi Dokumen Akademik DOCX

Skill pack untuk AI coding agents (OpenCode, Claude Code, dll) yang mengajarkan cara merevisi dokumen akademik (.docx) — skripsi maupun makalah — secara programatik — edit konten, perbaiki format, sinkronisasi Daftar Isi/Tabel/Gambar, cross-check sitasi, validasi konsistensi lintas bab, dan menjaga orisinalitas tulisan.

> **Catatan fork:** Repo ini adalah fork dari [andypratama3/office-skripsi](https://github.com/andypratama3/office-skripsi)
> yang dikelola dan dikembangkan aktif di sini. Perubahan utama di fork ini dibanding upstream:
> - `scripts/numbering.py` — penomoran multilevel otomatis untuk heading yang diketik manual (kontribusi utama)
> - Perbaikan `scripts/merge-runs.py` (false-positive prosa berangka, double-escape entity, paragraf terhapus saat merge)
> - Sinkronisasi rutin dari upstream

## Untuk Siapa

- **Mahasiswa** yang ingin merevisi skripsinya dengan bantuan AI agent
- **Mahasiswa** yang mengerjakan makalah / tugas kuliah (didukung penuh sebagai warga kelas satu)
- **AI coding agents** yang perlu tahu cara edit .docx tanpa Microsoft Office
- Siapapun yang bekerja dengan dokumen skripsi berbahasa Indonesia

## Jenis Dokumen yang Didukung

| Tipe | Struktur bab yang divalidasi | Contoh pakai |
|------|------------------------------|--------------|
| `skripsi` (default) | PENDAHULUAN, TINJAUAN PUSTAKA, METODE PENELITIAN, HASIL DAN PEMBAHASAN, KESIMPULAN, DAFTAR PUSTAKA | `python scripts/validate-docx.py revisi.docx` |
| `makalah` | PENDAHULUAN, PEMBAHASAN, PENUTUP, DAFTAR PUSTAKA | `python scripts/validate-docx.py makalah.docx --type makalah` |
| custom | sesukamu | `python scripts/validate-docx.py bebas.docx --sections "BAB 1, BAB 2"` |
| tanpa cek bab | — | `python scripts/validate-docx.py bebas.docx --type none` |

Teknik lain (numbering, italic-foreign, merge-runs) bekerja untuk kedua tipe karena beroperasi di level heading/style Word, bukan nama bab.

## Instalasi

### 1. Clone

```bash
git clone https://github.com/DarulQutni-Q/office-skripsi.git
cd office-skripsi
pip install python-docx lxml
```

### 2. Pasang skill (otomatis dikenali OpenCode)

**Cara A — Global (skill tersedia di semua project):**

```bash
mkdir -p ~/.config/opencode/skills/skripsi
ln -sf $(pwd)/SKILL.md ~/.config/opencode/skills/skripsi/SKILL.md
```

OpenCode otomatis scan `~/.config/opencode/skills/*/SKILL.md` — tidak perlu tambah config apapun.

**Cara B — Per project** (kalau `~/.config/opencode/skills/` belum ada):

Tambah ke `~/.config/opencode/opencode.json`:
```json
{
  "skills": {
    "paths": ["/path/ke/office-skripsi"]
  }
}
```

### 3. Install dependensi sistem

| OS | LibreOffice | Poppler |
|----|-------------|---------|
| **macOS** | `brew install --cask libreoffice` | `brew install poppler` |
| **Linux** | `sudo apt install libreoffice` | `sudo apt install poppler-utils` |
| **Windows** | [Download](https://www.libreoffice.org/download/) | [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases) |

### 4. Verifikasi

```bash
# Cek skill terdaftar
npx opencode debug skill
# atau kalau opencode terinstall global:
opencode debug skill
```

Update: `cd office-skripsi && git pull`

---

## Cara Pakai

### Sebagai AI Agent Skill

**Auto-configured**: Repo ini sudah punya `opencode.json` dengan `skills.paths: ["."]`. Buka folder `office-skripsi/` di OpenCode → skill langsung ter-load.

**Global** (semua project):
```bash
ln -sf $(pwd)/SKILL.md ~/.config/opencode/skills/skripsi/SKILL.md
```

Untuk Claude Code / agent lain:
```bash
# macOS/Linux
cat SKILL.md | pbcopy
# Windows PowerShell
Get-Content SKILL.md | Set-Clipboard
```

### Sebagai Script Tools Langsung

```bash
# Analisis dokumen skripsi
python scripts/docx-tools.py skripsi.docx all

# Validasi setelah edit
python scripts/validate-docx.py revised.docx
# atau untuk makalah:
# python scripts/validate-docx.py makalah.docx --type makalah

# Merge runs sebelum edit XML
python scripts/merge-runs.py work/unpacked/word/document.xml

# Penomoran multilevel otomatis (setelah merge-runs): audit / fix heading hardcoded
python scripts/numbering.py work/unpacked/ analyze
python scripts/numbering.py work/unpacked/ convert

# Italic hanya untuk istilah asing (audit dulu, lalu terapkan; .docx asli tidak ditimpa)
python scripts/italic-foreign.py audit skripsi.docx --strip-other
python scripts/italic-foreign.py apply skripsi.docx --strip-other

# Audit font PDF (deteksi Calibri/theme-font tersisa setelah konversi TNR)
pdftohtml -xml -i revised.pdf /tmp/doc.xml
python scripts/audit-pdf-fonts.py /tmp/doc.xml --target Times
```

Gunakan `python` (bukan `python3`) di Windows. Di macOS/Linux `python3`.

### Alur Kerja Singkat (macOS/Linux)

```bash
mkdir -p work && cd work
cp ../skripsi.docx ../skripsi_backup.docx
unzip -q ../skripsi.docx -d unpacked/
python3 ../scripts/merge-runs.py unpacked/word/document.xml
# edit XML...
cd unpacked && zip -Xr ../revised.docx . && cd ..
python3 ../scripts/validate-docx.py revised.docx
soffice --headless --convert-to pdf revised.docx
pdftotext -layout revised.pdf revised.txt
grep -n "BAB 1\|Tabel 4" revised.txt
```

### Alur Kerja Singkat (Windows PowerShell)

```powershell
New-Item -ItemType Directory -Path work -Force
Copy-Item ../skripsi.docx ../skripsi_backup.docx
Expand-Archive -Path ../skripsi.docx -DestinationPath unpacked/
python scripts/merge-runs.py unpacked/word/document.xml
# edit XML...
Compress-Archive -Path unpacked/* -DestinationPath ../revised.docx -Force
python scripts/validate-docx.py revised.docx
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf revised.docx
pdftotext -layout revised.pdf revised.txt
Select-String -Path revised.txt -Pattern "BAB 1|Tabel 4"
```

### Via `pip install` (Alternatif)

```bash
pip install git+https://github.com/DarulQutni-Q/office-skripsi.git
```

## Prasyarat

| Tool | Install | Untuk |
|------|---------|-------|
| Python 3 | [python.org](https://python.org) | Menjalankan scripts |
| `python-docx` | `pip install python-docx` | Baca/tulis DOCX |
| `lxml` | `pip install lxml` | Validasi XML |
| LibreOffice | [libreoffice.org](https://libreoffice.org) | DOCX → PDF |
| Poppler | `brew install poppler` / `sudo apt install poppler-utils` | PDF → teks |
| zip/unzip | Bawaan OS | Ekstrak/repack DOCX |

## Struktur Project

```
office-skripsi/
├── SKILL.md                        # Skill utama — panduan skripsi
├── package.json
├── index.js
├── .gitignore                      # Ignore file .docx/.pdf pribadi
├── README.md
├── skills/
│   ├── docx/SKILL.md               # Semua teknik edit DOCX
│   ├── pptx/SKILL.md               # PPTX (sidang)
│   ├── pdf/SKILL.md                # PDF handling
│   └── spreadsheet/SKILL.md        # XLSX (data)
├── scripts/
│   ├── validate-docx.py            # Validasi DOCX (--type skripsi/makalah)
│   ├── docx-tools.py               # Analisis isi
│   ├── merge-runs.py               # Gabung fragmentasi Word
│   ├── numbering.py                # Penomoran multilevel otomatis (heading hardcoded)
│   └── italic-foreign.py           # Italic hanya untuk istilah asing (aman untuk caption/TOC)
└── examples/
    ├── EXAMPLE_SKRIPSI_WORKFLOW.md
    └── EXAMPLE_MAKALAH_WORKFLOW.md
└── tests/
    ├── test_numbering.py
    ├── test_italic_foreign.py
    └── test_validate_docx.py
```

## Apa yang Bisa Dilakukan

| Masalah Skripsi | Solusi di Skill Ini |
|----------------|---------------------|
| Heading tidak bernomor | Edit XML: tambah `w:numPr` / jalankan `numbering.py convert` |
| Seluruh kalimat miring / istilah asing tidak miring | `italic-foreign.py audit` → `apply --strip-other` |
| Caption tabel/gambar tidak urut | Renumber SEQ field |
| Daftar Isi halamannya salah | Hardcode PAGEREF dari render PDF |
| Sitasi tidak punya pasangan | Cross-check otomatis |
| Definisi terdeteksi plagiarisme | Template paragraf anti-plagiarism |
| State di kode 9 tapi di teks 8 | Verifikasi langsung ke source code |
| Riwayat hidup duplikat | Deteksi & minta konfirmasi penulis |
| TOC hilang sub-bab baru | Cek bookmark `_Toc` |

## License

MIT
