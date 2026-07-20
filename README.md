# Skripsi Skill — AI Agent untuk Revisi Skripsi/Thesis DOCX

Skill pack untuk AI coding agents (OpenCode, Claude Code, dll) yang mengajarkan cara merevisi dokumen skripsi (.docx) secara programatik — edit konten, perbaiki format, sinkronisasi Daftar Isi/Tabel/Gambar, cross-check sitasi, validasi konsistensi lintas bab, dan menjaga orisinalitas tulisan.

## Untuk Siapa

- **Mahasiswa** yang ingin merevisi skripsinya dengan bantuan AI agent
- **AI coding agents** yang perlu tahu cara edit .docx tanpa Microsoft Office
- Siapapun yang bekerja dengan dokumen skripsi berbahasa Indonesia

## Instalasi

### Clone (Rekomendasi)

```bash
git clone https://github.com/andypratama3/office-skripsi.git
cd office-skripsi
pip install python-docx lxml
```

Install dependensi sistem sesuai OS kamu:

**macOS**
```bash
brew install --cask libreoffice poppler
```

**Linux (Ubuntu/Debian)**
```bash
sudo apt update
sudo apt install libreoffice poppler-utils zip unzip
```

**Windows**
```powershell
# PowerShell (Admin)
winget install LibreOffice.LibreOffice
winget install poppler
# atau download manual:
# LibreOffice: https://www.libreoffice.org/download/
# Poppler: https://github.com/oschwartz10612/poppler-windows/releases/
```

### Download ZIP

Download dari https://github.com/andypratama3/office-skripsi, extract, lalu:
```bash
cd office-skripsi-main
pip install python-docx lxml
```

Update: `cd office-skripsi && git pull`

## Cara Pakai

### Sebagai AI Agent Skill

**Auto-configured**: Repo ini sudah punya `opencode.json` yang refer ke `./SKILL.md`. Kalau kamu buka folder `office-skripsi/` di OpenCode, skill otomatis ter-load.

Untuk **global** (semua project), tambahkan ke `~/.config/opencode/opencode.json`:

```json
{
  "skills": {
    "skripsi": "/path/ke/office-skripsi/SKILL.md"
  }
}
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

# Merge runs sebelum edit XML
python scripts/merge-runs.py work/unpacked/word/document.xml
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
pip install git+https://github.com/andypratama3/office-skripsi.git
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
│   ├── validate-docx.py            # Validasi DOCX
│   ├── docx-tools.py               # Analisis isi
│   └── merge-runs.py               # Gabung fragmentasi Word
└── examples/
    └── EXAMPLE_SKRIPSI_WORKFLOW.md
```

## Apa yang Bisa Dilakukan

| Masalah Skripsi | Solusi di Skill Ini |
|----------------|---------------------|
| Heading tidak bernomor | Edit XML: tambah `w:numPr` |
| Caption tabel/gambar tidak urut | Renumber SEQ field |
| Daftar Isi halamannya salah | Hardcode PAGEREF dari render PDF |
| Sitasi tidak punya pasangan | Cross-check otomatis |
| Definisi terdeteksi plagiarisme | Template paragraf anti-plagiarism |
| State di kode 9 tapi di teks 8 | Verifikasi langsung ke source code |
| Riwayat hidup duplikat | Deteksi & minta konfirmasi penulis |
| TOC hilang sub-bab baru | Cek bookmark `_Toc` |

## License

MIT
