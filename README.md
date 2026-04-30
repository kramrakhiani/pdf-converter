# Offline Document Converter

A **privacy-first** offline document converter. No cloud uploads, no data collection — your documents never leave your machine. Designed for sensitive workflows like ID cards, personal records, and internal documents.

## Features

### Privacy & Security
- **100% offline** — all conversions happen locally
- **No telemetry** — no data is sent to external servers
- **Sensitive Document Mode** — 300 DPI for IDs, certificates, and appearance-critical files
- **SHA-256 verification** — every conversion includes checksum verification
- **File validation** — magic byte verification ensures file type integrity

### Supported Conversions

| From | To |
|------|-----|
| `.docx` / `.doc` | `.pdf`, `.txt`, `.docx` |
| `.pptx` / `.ppt` | `.pdf`, `.txt`, `.pptx` |
| `.xlsx` / `.xls` | `.pdf`, `.txt`, `.csv` |
| `.pdf` | `.docx`, `.txt`, `.html`, `.md`, `.rtf`, `.pptx` |
| `.txt/.md/.html/.rtf` | `.pdf`, `.docx`, `.pptx` |
| Images (PNG/JPG/BMP/GIF/TIFF/WebP) | `.pdf` |

### PDF Utilities
- **Merge PDFs** — combine multiple PDF files into one
- **Split PDFs** — extract individual pages as separate files
- **Compress PDFs** — reduce file size
- **Extract Images** — pull embedded images from PDF

### Usage

**Run the GUI:**
```bash
python3 offline_converter.py
```

**CLI conversion:**
```bash
# Single file - saves next to original
python3 offline_converter.py convert document.txt --to pdf

# Batch convert
python3 offline_converter.py convert file1.pdf file2.pdf --to docx

# Directory conversion
python3 offline_converter.py convert-dir ./docs --from-ext pdf --to txt
```

**Build standalone app:**
```bash
pip install pyinstaller
python build.py
# Output: dist/OfflineConverter.app (48MB)
```

### Requirements
```
pymupdf python-docx python-pptx reportlab pypdf pillow openpyxl pytesseract
```

Optional for better quality: `brew install tesseract libreoffice`

### Basic Usage

```bash
# Launch the GUI (default)
python3 offline_converter.py

# Launch the interactive terminal menu
python3 offline_converter.py --cli

# Convert a single file
python3 offline_converter.py convert document.docx --to pdf --output converted.pdf

# Show all supported conversions
python3 offline_converter.py list

# Check installed tools and libraries
python3 offline_converter.py doctor
```

## 📖 Detailed Usage

### Conversion Examples

```bash
# Convert with specific output path
python3 offline_converter.py convert input.docx --to pdf --output ./output.pdf

# Force native renderer (requires LibreOffice or MS Office)
python3 offline_converter.py convert input.docx --to pdf --mode native

# Use pure Python fallback only
python3 offline_converter.py convert input.docx --to pdf --mode python

# Sensitive Document Mode for IDs, certificates, etc.
python3 offline_converter.py convert id-card.pdf --to docx --sensitive --output id-card.docx

# OCR a scanned PDF with language support
python3 offline_converter.py convert scan.pdf --to txt --ocr --ocr-lang eng+hin

# Batch convert multiple files
python3 offline_converter.py convert a.pdf b.pdf c.pdf --to docx --output-dir ./converted

# Convert entire directory
python3 offline_converter.py convert-dir ./inbox --from-ext pdf --to txt --output-dir ./texts --recursive

# Preserve directory structure during batch conversion
python3 offline_converter.py convert-dir ./source --from-ext docx --to pdf --output-dir ./output --preserve-structure
```

### PDF Utilities

```bash
# Merge multiple PDFs
python3 offline_converter.py merge-pdf part1.pdf part2.pdf part3.pdf --output merged.pdf

# Split PDF into individual pages
python3 offline_converter.py split-pdf report.pdf --output-dir ./pages

# Compress a PDF
python3 offline_converter.py compress-pdf large.pdf --output smaller.pdf

# Extract embedded images from a PDF
python3 offline_converter.py extract-images document.pdf --output-dir ./images
```

### Output Control

```bash
# Overwrite existing files
python3 offline_converter.py convert input.docx --to pdf --overwrite

# Skip files that already exist (useful for batch operations)
python3 offline_converter.py convert-dir ./folder --from-ext pdf --to docx --output-dir ./out --skip-existing
```

## Architecture

```
offline_converter_app/
├── cli.py          # CLI interface
├── converter.py    # Core conversion logic
├── writers.py      # File generation (PDF, DOCX, PPTX, etc.)
├── extractors.py   # Text extraction
├── pdf_utils.py    # PDF utilities
├── renderers.py    # Native Office rendering
└── verify.py       # Output verification
```

### Conversion Flow

1. **Input Analysis** — Detect source format and available renderers
2. **Renderer Selection** — Choose best available conversion path based on mode
3. **Conversion** — Execute the conversion with selected engine
4. **Verification** — Validate output file and generate SHA-256 checksum
5. **Reporting** — Display engine used, fidelity level, and verification status

### Fidelity Levels

| Level | Description |
|-------|-------------|
| `native-renderer` | Used LibreOffice or MS Office for pixel-perfect output |
| `visual-preserving` | Embedded pages as images to maintain exact appearance |
| `fallback` | Python-based rendering with best-effort formatting |
| `reconstructed` | Text extracted and re-rendered (formatting may vary) |

## Optional Dependencies

| Tool | Purpose | Installation |
|------|---------|--------------|
| **Tesseract** | OCR for scanned PDFs | `brew install tesseract` (macOS) / `apt install tesseract-ocr` (Linux) |
| **LibreOffice** | High-fidelity DOCX/PPTX → PDF | `brew install --cask libreoffice` (macOS) / `apt install libreoffice` (Linux) |
| **Microsoft Word** | Native DOCX → PDF on macOS | Install Microsoft Word from App Store |
| **Microsoft PowerPoint** | Native PPTX → PDF on macOS | Install Microsoft PowerPoint from App Store |

## Sensitive Document Mode

Use `--sensitive` flag when converting:
- Personal identification documents (passports, driver's licenses, ID cards)
- Certificates and diplomas
- Mark sheets and transcripts
- Any document where visual appearance is more important than editability

This mode:
- Embeds PDF pages as images in DOCX/PPTX outputs (preserves exact appearance)
- Prefers native rendering for DOCX/PPTX → PDF conversions
- Reports verification status and SHA-256 checksums
- Prioritizes visual fidelity over text extraction

## Conversion Report

After each conversion, you'll see a detailed report:

```
Converted: document.docx -> document.pdf
Engine: LibreOffice
Fidelity: native-renderer
Verified: yes
SHA256: a1b2c3d4e5f6...
```

## Limitations

- **PDF → DOCX** embeds pages as images for visual fidelity (not editable)
- **PPTX → PDF** with Python fallback is not pixel-perfect; use LibreOffice for best results
- **Excel** complex formatting may not convert perfectly in Python mode

## Screenshots

### GUI Interface
The application includes a full-featured graphical interface with:
- File selection with drag-and-drop support
- Target format dropdown
- OCR and Sensitive Document Mode toggles
- Real-time renderer hints
- PDF utility buttons (Merge, Split, Compress)
- Tool status panel
- Conversion results viewer

### Terminal Menu
Interactive command-line interface with guided prompts for all operations.


## License

MIT License — feel free to use this tool for personal and commercial projects.

## Acknowledgments

- **PyMuPDF** — PDF processing and rendering
- **python-docx** — DOCX file manipulation
- **python-pptx** — PPTX file manipulation
- **ReportLab** — PDF generation
- **pypdf** — PDF utilities
- **Pillow** — Image processing
- **Tesseract** — OCR engine

---

# pdfconverter
