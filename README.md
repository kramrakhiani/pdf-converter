# Offline Document Converter

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-cross--platform-lightgrey.svg)](#installation)

A **privacy-first** Python CLI and GUI for local file conversion. No cloud uploads, no data collection — your documents never leave your machine. Designed for sensitive workflows like ID cards, personal records, and internal documents.

## Features

### Privacy & Security
- **100% offline** — all conversions happen locally on your machine
- **No telemetry** — no data is sent to external servers
- **Sensitive Document Mode** — optimized paths for IDs, certificates, and appearance-critical files
- **SHA-256 verification** — every conversion includes checksum verification

### Supported Conversions

| From | To |
|------|-----|
| `.docx` | `.pdf`, `.txt`, `.docx` |
| `.pptx` | `.pdf`, `.txt`, `.pptx` |
| `.pdf` | `.docx`, `.txt`, `.html`, `.md`, `.rtf`, `.pptx` |
| `.txt/.md/.html/.rtf` | `.pdf`, `.docx`, `.pptx` |
| Images (PNG/JPG/BMP/GIF/TIFF/WebP) | `.pdf` |

### PDF Utilities
- **Merge PDFs** — combine multiple PDF files into one
- **Split PDFs** — extract individual pages as separate files
- **Compress PDFs** — reduce file size with content stream compression
- **Extract Images** — pull embedded images from PDF documents

### Smart Rendering
- **Auto mode** — automatically selects the best available renderer
- **Native mode** — uses system Office applications for highest fidelity
- **Python fallback** — pure Python rendering when native tools aren't available

### Multiple Interfaces
- **CLI** — full-featured command-line interface
- **Interactive terminal menu** — guided step-by-step conversion
- **Desktop GUI** — user-friendly graphical interface built with Tkinter


### Installation

```bash
# Install Python dependencies from requirements.txt
python3 -m pip install --user -r requirements.txt

# Or install individually
python3 -m pip install --user pymupdf python-docx python-pptx reportlab pypdf pillow

# Optional: Install Tesseract for OCR support
# macOS: brew install tesseract
# Ubuntu: sudo apt install tesseract-ocr
# Windows: Download from GitHub releases

# Optional: Install LibreOffice for higher-fidelity conversions
# macOS: brew install --cask libreoffice
# Ubuntu: sudo apt install libreoffice
```

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
├── cli.py          # Command-line interface and argument parsing
├── common.py       # Shared types, utilities, and tool detection
├── converter.py    # Main conversion orchestration logic
├── extractors.py   # Text extraction from various formats
├── gui.py          # Tkinter-based graphical user interface
├── pdf_utils.py    # PDF merge, split, compress, and image extraction
├── renderers.py    # Native Office rendering (LibreOffice, MS Word/PPT)
├── verify.py       # Output verification and checksum generation
└── writers.py      # Output file generation for all formats
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

- **Legacy formats** (`.doc`, `.ppt`) are not supported in the pure-Python path
- **PDF → DOCX** prioritizes visual fidelity (pages as images) over editability
- **PPTX → PDF** Python fallback is not pixel-perfect; use native renderers for best results
- **Complex formatting** (advanced tables, custom fonts, embedded media) may not convert perfectly in fallback mode

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

