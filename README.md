# Offline Document Converter

A **privacy-first** document converter that runs 100% offline. Your files never leave your machine — no cloud uploads, no data collection. Perfect for sensitive documents like IDs, certificates, and personal records.

## Quick Start

```bash
# Launch the GUI (easiest way)
python3 offline_converter.py
```

That's it! The GUI opens automatically. Drag and drop your files, select the output format, and convert.

---

## How to Use

### Method 1: GUI (Recommended)

The easiest way — no command line needed.

```bash
python3 offline_converter.py
```

This opens a window where you can:
- **Drag & drop** files or click to browse
- **Select output format** from the dropdown
- **Enable OCR** for scanned PDFs
- **Sensitive Mode** for IDs and certificates (300 DPI)
- **Merge, split, or compress** PDFs with one click

### Method 2: Terminal Menu

Interactive menu with guided prompts.

```bash
python3 offline_converter.py --cli
```

Follow the on-screen instructions — perfect if you prefer menus over commands.

### Method 3: CLI Commands

Direct commands for automation and batch processing.

**Convert a single file:**
```bash
python3 offline_converter.py convert document.docx --to pdf
```

**Batch convert multiple files:**
```bash
python3 offline_converter.py convert file1.pdf file2.pdf --to docx
```

**Convert entire folder:**
```bash
python3 offline_converter.py convert-dir ./docs --from-ext pdf --to txt
```

**PDF utilities:**
```bash
python3 offline_converter.py merge-pdf a.pdf b.pdf --output merged.pdf
python3 offline_converter.py split-pdf document.pdf --output-dir ./pages
python3 offline_converter.py compress-pdf large.pdf --output small.pdf
```

---

## Supported Conversions

| From | To |
|------|-----|
| `.docx` / `.doc` | `.pdf`, `.txt`, `.docx` |
| `.pptx` / `.ppt` | `.pdf`, `.txt`, `.pptx` |
| `.xlsx` / `.xls` | `.pdf`, `.txt`, `.csv` |
| `.pdf` | `.docx`, `.txt`, `.html`, `.md`, `.rtf`, `.pptx` |
| `.txt` / `.md` / `.html` / `.rtf` | `.pdf`, `.docx`, `.pptx` |
| Images (PNG/JPG/BMP/GIF/TIFF/WebP) | `.pdf` |

### PDF Tools
- **Merge** — combine multiple PDFs into one
- **Split** — extract pages as separate files
- **Compress** — reduce file size
- **Extract Images** — pull images from PDF

---

## Features

- **100% Offline** — all processing happens locally on your machine
- **No Data Collection** — nothing is sent to any server
- **Sensitive Document Mode** — 300 DPI for IDs, certificates, and appearance-critical files
- **SHA-256 Verification** — every conversion is verified with a checksum
- **Magic Byte Validation** — confirms file types are correct

---

## Installation

```bash
pip install pymupdf python-docx python-pptx reportlab pypdf pillow openpyxl pytesseract
```

**Optional (better quality):**
```bash
# macOS
brew install tesseract libreoffice

# Linux
sudo apt install tesseract-ocr libreoffice
```

### Build a Standalone App

```bash
pip install pyinstaller
python build.py
# Output: dist/OfflineConverter.app
```

---

## Common Examples

```bash
# Convert with specific output path
python3 offline_converter.py convert input.docx --to pdf --output ./output.pdf

# OCR a scanned PDF
python3 offline_converter.py convert scan.pdf --to txt --ocr

# Sensitive mode for IDs/certificates
python3 offline_converter.py convert id-card.pdf --to docx --sensitive

# Skip existing files in batch mode
python3 offline_converter.py convert-dir ./folder --from-ext pdf --to docx --skip-existing
```

---

## Other Commands

```bash
python3 offline_converter.py list          # Show all supported conversions
python3 offline_converter.py doctor        # Check installed tools
```

---

## Architecture

```
offline_converter_app/
├── cli.py          # CLI interface
├── converter.py    # Core conversion logic
├── writers.py      # File generation
├── extractors.py   # Text extraction
├── pdf_utils.py    # PDF utilities
├── renderers.py    # Native Office rendering
└── verify.py       # Output verification
```

---

## Limitations

- **PDF → DOCX** embeds pages as images (not editable) for visual fidelity
- **PPTX → PDF** with Python fallback may not be pixel-perfect; use LibreOffice for best results
- Complex Excel formatting may vary in Python mode

---

## License

MIT License — free for personal and commercial use.

---

##Acknowledgments

- **PyMuPDF** — PDF processing
- **python-docx** — DOCX manipulation
- **python-pptx** — PPTX manipulation
- **ReportLab** — PDF generation
- **pypdf** — PDF utilities
- **Pillow** — Image processing
- **Tesseract** — OCR engine