from __future__ import annotations

from pathlib import Path

import fitz
from pypdf import PdfReader, PdfWriter

from .common import ConversionError, ensure_file


def merge_pdfs(inputs: list[Path], output_path: Path) -> None:
    if len(inputs) < 2:
        raise ConversionError("Please provide at least two PDF files to merge.")
    writer = PdfWriter()
    for path in inputs:
        ensure_file(path)
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as handle:
        writer.write(handle)


def split_pdf(input_path: Path, output_dir: Path) -> None:
    ensure_file(input_path)
    reader = PdfReader(str(input_path))
    output_dir.mkdir(parents=True, exist_ok=True)
    for index, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        out_path = output_dir / f"{input_path.stem}-{index:03d}.pdf"
        with out_path.open("wb") as handle:
            writer.write(handle)


def compress_pdf(input_path: Path, output_path: Path) -> None:
    ensure_file(input_path)
    reader = PdfReader(str(input_path))
    writer = PdfWriter()
    for page in reader.pages:
        try:
            page.compress_content_streams()
        except Exception:
            pass
        writer.add_page(page)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as handle:
        writer.write(handle)


def extract_pdf_images(input_path: Path, output_dir: Path) -> None:
    ensure_file(input_path)
    doc = fitz.open(str(input_path))
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        count = 0
        for page_index, page in enumerate(doc, start=1):
            for image_index, image_info in enumerate(page.get_images(full=True), start=1):
                xref = image_info[0]
                image = doc.extract_image(xref)
                ext = image["ext"]
                data = image["image"]
                out_path = output_dir / f"{input_path.stem}-p{page_index:03d}-{image_index:02d}.{ext}"
                out_path.write_bytes(data)
                count += 1
        if count == 0:
            raise ConversionError(f"No embedded images were found in {input_path}")
    finally:
        doc.close()

