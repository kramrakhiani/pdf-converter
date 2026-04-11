from __future__ import annotations

from pathlib import Path

from .common import ConversionError, ConversionOptions, ConversionReport, IMAGE_INPUTS, TEXT_INPUTS, WORD_INPUTS, ToolState, ensure_file, normalize_ext
from .extractors import extract_text
from .renderers import render_office_to_pdf
from .verify import verify_output
from .writers import (
    pdf_to_pptx,
    text_to_pptx,
    write_docx,
    write_docx_from_pdf_pages,
    write_html,
    write_image_pdf,
    write_md,
    write_pdf_from_docx,
    write_pdf_from_pptx,
    write_pdf_from_text,
    write_rtf,
    write_txt,
)


def convert_one(
    input_path: Path,
    target_ext: str,
    output_path: Path,
    tools: ToolState,
    *,
    ocr: bool = False,
    ocr_lang: str = "eng",
    mode: str = "auto",
) -> ConversionReport:
    return convert_one_with_options(
        input_path,
        target_ext,
        output_path,
        tools,
        ConversionOptions(ocr=ocr, ocr_lang=ocr_lang, mode=mode, sensitive=False),
    )


def convert_one_with_options(
    input_path: Path,
    target_ext: str,
    output_path: Path,
    tools: ToolState,
    options: ConversionOptions,
) -> ConversionReport:
    ensure_file(input_path)
    source_ext = normalize_ext(input_path)
    target_ext = target_ext.lower().lstrip(".")
    ocr = options.ocr
    ocr_lang = options.ocr_lang
    mode = options.mode
    sensitive = options.sensitive

    if source_ext == target_ext:
        raise ConversionError(f"Source and target formats are the same for {input_path.name}")

    if sensitive:
        if target_ext == "docx" and source_ext == "pdf":
            write_docx_from_pdf_pages(input_path, output_path)
            return verify_output(
                input_path,
                output_path,
                ConversionReport(
                    engine="python-page-image-docx",
                    fidelity="visual-preserving",
                    notes=(
                        "Sensitive Document Mode kept each PDF page as an image inside the DOCX.",
                        "This prioritizes appearance over editability.",
                    ),
                ),
            )
        if target_ext == "pptx" and source_ext == "pdf":
            pdf_to_pptx(input_path, output_path)
            return verify_output(
                input_path,
                output_path,
                ConversionReport(
                    engine="python-page-image-pptx",
                    fidelity="visual-preserving",
                    notes=(
                        "Sensitive Document Mode kept each PDF page as an image slide.",
                        "This prioritizes appearance over editability.",
                    ),
                ),
            )
        if target_ext == "pdf" and source_ext in {"docx", "pptx"} and mode == "auto":
            mode = "auto"

    if target_ext == "txt":
        write_txt(extract_text(input_path, tools, ocr=ocr, ocr_lang=ocr_lang), output_path)
        return verify_output(input_path, output_path, ConversionReport(engine="python-text", fidelity="reconstructed"))
    if target_ext == "md":
        write_md(extract_text(input_path, tools, ocr=ocr, ocr_lang=ocr_lang), output_path)
        return verify_output(input_path, output_path, ConversionReport(engine="python-text", fidelity="reconstructed"))
    if target_ext == "html":
        write_html(extract_text(input_path, tools, ocr=ocr, ocr_lang=ocr_lang), output_path)
        return verify_output(input_path, output_path, ConversionReport(engine="python-text", fidelity="reconstructed"))
    if target_ext == "rtf":
        write_rtf(extract_text(input_path, tools, ocr=ocr, ocr_lang=ocr_lang), output_path)
        return verify_output(input_path, output_path, ConversionReport(engine="python-text", fidelity="reconstructed"))
    if target_ext == "docx":
        if source_ext == "pdf":
            write_docx_from_pdf_pages(input_path, output_path)
            return verify_output(input_path, output_path, ConversionReport(
                engine="python-page-image-docx",
                fidelity="visual-preserving",
                notes=("Each PDF page was embedded as an image inside the DOCX.",),
            ))
        write_docx(extract_text(input_path, tools, ocr=ocr, ocr_lang=ocr_lang), output_path)
        return verify_output(input_path, output_path, ConversionReport(engine="python-text", fidelity="reconstructed"))
    if target_ext == "pdf":
        if source_ext in IMAGE_INPUTS:
            write_image_pdf(input_path, output_path)
            return verify_output(input_path, output_path, ConversionReport(engine="python-image-pdf", fidelity="visual-preserving"))
        if source_ext in {"docx", "pptx"} and mode in {"auto", "native"}:
            try:
                engine = render_office_to_pdf(input_path, output_path, tools)
                native_notes: tuple[str, ...] = ()
                if sensitive:
                    native_notes = ("Sensitive Document Mode used a native Office renderer for the highest available fidelity.",)
                return verify_output(input_path, output_path, ConversionReport(engine=engine, fidelity="native-renderer", notes=native_notes))
            except ConversionError as exc:
                if mode == "native":
                    raise
                if sensitive:
                    native_failure_note = (
                        "Sensitive Document Mode could not use a confirmed native renderer, so the safest available Python fallback was used: "
                        f"{exc}"
                    )
                else:
                    native_failure_note = f"Native renderer attempt failed, so Python fallback was used: {exc}"
            else:
                native_failure_note = None
        else:
            native_failure_note = None
        if source_ext == "docx":
            notes = write_pdf_from_docx(input_path, output_path)
            if native_failure_note:
                notes = (native_failure_note, *notes)
            return verify_output(input_path, output_path, ConversionReport(engine="python-docx-pdf", fidelity="fallback", notes=notes))
        if source_ext == "pptx":
            notes = write_pdf_from_pptx(input_path, output_path)
            if native_failure_note:
                notes = (native_failure_note, *notes)
            return verify_output(input_path, output_path, ConversionReport(engine="python-pptx-pdf", fidelity="fallback", notes=notes))
        write_pdf_from_text(
            extract_text(input_path, tools, ocr=ocr, ocr_lang=ocr_lang),
            output_path,
            title=input_path.stem,
        )
        return verify_output(input_path, output_path, ConversionReport(engine="python-text-pdf", fidelity="reconstructed"))
    if target_ext == "pptx":
        if source_ext == "pdf":
            pdf_to_pptx(input_path, output_path)
            return verify_output(input_path, output_path, ConversionReport(
                engine="python-page-image-pptx",
                fidelity="visual-preserving",
                notes=("Each PDF page became an image slide.",),
            ))
        if source_ext in TEXT_INPUTS | WORD_INPUTS:
            text_to_pptx(
                extract_text(input_path, tools, ocr=ocr, ocr_lang=ocr_lang),
                output_path,
                title=input_path.stem,
            )
            return verify_output(input_path, output_path, ConversionReport(engine="python-text-pptx", fidelity="reconstructed"))
        raise ConversionError(f".{source_ext} -> .pptx is not supported.")

    raise ConversionError(
        f"Unsupported conversion: .{source_ext} -> .{target_ext}. Run `list` to see supported pairs."
    )
