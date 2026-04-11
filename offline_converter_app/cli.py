from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

import fitz

from .common import ConversionError, ConversionOptions, ToolState, collect_directory_inputs, ensure_output_allowed, output_path_for
from .converter import convert_one, convert_one_with_options
from .gui import launch_gui
from .pdf_utils import compress_pdf, extract_pdf_images, merge_pdfs, split_pdf
from .renderers import renderer_candidates


def build_routes() -> list[tuple[str, str, str]]:
    return [
        ("docx", "pdf", "Renderer-based PDF if available, otherwise Python fallback"),
        ("pdf", "docx", "PDF pages embedded into DOCX for higher visual fidelity"),
        ("pdf", "txt", "PDF text extracted to plain text"),
        ("pdf", "html", "PDF text extracted to HTML"),
        ("pdf", "md", "PDF text extracted to Markdown-like text"),
        ("pdf", "rtf", "PDF text extracted to RTF"),
        ("pdf", "pptx", "Each PDF page becomes an image slide"),
        ("pptx", "pdf", "Renderer-based PDF if available, otherwise Python fallback"),
        ("pptx", "txt", "Slide text extracted to plain text"),
        ("txt", "pdf", "Plain text rendered to PDF"),
        ("txt", "docx", "Plain text written to DOCX"),
        ("txt", "pptx", "Plain text split into slides"),
        ("md", "pdf", "Markdown-like text rendered to PDF"),
        ("md", "docx", "Markdown-like text written to DOCX"),
        ("html", "pdf", "HTML text rendered to PDF"),
        ("html", "docx", "HTML text written to DOCX"),
        ("rtf", "pdf", "RTF text rendered to PDF"),
        ("rtf", "docx", "RTF text written to DOCX"),
        ("png/jpg/jpeg/bmp/gif/tiff/webp", "pdf", "Image wrapped into PDF"),
    ]


def format_report(report) -> list[str]:
    lines = [f"Engine: {report.engine}", f"Fidelity: {report.fidelity}", f"Verified: {'yes' if report.verified else 'no'}"]
    if report.checksum:
        lines.append(f"SHA256: {report.checksum}")
    lines.extend(f"Note: {note}" for note in report.notes)
    return lines


def print_routes() -> None:
    print("Supported cross-platform offline conversions")
    print("=" * 41)
    for source, target, note in build_routes():
        print(f"{source:>28} -> {target:<5} | {note}")
    print("\nPDF utilities")
    print("=" * 13)
    print("merge-pdf       Combine multiple PDFs into one")
    print("split-pdf       Split one PDF into one file per page")
    print("compress-pdf    Apply stream compression to a PDF")
    print("extract-images  Extract embedded images from a PDF")
    print("\nNotes")
    print("=" * 5)
    print("Legacy binary formats like .doc and .ppt are not included in the pure-Python path.")
    print("The fully cross-platform Office path targets modern .docx and .pptx files.")


def doctor(tools: ToolState) -> int:
    print("Cross-platform converter check")
    print("=" * 30)
    modules = {
        "PyMuPDF": fitz.__doc__.splitlines()[0] if fitz.__doc__ else "installed",
        "python-docx": "installed",
        "python-pptx": "installed",
        "reportlab": "installed",
        "pypdf": "installed",
    }
    for name, detail in modules.items():
        print(f"[ok] {name}: {detail}")
    if tools.tesseract:
        print(f"[ok] tesseract: {tools.tesseract}")
    else:
        print("[optional] tesseract: missing")
    if tools.libreoffice:
        print(f"[ok] libreoffice: {tools.libreoffice}")
    else:
        print("[optional] libreoffice: missing")
    if tools.platform == "Darwin":
        word_status = tools.mac_word_app if tools.osascript and tools.mac_word_app else "missing"
        ppt_status = tools.mac_powerpoint_app if tools.osascript and tools.mac_powerpoint_app else "missing"
        print(f"[optional] Microsoft Word (macOS): {word_status}")
        print(f"[optional] Microsoft PowerPoint (macOS): {ppt_status}")
    if tools.platform == "Windows":
        print(f"[optional] Windows Office automation host: {'available' if tools.powershell else 'missing'}")
    print("\nNotes")
    print("=" * 5)
    print("DOCX/PPTX -> PDF tries native renderers first when available, then falls back to Python mode.")
    print("OCR for scanned PDFs still needs the optional tesseract binary.")
    return 0


def cmd_convert(args: argparse.Namespace, tools: ToolState) -> int:
    inputs = [Path(item).expanduser().resolve() for item in args.inputs]
    target_ext = args.to.lower().lstrip(".")
    if args.output and len(inputs) > 1:
        raise ValueError("Use --output-dir for multiple input files.")

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
        if output_path.exists() and output_path.is_dir():
            raise ValueError("--output must be a file path, not a directory.")
        if ensure_output_allowed(output_path, overwrite=args.overwrite, skip_existing=args.skip_existing):
            report = convert_one_with_options(
                inputs[0],
                target_ext,
                output_path,
                tools,
                ConversionOptions(
                    ocr=args.ocr,
                    ocr_lang=args.ocr_lang,
                    mode=args.mode,
                    sensitive=args.sensitive,
                ),
            )
            print(f"Converted: {inputs[0]} -> {output_path}")
            for line in format_report(report):
                print(line)
        else:
            print(f"Skipped existing file: {output_path}")
        return 0

    output_dir = Path(args.output_dir or ".").expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    for input_path in inputs:
        destination = output_path_for(input_path, output_dir, target_ext)
        if ensure_output_allowed(destination, overwrite=args.overwrite, skip_existing=args.skip_existing):
            report = convert_one_with_options(
                input_path,
                target_ext,
                destination,
                tools,
                ConversionOptions(
                    ocr=args.ocr,
                    ocr_lang=args.ocr_lang,
                    mode=args.mode,
                    sensitive=args.sensitive,
                ),
            )
            print(f"Converted: {input_path} -> {destination}")
            for line in format_report(report):
                print(line)
        else:
            print(f"Skipped existing file: {destination}")
    return 0


def cmd_convert_dir(args: argparse.Namespace, tools: ToolState) -> int:
    input_dir = Path(args.input_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    inputs = collect_directory_inputs(input_dir, args.from_ext, args.recursive)
    converted = 0
    skipped = 0
    for input_path in inputs:
        relative_parent = input_path.parent.relative_to(input_dir) if args.preserve_structure else Path()
        destination = output_path_for(input_path, output_dir / relative_parent, args.to)
        if ensure_output_allowed(destination, overwrite=args.overwrite, skip_existing=args.skip_existing):
            report = convert_one_with_options(
                input_path,
                args.to,
                destination,
                tools,
                ConversionOptions(
                    ocr=args.ocr,
                    ocr_lang=args.ocr_lang,
                    mode=args.mode,
                    sensitive=args.sensitive,
                ),
            )
            print(f"Converted: {input_path} -> {destination}")
            for line in format_report(report):
                print(line)
            converted += 1
        else:
            print(f"Skipped existing file: {destination}")
            skipped += 1
    print(f"Finished. Converted: {converted}, skipped: {skipped}")
    return 0


def cmd_merge_pdf(args: argparse.Namespace, _tools: ToolState) -> int:
    inputs = [Path(item).expanduser().resolve() for item in args.inputs]
    output_path = Path(args.output).expanduser().resolve()
    ensure_output_allowed(output_path, overwrite=args.overwrite, skip_existing=False)
    merge_pdfs(inputs, output_path)
    print(f"Merged {len(inputs)} PDFs into {output_path}")
    return 0


def cmd_split_pdf(args: argparse.Namespace, _tools: ToolState) -> int:
    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    split_pdf(input_path, output_dir)
    print(f"Split {input_path} into single-page PDFs under {output_dir}")
    return 0


def cmd_compress_pdf(args: argparse.Namespace, _tools: ToolState) -> int:
    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    ensure_output_allowed(output_path, overwrite=args.overwrite, skip_existing=False)
    compress_pdf(input_path, output_path)
    print(f"Compressed {input_path} -> {output_path}")
    return 0


def cmd_extract_images(args: argparse.Namespace, _tools: ToolState) -> int:
    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    extract_pdf_images(input_path, output_dir)
    print(f"Extracted embedded images from {input_path} into {output_dir}")
    return 0


def cmd_gui(_args: argparse.Namespace, tools: ToolState) -> int:
    return launch_gui(tools)


def prompt_bool(message: str, default: bool = False) -> bool:
    suffix = " [Y/n]: " if default else " [y/N]: "
    value = input(message + suffix).strip().lower()
    if not value:
        return default
    return value in {"y", "yes"}


def choose_path_with_dialog(*, title: str, select_dir: bool) -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        if select_dir:
            selected = filedialog.askdirectory(title=title, mustexist=False)
        else:
            selected = filedialog.askopenfilename(title=title)
    finally:
        root.destroy()

    if not selected:
        return None
    return Path(selected).expanduser().resolve()


def prompt_path(
    message: str,
    *,
    must_exist: bool = False,
    is_dir: bool | None = None,
    default: Path | None = None,
) -> Path:
    chooser_title = "Choose folder" if is_dir else "Choose file"
    if is_dir is not None:
        use_picker = prompt_bool(f"{message} Open system {'folder' if is_dir else 'file'} picker?", default=True)
        if use_picker:
            selected = choose_path_with_dialog(title=chooser_title, select_dir=is_dir)
            if selected is not None:
                if must_exist and not selected.exists():
                    print(f"Path does not exist: {selected}")
                elif is_dir is True and selected.exists() and not selected.is_dir():
                    print(f"Expected a directory: {selected}")
                elif is_dir is False and selected.exists() and not selected.is_file():
                    print(f"Expected a file: {selected}")
                else:
                    print(f"Selected: {selected}")
                    return selected
            else:
                print("No selection made. Falling back to manual path entry.")

    while True:
        value = input(message).strip()
        if not value:
            if default is not None:
                print(f"Using default: {default}")
                return default.resolve()
            print("Please enter a path.")
            continue
        path = Path(value).expanduser().resolve()
        if must_exist and not path.exists():
            print(f"Path does not exist: {path}")
            continue
        if is_dir is True and path.exists() and not path.is_dir():
            print(f"Expected a directory: {path}")
            continue
        if is_dir is False and path.exists() and not path.is_file():
            print(f"Expected a file: {path}")
            continue
        return path


def interactive_convert(tools: ToolState) -> int:
    input_path = prompt_path("Input file path: ", must_exist=True, is_dir=False)
    target = input("Convert to extension (pdf/docx/txt/html/md/rtf/pptx): ").strip().lower().lstrip(".")
    if not target:
        print("No target format entered.")
        return 1
    output_path = prompt_path(
        "Output file path: ",
        must_exist=False,
        default=(Path.cwd() / f"{input_path.stem}.{target}"),
    )
    ocr = prompt_bool("Use OCR for PDFs?", default=False)
    ocr_lang = "eng"
    if ocr:
        ocr_lang = input("OCR language code [eng]: ").strip() or "eng"
    print(f"Available renderer candidates for this file: {', '.join(renderer_candidates(input_path.suffix.lower().lstrip('.'), tools)) or 'none'}")
    mode = input("Mode [auto/native/python]: ").strip().lower() or "auto"
    sensitive = prompt_bool("Sensitive Document Mode? Use the safest visual-preserving path.", default=False)
    overwrite = prompt_bool("Overwrite output if it exists?", default=False)
    skip_existing = False if overwrite else prompt_bool("Skip existing output instead of failing?", default=True)
    if ensure_output_allowed(output_path, overwrite=overwrite, skip_existing=skip_existing):
        report = convert_one_with_options(
            input_path,
            target,
            output_path,
            tools,
            ConversionOptions(ocr=ocr, ocr_lang=ocr_lang, mode=mode, sensitive=sensitive),
        )
        print(f"Converted: {input_path} -> {output_path}")
        for line in format_report(report):
            print(line)
    else:
        print(f"Skipped existing file: {output_path}")
    return 0


def interactive_convert_dir(tools: ToolState) -> int:
    input_dir = prompt_path("Input folder: ", must_exist=True, is_dir=True)
    from_ext = input("Find files with extension (for example pdf): ").strip().lower().lstrip(".")
    target = input("Convert to extension: ").strip().lower().lstrip(".")
    output_dir = prompt_path("Output folder: ", must_exist=False, is_dir=True, default=Path.cwd())
    recursive = prompt_bool("Search subfolders too?", default=True)
    preserve_structure = prompt_bool("Preserve subfolder structure?", default=True)
    ocr = prompt_bool("Use OCR for PDFs?", default=False)
    ocr_lang = input("OCR language code [eng]: ").strip() if ocr else ""
    ocr_lang = ocr_lang or "eng"
    mode = input("Mode [auto/native/python]: ").strip().lower() or "auto"
    sensitive = prompt_bool("Sensitive Document Mode? Use the safest visual-preserving path.", default=False)
    overwrite = prompt_bool("Overwrite existing outputs?", default=False)
    skip_existing = False if overwrite else prompt_bool("Skip existing outputs?", default=True)

    inputs = collect_directory_inputs(input_dir, from_ext, recursive)
    converted = 0
    skipped = 0
    for input_path in inputs:
        relative_parent = input_path.parent.relative_to(input_dir) if preserve_structure else Path()
        destination = output_path_for(input_path, output_dir / relative_parent, target)
        if ensure_output_allowed(destination, overwrite=overwrite, skip_existing=skip_existing):
            report = convert_one_with_options(
                input_path,
                target,
                destination,
                tools,
                ConversionOptions(ocr=ocr, ocr_lang=ocr_lang, mode=mode, sensitive=sensitive),
            )
            print(f"Converted: {input_path} -> {destination}")
            for line in format_report(report):
                print(line)
            converted += 1
        else:
            print(f"Skipped existing file: {destination}")
            skipped += 1
    print(f"Finished. Converted: {converted}, skipped: {skipped}")
    return 0


def interactive_merge_pdf() -> int:
    raw = input("PDF files to merge, separated by commas: ").strip()
    inputs = [Path(item.strip()).expanduser().resolve() for item in raw.split(",") if item.strip()]
    if len(inputs) < 2:
        print("Please provide at least two PDF files.")
        return 1
    output_path = prompt_path("Merged PDF output path: ", must_exist=False, default=(Path.cwd() / "merged.pdf"))
    overwrite = prompt_bool("Overwrite output if it exists?", default=False)
    ensure_output_allowed(output_path, overwrite=overwrite, skip_existing=False)
    merge_pdfs(inputs, output_path)
    print(f"Merged {len(inputs)} PDFs into {output_path}")
    return 0


def interactive_split_pdf() -> int:
    input_path = prompt_path("PDF file to split: ", must_exist=True, is_dir=False)
    output_dir = prompt_path("Output folder for split PDFs: ", must_exist=False, is_dir=True, default=Path.cwd())
    split_pdf(input_path, output_dir)
    print(f"Split {input_path} into single-page PDFs under {output_dir}")
    return 0


def interactive_compress_pdf() -> int:
    input_path = prompt_path("PDF file to compress: ", must_exist=True, is_dir=False)
    output_path = prompt_path(
        "Compressed PDF output path: ",
        must_exist=False,
        default=(Path.cwd() / f"{input_path.stem}-compressed.pdf"),
    )
    overwrite = prompt_bool("Overwrite output if it exists?", default=False)
    ensure_output_allowed(output_path, overwrite=overwrite, skip_existing=False)
    compress_pdf(input_path, output_path)
    print(f"Compressed {input_path} -> {output_path}")
    return 0


def interactive_extract_images() -> int:
    input_path = prompt_path("PDF file to inspect: ", must_exist=True, is_dir=False)
    output_dir = prompt_path(
        "Output folder for extracted images: ",
        must_exist=False,
        is_dir=True,
        default=Path.cwd(),
    )
    extract_pdf_images(input_path, output_dir)
    print(f"Extracted embedded images from {input_path} into {output_dir}")
    return 0


def interactive_menu(tools: ToolState) -> int:
    actions = {
        "1": ("Convert one file", lambda: interactive_convert(tools)),
        "2": ("Convert a whole folder", lambda: interactive_convert_dir(tools)),
        "3": ("Merge PDFs", interactive_merge_pdf),
        "4": ("Split a PDF", interactive_split_pdf),
        "5": ("Compress a PDF", interactive_compress_pdf),
        "6": ("Extract images from a PDF", interactive_extract_images),
        "7": ("Show supported conversions", lambda: (print_routes(), 0)[1]),
        "8": ("Run doctor check", lambda: doctor(tools)),
        "9": ("Launch GUI", lambda: launch_gui(tools)),
        "0": ("Exit", lambda: 0),
    }

    while True:
        print("\nOffline Converter Menu")
        print("======================")
        for key, (label, _) in actions.items():
            print(f"{key}. {label}")
        choice = input("Choose an option: ").strip()
        if choice == "0":
            return 0
        action = actions.get(choice)
        if not action:
            print("Invalid choice. Please pick one of the menu numbers.")
            continue
        try:
            result = action[1]()
            if result != 0:
                print(f"Finished with status {result}.")
        except (ConversionError, ValueError) as exc:
            print(f"Error: {exc}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Cross-platform offline document converter for privacy-sensitive files."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    cli_parser = subparsers.add_parser("cli", help="Launch the interactive terminal menu.")
    cli_parser.set_defaults(handler=lambda _args, tools: interactive_menu(tools))

    list_parser = subparsers.add_parser("list", help="Show supported offline conversions.")
    list_parser.set_defaults(handler=lambda _args, _tools: (print_routes(), 0)[1])

    doctor_parser = subparsers.add_parser("doctor", help="Check installed Python conversion libraries.")
    doctor_parser.set_defaults(handler=lambda _args, tools: doctor(tools))

    convert_parser = subparsers.add_parser("convert", help="Convert one or more files to another format.")
    convert_parser.add_argument("inputs", nargs="+", help="Input file path(s).")
    convert_parser.add_argument("--to", required=True, help="Target extension, for example pdf, docx, txt, or pptx.")
    convert_parser.add_argument("--output", help="Output file path for a single input file.")
    convert_parser.add_argument("--output-dir", help="Output directory for batch conversion.")
    convert_parser.add_argument("--ocr", action="store_true", help="Use OCR for scanned PDFs.")
    convert_parser.add_argument("--ocr-lang", default="eng", help="Tesseract OCR language code, for example eng or eng+hin.")
    convert_parser.add_argument(
        "--mode",
        choices=["auto", "native", "python"],
        default="auto",
        help="auto = try best-quality renderer first, native = require renderer, python = use pure-Python fallback only.",
    )
    convert_parser.add_argument(
        "--sensitive",
        action="store_true",
        help="Prefer visual-preserving conversion paths for sensitive documents like IDs, cards, and certificates.",
    )
    convert_parser.add_argument("--overwrite", action="store_true", help="Overwrite output files if they already exist.")
    convert_parser.add_argument("--skip-existing", action="store_true", help="Skip outputs that already exist.")
    convert_parser.set_defaults(handler=cmd_convert)

    convert_dir_parser = subparsers.add_parser("convert-dir", help="Convert matching files in a directory.")
    convert_dir_parser.add_argument("input_dir", help="Folder containing input files.")
    convert_dir_parser.add_argument("--from-ext", required=True, help="Source extension to scan for.")
    convert_dir_parser.add_argument("--to", required=True, help="Target extension.")
    convert_dir_parser.add_argument("--output-dir", required=True, help="Folder to place converted files into.")
    convert_dir_parser.add_argument("--recursive", action="store_true", help="Search subfolders too.")
    convert_dir_parser.add_argument("--preserve-structure", action="store_true", help="Preserve source subfolder layout.")
    convert_dir_parser.add_argument("--ocr", action="store_true", help="Use OCR for scanned PDFs.")
    convert_dir_parser.add_argument("--ocr-lang", default="eng", help="Tesseract OCR language code.")
    convert_dir_parser.add_argument(
        "--mode",
        choices=["auto", "native", "python"],
        default="auto",
        help="auto = try best-quality renderer first, native = require renderer, python = use pure-Python fallback only.",
    )
    convert_dir_parser.add_argument(
        "--sensitive",
        action="store_true",
        help="Prefer visual-preserving conversion paths for sensitive documents like IDs, cards, and certificates.",
    )
    convert_dir_parser.add_argument("--overwrite", action="store_true", help="Overwrite output files.")
    convert_dir_parser.add_argument("--skip-existing", action="store_true", help="Skip existing outputs.")
    convert_dir_parser.set_defaults(handler=cmd_convert_dir)

    merge_parser = subparsers.add_parser("merge-pdf", help="Merge multiple PDFs into one file.")
    merge_parser.add_argument("inputs", nargs="+", help="PDF files to merge.")
    merge_parser.add_argument("--output", required=True, help="Merged PDF output path.")
    merge_parser.add_argument("--overwrite", action="store_true", help="Overwrite the output file if it exists.")
    merge_parser.set_defaults(handler=cmd_merge_pdf)

    split_parser = subparsers.add_parser("split-pdf", help="Split a PDF into one file per page.")
    split_parser.add_argument("input", help="PDF file to split.")
    split_parser.add_argument("--output-dir", required=True, help="Directory for split PDFs.")
    split_parser.set_defaults(handler=cmd_split_pdf)

    compress_parser = subparsers.add_parser("compress-pdf", help="Apply stream compression to a PDF.")
    compress_parser.add_argument("input", help="PDF file to compress.")
    compress_parser.add_argument("--output", required=True, help="Compressed PDF output path.")
    compress_parser.add_argument("--overwrite", action="store_true", help="Overwrite the output file if it exists.")
    compress_parser.set_defaults(handler=cmd_compress_pdf)

    images_parser = subparsers.add_parser("extract-images", help="Extract embedded images from a PDF.")
    images_parser.add_argument("input", help="PDF file to inspect.")
    images_parser.add_argument("--output-dir", required=True, help="Directory for extracted image files.")
    images_parser.set_defaults(handler=cmd_extract_images)

    gui_parser = subparsers.add_parser("gui", help="Launch the local desktop GUI.")
    gui_parser.set_defaults(handler=cmd_gui)

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    tools = ToolState.detect()
    if not argv:
        return launch_gui(tools)
    if argv and argv[0] == "--cli":
        if len(argv) == 1:
            return interactive_menu(tools)
        argv = argv[1:]

    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.handler(args, tools)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
