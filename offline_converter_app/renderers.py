from __future__ import annotations

import subprocess
from pathlib import Path

from .common import ConversionError, ToolState


def run_command(args: list[str]) -> None:
    try:
        subprocess.run(args, check=True, text=True, capture_output=True)
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.strip() if exc.stderr else ""
        stdout = exc.stdout.strip() if exc.stdout else ""
        detail = stderr or stdout or str(exc)
        raise ConversionError(detail) from exc


def libreoffice_to_pdf(input_path: Path, output_path: Path, soffice: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    run_command([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_path.parent), str(input_path)])
    produced = output_path.parent / f"{input_path.stem}.pdf"
    if not produced.exists():
        raise ConversionError("LibreOffice did not produce the expected PDF.")
    if produced != output_path:
        produced.replace(output_path)


def mac_word_to_pdf(input_path: Path, output_path: Path, osascript: str) -> None:
    script = f'''
tell application "Microsoft Word"
    activate
    open POSIX file "{input_path}"
    set docRef to active document
    save as docRef file name "{output_path}" file format format PDF
    close docRef saving no
end tell
'''
    run_command([osascript, "-e", script])


def mac_powerpoint_to_pdf(input_path: Path, output_path: Path, osascript: str) -> None:
    script = f'''
tell application "Microsoft PowerPoint"
    activate
    open POSIX file "{input_path}"
    save active presentation in "{output_path}" as save as PDF
    close active presentation
end tell
'''
    run_command([osascript, "-e", script])


def windows_word_to_pdf(input_path: Path, output_path: Path, powershell: str) -> None:
    src = str(input_path).replace("'", "''")
    dst = str(output_path).replace("'", "''")
    command = (
        "$word = New-Object -ComObject Word.Application; "
        "$word.Visible = $false; "
        f"$doc = $word.Documents.Open('{src}'); "
        f"$doc.SaveAs([ref] '{dst}', [ref] 17); "
        "$doc.Close(); "
        "$word.Quit();"
    )
    run_command([powershell, "-NoProfile", "-Command", command])


def windows_powerpoint_to_pdf(input_path: Path, output_path: Path, powershell: str) -> None:
    src = str(input_path).replace("'", "''")
    dst = str(output_path).replace("'", "''")
    command = (
        "$ppt = New-Object -ComObject PowerPoint.Application; "
        f"$presentation = $ppt.Presentations.Open('{src}', $true, $false, $false); "
        f"$presentation.SaveAs('{dst}', 32); "
        "$presentation.Close(); "
        "$ppt.Quit();"
    )
    run_command([powershell, "-NoProfile", "-Command", command])


def renderer_candidates(source_ext: str, tools: ToolState) -> list[str]:
    candidates: list[str] = []
    if tools.libreoffice and source_ext in {"docx", "pptx"}:
        candidates.append("LibreOffice")
    if tools.platform == "Darwin" and tools.osascript:
        if source_ext == "docx" and tools.mac_word_app:
            candidates.append("Microsoft Word (macOS)")
        elif source_ext == "pptx" and tools.mac_powerpoint_app:
            candidates.append("Microsoft PowerPoint (macOS)")
    if tools.platform == "Windows" and tools.powershell:
        if source_ext == "docx":
            candidates.append("Microsoft Word (Windows)")
        elif source_ext == "pptx":
            candidates.append("Microsoft PowerPoint (Windows)")
    return candidates


def render_office_to_pdf(input_path: Path, output_path: Path, tools: ToolState) -> str:
    source_ext = input_path.suffix.lower().lstrip(".")
    errors: list[str] = []

    if tools.libreoffice and source_ext in {"docx", "pptx"}:
        try:
            libreoffice_to_pdf(input_path, output_path, tools.libreoffice)
            return "LibreOffice"
        except ConversionError as exc:
            errors.append(f"LibreOffice: {exc}")

    if tools.platform == "Darwin" and tools.osascript:
        try:
            if source_ext == "docx" and tools.mac_word_app:
                mac_word_to_pdf(input_path, output_path, tools.osascript)
                return "Microsoft Word (macOS)"
            if source_ext == "pptx" and tools.mac_powerpoint_app:
                mac_powerpoint_to_pdf(input_path, output_path, tools.osascript)
                return "Microsoft PowerPoint (macOS)"
        except ConversionError as exc:
            errors.append(f"macOS Office automation: {exc}")

    if tools.platform == "Windows" and tools.powershell:
        try:
            if source_ext == "docx":
                windows_word_to_pdf(input_path, output_path, tools.powershell)
                return "Microsoft Word (Windows)"
            if source_ext == "pptx":
                windows_powerpoint_to_pdf(input_path, output_path, tools.powershell)
                return "Microsoft PowerPoint (Windows)"
        except ConversionError as exc:
            errors.append(f"Windows Office automation: {exc}")

    if errors:
        raise ConversionError(" ; ".join(errors))
    raise ConversionError("No renderer is available for high-fidelity Office-to-PDF export.")
