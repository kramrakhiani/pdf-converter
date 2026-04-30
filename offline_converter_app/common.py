from __future__ import annotations

import hashlib
import re
import shutil
import platform
from dataclasses import dataclass
from pathlib import Path


TEXT_INPUTS = {"txt", "md", "markdown", "html", "htm", "rtf"}
WORD_INPUTS = {"docx", "doc"}
PRESENTATION_INPUTS = {"pptx", "ppt"}
IMAGE_INPUTS = {"png", "jpg", "jpeg", "bmp", "gif", "tiff", "webp", "heic", "heif", "svg", "ico", "jfif"}
SPREADSHEET_INPUTS = {"xlsx", "xls"}

MAGIC_BYTES = {
    "pdf": b"%PDF",
    "docx": b"PK\x03\x04",
    "pptx": b"PK\x03\x04",
    "xlsx": b"PK\x03\x04",
    "doc": b"\xd0\xcf\x11\xe0",
    "ppt": b"\xd0\xcf\x11\xe0",
    "xls": b"\xd0\xcf\x11\xe0",
    "png": b"\x89PNG",
    "jpg": b"\xff\xd8\xff",
    "jpeg": b"\xff\xd8\xff",
    "gif": b"GIF8",
    "bmp": b"BM",
    "tiff": b"II\x2a\x00",
    "webp": b"RIFF",
    "heic": b"ftyp",
    "heif": b"ftyp",
}


class ConversionError(RuntimeError):
    pass


@dataclass(frozen=True)
class ToolState:
    tesseract: str | None
    libreoffice: str | None
    osascript: str | None
    powershell: str | None
    mac_word_app: str | None
    mac_powerpoint_app: str | None
    platform: str

    @classmethod
    def detect(cls) -> "ToolState":
        mac_word_app = None
        mac_powerpoint_app = None
        if platform.system() == "Darwin":
            for candidate in (
                Path("/Applications/Microsoft Word.app"),
                Path.home() / "Applications/Microsoft Word.app",
            ):
                if candidate.exists():
                    mac_word_app = str(candidate)
                    break
            for candidate in (
                Path("/Applications/Microsoft PowerPoint.app"),
                Path.home() / "Applications/Microsoft PowerPoint.app",
            ):
                if candidate.exists():
                    mac_powerpoint_app = str(candidate)
                    break
        return cls(
            tesseract=shutil.which("tesseract"),
            libreoffice=shutil.which("libreoffice") or shutil.which("soffice"),
            osascript=shutil.which("osascript"),
            powershell=shutil.which("powershell") or shutil.which("pwsh"),
            mac_word_app=mac_word_app,
            mac_powerpoint_app=mac_powerpoint_app,
            platform=platform.system(),
        )


@dataclass(frozen=True)
class ConversionReport:
    engine: str
    fidelity: str
    notes: tuple[str, ...] = ()
    verified: bool = False
    checksum: str = ""


@dataclass(frozen=True)
class ConversionOptions:
    ocr: bool = False
    ocr_lang: str = "eng"
    mode: str = "auto"
    sensitive: bool = False


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def ensure_file(path: Path) -> None:
    if not path.is_file():
        raise ConversionError(f"Input file not found: {path}")


def normalize_ext(path: Path) -> str:
    return path.suffix.lower().lstrip(".")


def output_path_for(input_path: Path, output_dir: Path, target_ext: str) -> Path:
    return output_dir / f"{input_path.stem}.{target_ext}"


def ensure_output_allowed(output_path: Path, *, overwrite: bool, skip_existing: bool) -> bool:
    if output_path.exists():
        if skip_existing:
            return False
        if not overwrite:
            raise ConversionError(
                f"Output already exists: {output_path}. Use --overwrite or --skip-existing."
            )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return True


def collect_directory_inputs(directory: Path, from_ext: str, recursive: bool) -> list[Path]:
    if not directory.is_dir():
        raise ConversionError(f"Input directory not found: {directory}")
    pattern = f"*.{from_ext.lower().lstrip('.')}"
    matches = directory.rglob(pattern) if recursive else directory.glob(pattern)
    files = sorted(path.resolve() for path in matches if path.is_file())
    if not files:
        raise ConversionError(f"No files matching {pattern} were found under {directory}")
    return files


def clean_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def validate_file_type(path: Path) -> str:
    ext = normalize_ext(path)
    with path.open("rb") as f:
        header = f.read(12)
    for file_type, magic in MAGIC_BYTES.items():
        if header.startswith(magic):
            return file_type
    return ext


def sanitize_path(path: Path, base_dir: Path) -> Path:
    resolved = path.resolve()
    base_resolved = base_dir.resolve()
    if not str(resolved).startswith(str(base_resolved)):
        raise ConversionError(f"Path traversal attempt detected: {path}")
    return resolved


def validate_output_path(output_path: Path) -> None:
    dangerous_chars = ["..", "~", "$"]
    for char in dangerous_chars:
        if char in str(output_path):
            raise ConversionError(f"Dangerous path component detected: {char}")
