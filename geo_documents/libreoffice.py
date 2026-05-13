from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def find_soffice(*, preferred: str | Path | None = None) -> str | None:
    """Ищет soffice.exe: сначала явный путь из UI, затем переменные окружения, стандартные каталоги, PATH."""
    if preferred:
        p = Path(str(preferred).strip().strip('"'))
        if p.is_file():
            return str(p.resolve())
    env = os.environ.get("LIBREOFFICE_EXECUTABLE") or os.environ.get("SOFFICE_PATH")
    if env and Path(env).is_file():
        return env
    for candidate in (
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ):
        if candidate.is_file():
            return str(candidate)
    which = shutil.which("soffice") or shutil.which("soffice.exe")
    return which


def convert_to(
    soffice: str,
    source: str | Path,
    outdir: str | Path,
    target_ext: str,
    timeout_sec: int = 120,
) -> Path:
    """Конвертирует один файл в outdir (docx, pdf, ...). Возвращает путь к результату."""
    source = Path(source).resolve()
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        target_ext,
        "--outdir",
        str(outdir),
        str(source),
    ]
    subprocess.run(cmd, check=True, timeout=timeout_sec, capture_output=True)
    out_name = source.with_suffix("." + target_ext.lstrip(".")).name
    out_path = outdir / out_name
    if not out_path.is_file():
        raise FileNotFoundError(f"LibreOffice не создал файл: {out_path}")
    return out_path


def doc_to_docx(soffice: str, doc_path: str | Path) -> Path:
    tmp = Path(tempfile.mkdtemp(prefix="geo_doc_"))
    try:
        return convert_to(soffice, doc_path, tmp, "docx")
    except Exception:
        for p in tmp.glob("*"):
            p.unlink(missing_ok=True)
        tmp.rmdir()
        raise


def docx_to_pdf(soffice: str, docx_path: str | Path, outdir: str | Path | None = None) -> Path:
    docx_path = Path(docx_path).resolve()
    outdir = Path(outdir or docx_path.parent)
    return convert_to(soffice, docx_path, outdir, "pdf")
