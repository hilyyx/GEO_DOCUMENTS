from __future__ import annotations

import io
import shutil
import traceback
from pathlib import Path

import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Inches
from docxcompose.composer import Composer

from geo_documents.libreoffice import doc_to_docx, docx_to_pdf, find_soffice


def _prepend_heading(doc: Document, text: str) -> None:
    h = doc.add_heading(text, level=2)
    el = h._element
    body = doc.element.body
    body.remove(el)
    body.insert(0, el)


def _append_page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _append_pdf_as_images(
    doc: Document,
    pdf_path: Path,
    *,
    dpi: int = 120,
    max_width_inches: float = 6.5,
) -> None:
    src = fitz.open(pdf_path)
    try:
        for i in range(len(src)):
            page = src[i]
            pix = page.get_pixmap(dpi=dpi)
            bio = io.BytesIO(pix.tobytes("png"))
            par = doc.add_paragraph()
            run = par.add_run()
            run.add_picture(bio, width=Inches(max_width_inches))
    finally:
        src.close()


def merge_to_docx_and_pdf(
    paths: list[Path],
    output_docx: Path,
    output_pdf: Path | None,
    *,
    page_break_between_parts: bool = True,
    insert_titles: bool = False,
    pdf_render_dpi: int = 120,
    libreoffice_executable: str | None = None,
) -> tuple[list[str], list[str]]:
    """
    Склеивает файлы по порядку `paths`.
    Возвращает (warnings, errors). errors непустой — часть шагов не выполнена.
    """
    warnings: list[str] = []
    errors: list[str] = []
    soffice = find_soffice(preferred=libreoffice_executable)

    work_items: list[tuple[Path, str]] = []
    temp_to_delete: list[Path] = []

    for p in paths:
        p = Path(p)
        if not p.is_file():
            warnings.append(f"Пропуск (файл не найден): {p}")
            continue
        ext = p.suffix.lower()
        if ext == ".doc":
            if not soffice:
                errors.append(
                    f"Нужен LibreOffice для .doc: {p.name} "
                    "(укажите путь к soffice.exe в поле ниже или установите LibreOffice)."
                )
                continue
            try:
                cx = doc_to_docx(soffice, p)
                temp_to_delete.append(cx)
                work_items.append((cx, p.name))
            except Exception as e:
                errors.append(f"Конвертация .doc → docx не удалась ({p.name}): {e}")
            continue
        if ext == ".docx":
            work_items.append((p, p.name))
            continue
        if ext == ".pdf":
            work_items.append((p, p.name))
            continue
        warnings.append(f"Пропуск (неподдерживаемый тип): {p.name}")

    if not work_items:
        errors.append("Нет ни одного поддерживаемого файла для склейки.")
        return warnings, errors

    merged: Document | None = None
    composer: Composer | None = None

    for idx, (src_path, display_name) in enumerate(work_items):
        ext = src_path.suffix.lower()

        if idx > 0 and page_break_between_parts and merged is not None:
            _append_page_break(merged)

        if ext == ".pdf":
            if merged is None:
                merged = Document()
                composer = None
            if insert_titles:
                merged.add_heading(display_name, level=2)
            _append_pdf_as_images(merged, src_path, dpi=pdf_render_dpi)
            continue

        if ext == ".docx":
            doc = Document(str(src_path))
            if insert_titles:
                _prepend_heading(doc, display_name)
            if merged is None:
                merged = doc
                composer = Composer(merged)
            else:
                if composer is None:
                    composer = Composer(merged)
                composer.append(doc)
            continue

    assert merged is not None

    output_docx = Path(output_docx)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    try:
        merged.save(str(output_docx))
    except Exception as e:
        errors.append(f"Сохранение DOCX не удалось: {e}\n{traceback.format_exc()}")
        for t in temp_to_delete:
            t.unlink(missing_ok=True)
        return warnings, errors

    if output_pdf is not None:
        output_pdf = Path(output_pdf)
        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        if not soffice:
            warnings.append(
                "LibreOffice (soffice) не найден — PDF не создан. "
                "Укажите полный путь к soffice.exe в настройках окна, "
                "переменную LIBREOFFICE_EXECUTABLE или установите LibreOffice."
            )
        else:
            try:
                docx_to_pdf(soffice, output_docx, output_pdf.parent)
                lo_out = output_pdf.parent / f"{output_docx.stem}.pdf"
                if lo_out.is_file():
                    if lo_out.resolve() != output_pdf.resolve():
                        shutil.move(str(lo_out), str(output_pdf))
                else:
                    errors.append(f"LibreOffice не создал PDF: {lo_out}")
            except Exception as e:
                errors.append(f"Экспорт PDF не удался: {e}")

    for t in temp_to_delete:
        try:
            t.unlink(missing_ok=True)
            if t.parent.is_dir() and not any(t.parent.iterdir()):
                t.parent.rmdir()
        except OSError:
            warnings.append(f"Не удалось удалить временный файл: {t}")

    return warnings, errors
