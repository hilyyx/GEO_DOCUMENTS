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

from geo_documents.docx_fit import fit_document_content, section_content_inches
from geo_documents.libreoffice import docx_to_pdf, find_soffice


def _prepend_heading(doc: Document, text: str) -> None:
    h = doc.add_heading(text, level=2)
    el = h._element
    body = doc.element.body
    body.remove(el)
    body.insert(0, el)


def _append_page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _pdf_content_inches(doc: Document) -> tuple[float, float]:
    """Минимальная область печати — PDF-страницы вставляются на любую секцию."""
    widths: list[float] = []
    heights: list[float] = []
    for section in doc.sections:
        w, h = section_content_inches(section)
        widths.append(w)
        heights.append(h)
    if not widths:
        return 6.0, 9.0
    return min(widths), min(heights)


def _fit_picture_inches(
    width_px: int,
    height_px: int,
    *,
    dpi: int,
    max_width_inches: float,
    max_height_inches: float,
) -> tuple[float, float]:
    w_in = width_px / dpi
    h_in = height_px / dpi
    if w_in <= 0 or h_in <= 0:
        return max_width_inches, max_height_inches
    scale = min(max_width_inches / w_in, max_height_inches / h_in, 1.0)
    return w_in * scale, h_in * scale


def _pdf_render_rect(page: fitz.Page) -> fitz.Rect:
    """Область рендера PDF: не обрезать контент из-за узкого CropBox."""
    rect = page.rect | page.mediabox
    try:
        bound = page.bound()
        if bound.width > 0 and bound.height > 0:
            rect |= bound
    except Exception:
        pass
    return rect


def _append_pdf_as_images(
    doc: Document,
    pdf_path: Path,
    *,
    dpi: int = 120,
) -> None:
    max_w_in, max_h_in = _pdf_content_inches(doc)
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)

    src = fitz.open(pdf_path)
    try:
        for i in range(len(src)):
            page = src[i]
            pix = page.get_pixmap(matrix=matrix, clip=_pdf_render_rect(page), alpha=False)
            w_in, h_in = _fit_picture_inches(
                pix.width,
                pix.height,
                dpi=dpi,
                max_width_inches=max_w_in,
                max_height_inches=max_h_in,
            )
            bio = io.BytesIO(pix.tobytes("png"))
            bio.seek(0)
            par = doc.add_paragraph()
            run = par.add_run()
            run.add_picture(bio, width=Inches(w_in), height=Inches(h_in))
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
            warnings.append(f"Пропуск .doc без чтения: {p.name}")
            continue
        if ext == ".docx":
            work_items.append((p, p.name))
            continue
        if ext == ".pdf":
            work_items.append((p, p.name))
            continue
        warnings.append(f"Пропуск (неподдерживаемый тип): {p.name}")

    if not work_items:
        errors.append("Нет ни одного поддерживаемого файла для склейки (.doc пропускаются).")
        return warnings, errors

    merged: Document | None = None
    composer: Composer | None = None

    output_docx = Path(output_docx)
    try:
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
                fit_document_content(doc)
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

        if merged is None:
            errors.append("Не удалось сформировать документ (неизвестная причина).")
            for t in temp_to_delete:
                t.unlink(missing_ok=True)
            return warnings, errors

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        fit_document_content(merged)
        merged.save(str(output_docx))
    except Exception as e:
        errors.append(f"Ошибка при склейке или сохранении DOCX: {e}\n{traceback.format_exc()}")
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
