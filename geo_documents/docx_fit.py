"""Подгонка изображений и таблиц в DOCX под область печати страницы."""

from __future__ import annotations

from docx import Document
from docx.oxml.ns import qn
from docx.section import Section
_EMU_PER_INCH = 914400
_TWIPS_PER_INCH = 1440
_DEFAULT_PAGE_WIDTH_EMU = int(8.5 * _EMU_PER_INCH)
_DEFAULT_PAGE_HEIGHT_EMU = int(11 * _EMU_PER_INCH)
_DEFAULT_MARGIN_EMU = int(1 * _EMU_PER_INCH)


def _section_content_emu(section: Section) -> tuple[int, int]:
    page_width = section.page_width or _DEFAULT_PAGE_WIDTH_EMU
    page_height = section.page_height or _DEFAULT_PAGE_HEIGHT_EMU
    left_margin = section.left_margin or _DEFAULT_MARGIN_EMU
    right_margin = section.right_margin or _DEFAULT_MARGIN_EMU
    top_margin = section.top_margin or _DEFAULT_MARGIN_EMU
    bottom_margin = section.bottom_margin or _DEFAULT_MARGIN_EMU

    w = page_width - left_margin - right_margin
    h = page_height - top_margin - bottom_margin
    if w <= 0:
        w = _DEFAULT_PAGE_WIDTH_EMU - 2 * _DEFAULT_MARGIN_EMU
    if h <= 0:
        h = _DEFAULT_PAGE_HEIGHT_EMU - 2 * _DEFAULT_MARGIN_EMU
    return int(w), int(h)


def section_content_inches(section: Section) -> tuple[float, float]:
    w, h = _section_content_emu(section)
    return w / _EMU_PER_INCH, h / _EMU_PER_INCH


def _scale_int_attr(el, attr_qn: str, scale: float) -> None:
    raw = el.get(attr_qn)
    if raw is None:
        return
    try:
        val = int(raw)
    except (TypeError, ValueError):
        return
    el.set(attr_qn, str(max(1, int(val * scale))))


def _scale_drawings(root, max_cx: int, max_cy: int) -> None:
    for tag in (qn("wp:inline"), qn("wp:anchor")):
        for drawing in root.iter(tag):
            extent = drawing.find(qn("wp:extent"))
            if extent is None:
                continue
            cx_attr, cy_attr = "cx", "cy"
            try:
                cx = int(extent.get(cx_attr, 0))
                cy = int(extent.get(cy_attr, 0))
            except (TypeError, ValueError):
                continue
            if cx <= 0 or cy <= 0:
                continue
            scale = min(max_cx / cx, max_cy / cy, 1.0)
            if scale < 1.0:
                _scale_int_attr(extent, cx_attr, scale)
                _scale_int_attr(extent, cy_attr, scale)
                for pic_extent in drawing.iter(qn("a:ext")):
                    _scale_int_attr(pic_extent, cx_attr, scale)
                    _scale_int_attr(pic_extent, cy_attr, scale)


def _table_width_twips(tbl_el, *, content_width_twips: int) -> int | None:
    tbl_grid = tbl_el.find(qn("w:tblGrid"))
    if tbl_grid is not None:
        total = 0
        for col in tbl_grid.findall(qn("w:gridCol")):
            w = col.get(qn("w:w"))
            if w:
                try:
                    total += int(w)
                except ValueError:
                    pass
        if total > 0:
            return total

    tbl_pr = tbl_el.find(qn("w:tblPr"))
    if tbl_pr is None:
        return None
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        return None
    w_type = tbl_w.get(qn("w:type"), "auto")
    if w_type == "dxa":
        try:
            return int(tbl_w.get(qn("w:w"), 0))
        except ValueError:
            return None
    if w_type == "pct":
        try:
            pct = int(tbl_w.get(qn("w:w"), 5000))
        except ValueError:
            return None
        return int(content_width_twips * pct / 5000)
    return None


def _scale_table_columns(tbl_el, scale: float) -> None:
    tbl_grid = tbl_el.find(qn("w:tblGrid"))
    if tbl_grid is not None:
        for col in tbl_grid.findall(qn("w:gridCol")):
            _scale_int_attr(col, qn("w:w"), scale)

    for tc in tbl_el.iter(qn("w:tc")):
        tc_pr = tc.find(qn("w:tcPr"))
        if tc_pr is None:
            continue
        tc_w = tc_pr.find(qn("w:tcW"))
        if tc_w is not None and tc_w.get(qn("w:type")) == "dxa":
            _scale_int_attr(tc_w, qn("w:w"), scale)

    tbl_pr = tbl_el.find(qn("w:tblPr"))
    if tbl_pr is not None:
        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is not None and tbl_w.get(qn("w:type")) == "dxa":
            _scale_int_attr(tbl_w, qn("w:w"), scale)


def _scale_tables(root, max_width_twips: int) -> None:
    for tbl_el in root.iter(qn("w:tbl")):
        width = _table_width_twips(tbl_el, content_width_twips=max_width_twips)
        if width is None or width <= max_width_twips:
            continue
        scale = max_width_twips / width
        _scale_table_columns(tbl_el, scale)


def fit_document_content(doc: Document) -> None:
    """Уменьшает слишком широкие таблицы и рисунки под поля секций документа."""
    if not doc.sections:
        return
    widths: list[int] = []
    heights: list[int] = []
    for section in doc.sections:
        w, h = _section_content_emu(section)
        widths.append(w)
        heights.append(h)
    # Рисунки должны помещаться и в портрет, и в альбомную секцию.
    img_max_w, img_max_h = min(widths), min(heights)
    # Таблицы в альбомной ориентации могут быть шире — берём максимальную ширину поля.
    tbl_max_w = max(widths)
    max_width_twips = int(tbl_max_w / _EMU_PER_INCH * _TWIPS_PER_INCH)

    root = doc.element.body
    _scale_drawings(root, img_max_w, img_max_h)
    _scale_tables(root, max_width_twips)
