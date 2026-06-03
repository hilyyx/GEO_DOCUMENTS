"""Подгонка изображений и таблиц в DOCX под область печати страницы."""

from __future__ import annotations

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.section import Section
_EMU_PER_INCH = 914400
_TWIPS_PER_INCH = 1440
_DEFAULT_PAGE_WIDTH_EMU = int(8.5 * _EMU_PER_INCH)
_DEFAULT_PAGE_HEIGHT_EMU = int(11 * _EMU_PER_INCH)
_DEFAULT_MARGIN_EMU = int(1 * _EMU_PER_INCH)
_TABLE_SAFE_WIDTH_RATIO = 0.98
_DRAWING_TABLE_IND_TWIPS = 108
_GRAPH_TABLE_COLS_TWIPS = (5478, 4777)
_LANDSCAPE_TABLE_THRESHOLD = 1.02
_ANCHOR_POSITION_TAGS = {
    qn("wp:simplePos"),
    qn("wp:positionH"),
    qn("wp:positionV"),
    qn("wp:wrapNone"),
    qn("wp:wrapSquare"),
    qn("wp:wrapTight"),
    qn("wp:wrapThrough"),
    qn("wp:wrapTopAndBottom"),
}
_SIGNATURE_MARKER = "составил"


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


def _set_section_landscape(section: Section) -> None:
    if not section.page_width or not section.page_height:
        return
    if section.page_width > section.page_height:
        return
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width


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


def _scale_drawing_extents(root, scale: float) -> None:
    for tag in (qn("wp:inline"), qn("wp:anchor")):
        for drawing in root.iter(tag):
            extent = drawing.find(qn("wp:extent"))
            if extent is not None:
                _scale_int_attr(extent, "cx", scale)
                _scale_int_attr(extent, "cy", scale)
            for pic_extent in drawing.iter(qn("a:ext")):
                _scale_int_attr(pic_extent, "cx", scale)
                _scale_int_attr(pic_extent, "cy", scale)


def _normalize_floating_layout(root) -> None:
    """Переводит плавающие объекты в поток, чтобы они не накладывались на текст/таблицы."""
    for tbl_pr in root.iter(qn("w:tblPr")):
        for tblp_pr in list(tbl_pr.findall(qn("w:tblpPr"))):
            tbl_pr.remove(tblp_pr)

    for p_pr in root.iter(qn("w:pPr")):
        for frame_pr in list(p_pr.findall(qn("w:framePr"))):
            p_pr.remove(frame_pr)

    for anchor in list(root.iter(qn("wp:anchor"))):
        for child in list(anchor):
            if child.tag in _ANCHOR_POSITION_TAGS:
                anchor.remove(child)
        for attr in list(anchor.attrib):
            if attr not in {"distT", "distB", "distL", "distR"}:
                del anchor.attrib[attr]
        anchor.tag = qn("wp:inline")


def _element_text(el) -> str:
    return "".join(text_el.text or "" for text_el in el.iter(qn("w:t")))


def _is_signature_marker_paragraph(p_el) -> bool:
    text = " ".join(_element_text(p_el).lower().split())
    return _SIGNATURE_MARKER in text and ":" in text


def _run_has_drawing(run_el) -> bool:
    drawing_tags = {qn("w:drawing"), qn("w:pict"), qn("w:object")}
    return any(el.tag in drawing_tags for el in run_el.iter())


def _drawing_runs(p_el) -> list:
    return [run_el for run_el in p_el.findall(qn("w:r")) if _run_has_drawing(run_el)]


def _paragraph_has_text(p_el) -> bool:
    return bool(_element_text(p_el).strip())


def _is_slash_name_paragraph(p_el) -> bool:
    text = " ".join(_element_text(p_el).split())
    return 3 <= len(text) <= 80 and text.startswith("/") and text.endswith("/")


def _append_tab_run(p_el) -> None:
    run = OxmlElement("w:r")
    tab = OxmlElement("w:tab")
    run.append(tab)
    p_el.append(run)


def _move_signature_runs(target_p, source_p, body) -> bool:
    runs = _drawing_runs(source_p)
    if not runs:
        return False
    _append_tab_run(target_p)
    for run in runs:
        target_p.append(run)
    if _is_empty_paragraph(source_p):
        body.remove(source_p)
    return True


def _move_slash_name_to_signature_line(target_p, source_p, body) -> bool:
    if not _is_slash_name_paragraph(source_p):
        return False
    _append_tab_run(target_p)
    for run in list(source_p.findall(qn("w:r"))):
        target_p.append(run)
    if _is_empty_paragraph(source_p):
        body.remove(source_p)
    return True


def _place_signatures_next_to_marker(doc: Document) -> None:
    """Ставит картинку-подпись справа от строки вида "Составил: ...".

    В шаблонах подпись часто хранится отдельным плавающим рисунком рядом со строкой
    исполнителя. После слияния такой рисунок может оказаться поверх таблицы, поэтому
    переносим ближайший рисунок в сам абзац "Составил:".
    """
    body = doc.element.body
    children = _content_children(body)

    for idx, child in enumerate(list(children)):
        if child.tag != qn("w:p") or not _is_signature_marker_paragraph(child):
            continue
        if _drawing_runs(child):
            continue

        moved = False
        for prev_idx in range(idx - 1, max(-1, idx - 4), -1):
            candidate = children[prev_idx]
            if candidate.tag == qn("w:tbl"):
                continue
            if candidate.tag == qn("w:p"):
                if _move_signature_runs(child, candidate, body):
                    moved = True
                    break
                if _paragraph_has_text(candidate):
                    break
                continue
            break

        if not moved:
            for next_idx in range(idx + 1, min(len(children), idx + 5)):
                candidate = children[next_idx]
                if candidate.tag != qn("w:p"):
                    break
                if _move_signature_runs(child, candidate, body):
                    following_idx = next_idx + 1
                    if following_idx < len(children):
                        _move_slash_name_to_signature_line(child, children[following_idx], body)
                    break
                if _paragraph_has_text(candidate):
                    break


def _move_signature_blocks_under_tables(doc: Document) -> None:
    """Если строка "Составил:" стоит перед таблицей, переносит её сразу под таблицу."""
    body = doc.element.body

    for child in list(_content_children(body)):
        if child.tag != qn("w:p") or not _is_signature_marker_paragraph(child):
            continue

        children = _content_children(body)
        try:
            idx = children.index(child)
        except ValueError:
            continue

        target_table = None
        for next_idx in range(idx + 1, min(len(children), idx + 8)):
            candidate = children[next_idx]
            if candidate.tag == qn("w:tbl"):
                target_table = candidate
                continue
            if candidate.tag == qn("w:p") and _is_empty_paragraph(candidate):
                continue
            break

        if target_table is not None:
            body.remove(child)
            target_table.addnext(child)


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


def _remove_children(parent, child_tag: str) -> None:
    for child in list(parent.findall(child_tag)):
        parent.remove(child)


def _table_has_drawings(tbl_el) -> bool:
    return any(
        True
        for _ in (
            list(tbl_el.iter(qn("w:drawing")))
            + list(tbl_el.iter(qn("w:pict")))
            + list(tbl_el.iter(qn("w:object")))
        )
    )


def _is_graph_table(tbl_el) -> bool:
    text = " ".join(_element_text(tbl_el).lower().split())
    drawings = sum(1 for _ in tbl_el.iter(qn("w:drawing"))) + sum(
        1 for _ in tbl_el.iter(qn("w:pict"))
    )
    return drawings >= 2 and (
        "консолидированно" in text or "график зависимости" in text
    )


def _set_graph_table_reference_geometry(tbl_el) -> None:
    grid = tbl_el.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        if len(cols) == len(_GRAPH_TABLE_COLS_TWIPS):
            for col, width in zip(cols, _GRAPH_TABLE_COLS_TWIPS):
                col.set(qn("w:w"), str(width))

    first_row = tbl_el.find(qn("w:tr"))
    if first_row is None:
        return
    cells = first_row.findall(qn("w:tc"))
    if len(cells) != len(_GRAPH_TABLE_COLS_TWIPS):
        return
    for cell, width in zip(cells, _GRAPH_TABLE_COLS_TWIPS):
        tc_pr = cell.find(qn("w:tcPr"))
        if tc_pr is None:
            continue
        tc_w = tc_pr.find(qn("w:tcW"))
        if tc_w is not None:
            tc_w.set(qn("w:type"), "dxa")
            tc_w.set(qn("w:w"), str(width))


def _normalize_table_readability(tbl_el, max_width_twips: int) -> None:
    has_drawings = _table_has_drawings(tbl_el)
    tbl_pr = tbl_el.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl_el.insert(0, tbl_pr)

    if has_drawings:
        tbl_ind = tbl_pr.find(qn("w:tblInd"))
        if tbl_ind is None:
            tbl_ind = OxmlElement("w:tblInd")
            tbl_pr.append(tbl_ind)
        tbl_ind.set(qn("w:type"), "dxa")
        tbl_ind.set(qn("w:w"), str(_DRAWING_TABLE_IND_TWIPS))
    else:
        # Отступ таблицы часто даёт визуальную обрезку справа/слева после экспорта в PDF.
        _remove_children(tbl_pr, qn("w:tblInd"))

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    if has_drawings:
        tbl_w.set(qn("w:type"), "auto")
        tbl_w.set(qn("w:w"), "0")
    else:
        tbl_w.set(qn("w:type"), "dxa")
        tbl_w.set(qn("w:w"), str(max_width_twips))

    tbl_layout = tbl_pr.find(qn("w:tblLayout"))
    if tbl_layout is None:
        tbl_layout = OxmlElement("w:tblLayout")
        tbl_pr.append(tbl_layout)
    tbl_layout.set(qn("w:type"), "fixed" if has_drawings else "autofit")

    for tr in tbl_el.iter(qn("w:tr")):
        tr_pr = tr.find(qn("w:trPr"))
        if tr_pr is None:
            continue
        for tr_height in tr_pr.findall(qn("w:trHeight")):
            # "exact" режет многострочные/повёрнутые подписи в ячейках.
            if tr_height.get(qn("w:hRule")) == "exact":
                tr_height.set(qn("w:hRule"), "atLeast")

    for tc_pr in tbl_el.iter(qn("w:tcPr")):
        _remove_children(tc_pr, qn("w:noWrap"))


def _scale_tables(root, max_width_twips: int) -> None:
    for tbl_el in root.iter(qn("w:tbl")):
        _normalize_table_readability(tbl_el, max_width_twips)
        if _is_graph_table(tbl_el):
            _set_graph_table_reference_geometry(tbl_el)
            continue
        if _table_has_drawings(tbl_el):
            continue
        width = _table_width_twips(tbl_el, content_width_twips=max_width_twips)
        if width is None:
            continue
        if width <= max_width_twips:
            continue
        scale = max_width_twips / width
        _scale_table_columns(tbl_el, scale)


def _make_wide_table_sections_landscape(doc: Document) -> None:
    if not doc.sections:
        return

    content_widths = [_section_content_emu(section)[0] for section in doc.sections]
    if not content_widths:
        return
    max_current_width_twips = int(max(content_widths) / _EMU_PER_INCH * _TWIPS_PER_INCH)

    widest_table = 0
    for tbl_el in doc.element.body.iter(qn("w:tbl")):
        width = _table_width_twips(tbl_el, content_width_twips=max_current_width_twips)
        if width is not None:
            widest_table = max(widest_table, width)

    if widest_table <= max_current_width_twips * _LANDSCAPE_TABLE_THRESHOLD:
        return

    for section in doc.sections:
        _set_section_landscape(section)


def _paragraph_has_visible_content(p_el) -> bool:
    for text_el in p_el.iter(qn("w:t")):
        if text_el.text and text_el.text.strip():
            return True
    visible_tags = {
        qn("w:drawing"),
        qn("w:pict"),
        qn("w:object"),
        qn("w:tbl"),
    }
    return any(el.tag in visible_tags for el in p_el.iter())


def _paragraph_has_page_break(p_el) -> bool:
    for br in p_el.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _is_empty_paragraph(p_el) -> bool:
    return p_el.tag == qn("w:p") and not _paragraph_has_visible_content(p_el)


def _is_page_break_only_paragraph(p_el) -> bool:
    return (
        p_el.tag == qn("w:p")
        and _paragraph_has_page_break(p_el)
        and not _paragraph_has_visible_content(p_el)
    )


def _is_section_break_only_paragraph(p_el) -> bool:
    if p_el.tag != qn("w:p") or _paragraph_has_visible_content(p_el):
        return False
    p_pr = p_el.find(qn("w:pPr"))
    return p_pr is not None and p_pr.find(qn("w:sectPr")) is not None


def _is_trailing_artifact_paragraph(p_el) -> bool:
    return (
        _is_empty_paragraph(p_el)
        or _is_page_break_only_paragraph(p_el)
        or _is_section_break_only_paragraph(p_el)
    )


def _content_children(body) -> list:
    return [child for child in body if child.tag != qn("w:sectPr")]


def _strip_trailing_artifacts(body) -> None:
    while True:
        children = _content_children(body)
        if not children or not _is_trailing_artifact_paragraph(children[-1]):
            break
        body.remove(children[-1])


def _strip_page_break_from_last_content_paragraph(body) -> None:
    children = _content_children(body)
    if not children:
        return
    last = children[-1]
    if last.tag != qn("w:p") or not _paragraph_has_visible_content(last):
        return
    if not _paragraph_has_page_break(last):
        return
    for br in list(last.iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page":
            parent = br.getparent()
            if parent is not None:
                parent.remove(br)


def document_ends_with_page_forcing_break(doc: Document) -> bool:
    """True, если документ уже заканчивается разрывом страницы или секции."""
    children = _content_children(doc.element.body)
    if not children:
        return False
    last = children[-1]
    return _is_page_break_only_paragraph(last) or _is_section_break_only_paragraph(last)


def remove_empty_pages(doc: Document) -> None:
    """Убирает пустые страницы: лишние разрывы, пустые абзацы, хвостовые артефакты."""
    body = doc.element.body
    seen_content = False
    previous_page_break = False

    for child in list(_content_children(body)):
        if _is_page_break_only_paragraph(child):
            if not seen_content or previous_page_break:
                body.remove(child)
                continue
            previous_page_break = True
            continue

        if _is_empty_paragraph(child):
            if not seen_content or previous_page_break:
                body.remove(child)
            continue

        seen_content = True
        previous_page_break = False

    _strip_page_break_from_last_content_paragraph(body)
    _strip_trailing_artifacts(body)


def fit_document_content(doc: Document) -> None:
    """Уменьшает слишком широкие таблицы и рисунки под поля секций документа."""
    if not doc.sections:
        return
    _make_wide_table_sections_landscape(doc)
    widths: list[int] = []
    heights: list[int] = []
    for section in doc.sections:
        w, h = _section_content_emu(section)
        widths.append(w)
        heights.append(h)
    # Сохраняем доступную ширину альбомных секций: широкие листы не должны
    # принудительно ужиматься под портретную страницу.
    img_max_w, img_max_h = max(widths), max(heights)
    tbl_max_w = max(widths)
    max_width_twips = int(tbl_max_w / _EMU_PER_INCH * _TWIPS_PER_INCH * _TABLE_SAFE_WIDTH_RATIO)

    root = doc.element.body
    _place_signatures_next_to_marker(doc)
    _move_signature_blocks_under_tables(doc)
    _normalize_floating_layout(root)
    _scale_drawings(root, img_max_w, img_max_h)
    _scale_tables(root, max_width_twips)
    remove_empty_pages(doc)
