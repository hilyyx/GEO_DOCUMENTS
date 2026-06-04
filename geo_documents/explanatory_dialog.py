from __future__ import annotations

import traceback
from pathlib import Path

from PyQt6.QtCore import QEvent, QObject, QPointF, QThread, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QCursor, QTextCharFormat, QTextCursor, QTextFormat
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTextEdit,
    QVBoxLayout,
)

from geo_documents.explanatory_generator import (
    GenerationResult,
    generate_explanatory_note,
    save_generation_result_docx,
)
from geo_documents.llm_ollama import DEFAULT_MODEL, DEFAULT_OLLAMA_HOST, OllamaError, list_local_models
from geo_documents.template_model import BlockType, ExplanationTemplate, TemplateBlock, slugify_template_id
from geo_documents.template_store import (
    delete_template,
    ensure_default_templates,
    list_templates,
    save_template,
)


TYPE_LABELS: dict[BlockType, str] = {
    "plain": "Обычный",
    "fixed": "Жёлтый: неизменно",
    "generated": "Зелёный: найти/сгенерировать",
}
LABEL_TO_TYPE = {label: key for key, label in TYPE_LABELS.items()}
TYPE_PROPERTY = int(QTextFormat.Property.UserProperty) + 1
NOTE_PROPERTY = int(QTextFormat.Property.UserProperty) + 2
COMMENT_MARKER_PROPERTY = int(QTextFormat.Property.UserProperty) + 3
FIXED_COLOR = QColor("#fff4a3")
GENERATED_COLOR = QColor("#b8f5b1")
COMMENT_MARKER = "💬"


class _EditorCommentClickFilter(QObject):
    """Открывает заметку по клику на значок комментария в редакторе шаблона."""

    def __init__(self, dialog: "ExplanatoryNoteDialog") -> None:
        super().__init__()
        self._dialog = dialog

    def eventFilter(self, obj, event) -> bool:
        if obj is not self._dialog.editor.viewport():
            return super().eventFilter(obj, event)

        if event.type() == QEvent.Type.MouseMove:
            pos = self._dialog._doc_pos_from_viewport(event.position())
            text = self._dialog.editor.toPlainText()
            if pos is not None and 0 <= pos < len(text) and text[pos] == COMMENT_MARKER:
                self._dialog.editor.viewport().setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            else:
                self._dialog.editor.viewport().unsetCursor()
            return False

        if event.type() == QEvent.Type.MouseButtonRelease and event.button() == Qt.MouseButton.LeftButton:
            pos = self._dialog._doc_pos_from_viewport(event.position())
            if pos is not None and self._dialog._open_note_at_doc_pos(pos):
                return True
        return super().eventFilter(obj, event)


class _GenerateThread(QThread):
    done = pyqtSignal(str, str)
    crashed = pyqtSignal(str)

    def __init__(
        self,
        *,
        template: ExplanationTemplate,
        context_path: Path,
        output_path: Path,
        model: str,
        host: str,
        low_memory: bool,
        keep_highlight: bool,
    ) -> None:
        super().__init__()
        self._template = template
        self._context_path = context_path
        self._output_path = output_path
        self._model = model
        self._host = host
        self._low_memory = low_memory
        self._keep_highlight = keep_highlight

    def run(self) -> None:
        try:
            result: GenerationResult = generate_explanatory_note(
                self._template,
                [self._context_path],
                model=self._model,
                host=self._host,
                low_memory=self._low_memory,
            )
            save_generation_result_docx(result, self._output_path, keep_highlight=self._keep_highlight)
            self.done.emit(str(self._output_path), result.text)
        except Exception:
            self.crashed.emit(traceback.format_exc())


class ExplanatoryNoteDialog(QDialog):
    def __init__(self, *, initial_folder: Path, suggested_context: Path | None = None, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Пояснительная записка через локальную LLM")
        self.resize(1100, 760)
        self._initial_folder = initial_folder
        self._templates: list[ExplanationTemplate] = []
        self._thread: _GenerateThread | None = None

        ensure_default_templates()

        root = QVBoxLayout(self)

        top = QHBoxLayout()
        self.cb_template = QComboBox()
        self.cb_template.currentIndexChanged.connect(self._load_selected_template)
        btn_reload = QPushButton("Обновить")
        btn_reload.clicked.connect(self._reload_templates)
        btn_new = QPushButton("Новый")
        btn_new.clicked.connect(self._new_template)
        btn_duplicate = QPushButton("Дублировать")
        btn_duplicate.clicked.connect(self._duplicate_template)
        btn_delete = QPushButton("Удалить")
        btn_delete.clicked.connect(self._delete_template)
        top.addWidget(QLabel("Шаблон:"))
        top.addWidget(self.cb_template, stretch=1)
        top.addWidget(btn_reload)
        top.addWidget(btn_new)
        top.addWidget(btn_duplicate)
        top.addWidget(btn_delete)
        root.addLayout(top)

        meta = QGroupBox("Параметры шаблона")
        meta_form = QFormLayout(meta)
        self.ed_name = QLineEdit()
        self.ed_description = QLineEdit()
        meta_form.addRow("Название:", self.ed_name)
        meta_form.addRow("Описание:", self.ed_description)
        root.addWidget(meta)

        editor_group = QGroupBox("Текст шаблона")
        editor_layout = QVBoxLayout(editor_group)
        hint = QLabel(
            "Пишите цельный текст шаблона. Выделите фрагмент и нажмите: "
            "жёлтый — неизменяемый текст, зелёный — место для поиска/генерации по документу. "
            "К фрагменту можно добавить заметку кнопкой «💬 Заметка» или нажатием на значок 💬."
        )
        hint.setWordWrap(True)
        editor_layout.addWidget(hint)
        self.editor = QTextEdit()
        self.editor.setAcceptRichText(False)
        self.editor.setMouseTracking(True)
        self._editor_click_filter = _EditorCommentClickFilter(self)
        self.editor.viewport().installEventFilter(self._editor_click_filter)
        editor_layout.addWidget(self.editor, stretch=1)
        root.addWidget(editor_group, stretch=2)

        block_buttons = QHBoxLayout()
        btn_fixed = QPushButton("Жёлтый: неизменно")
        btn_fixed.clicked.connect(lambda: self._mark_selection("fixed"))
        btn_generated = QPushButton("Зелёный: найти/сгенерировать")
        btn_generated.clicked.connect(lambda: self._mark_selection("generated"))
        btn_plain = QPushButton("Очистить выделение")
        btn_plain.clicked.connect(self._clear_selection_format)
        btn_note = QPushButton(f"{COMMENT_MARKER} Заметка")
        btn_note.clicked.connect(self._set_selection_note)
        btn_save = QPushButton("Сохранить шаблон")
        btn_save.clicked.connect(self._save_current_template)
        block_buttons.addWidget(btn_fixed)
        block_buttons.addWidget(btn_generated)
        block_buttons.addWidget(btn_plain)
        block_buttons.addWidget(btn_note)
        block_buttons.addStretch(1)
        block_buttons.addWidget(btn_save)
        root.addLayout(block_buttons)

        gen = QGroupBox("Генерация")
        gen_form = QFormLayout(gen)
        row_context = QHBoxLayout()
        self.ed_context = QLineEdit(str(suggested_context or ""))
        btn_context = QPushButton("Обзор…")
        btn_context.clicked.connect(self._pick_context)
        row_context.addWidget(self.ed_context, stretch=1)
        row_context.addWidget(btn_context)
        gen_form.addRow("Документ-контекст:", row_context)

        row_output = QHBoxLayout()
        default_out = initial_folder / "explanatory_note_generated.docx"
        self.ed_output = QLineEdit(str(default_out))
        btn_output = QPushButton("Куда сохранить…")
        btn_output.clicked.connect(self._pick_output)
        row_output.addWidget(self.ed_output, stretch=1)
        row_output.addWidget(btn_output)
        gen_form.addRow("Результат DOCX:", row_output)

        self.cb_model = QComboBox()
        self.btn_refresh_models = QPushButton("Обновить модели")
        self.btn_refresh_models.clicked.connect(self._refresh_models)
        self.ed_host = QLineEdit(DEFAULT_OLLAMA_HOST)
        self.cb_low_memory = QCheckBox("Экономия памяти")
        self.cb_low_memory.setChecked(True)
        self.cb_keep_highlight = QCheckBox("Сохранять подсветку в результате")
        row_model = QHBoxLayout()
        row_model.addWidget(self.cb_model, stretch=1)
        row_model.addWidget(self.btn_refresh_models)
        gen_form.addRow("Модель Ollama:", row_model)
        gen_form.addRow("Ollama host:", self.ed_host)
        gen_form.addRow(self.cb_low_memory)
        gen_form.addRow(self.cb_keep_highlight)
        root.addWidget(gen)

        action_row = QHBoxLayout()
        self.btn_generate = QPushButton("Сгенерировать пояснительную записку")
        self.btn_generate.clicked.connect(self._generate)
        action_row.addStretch(1)
        action_row.addWidget(self.btn_generate)
        root.addLayout(action_row)

        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setPlaceholderText("Здесь появится текст результата после генерации.")
        root.addWidget(self.preview, stretch=1)

        self._reload_templates()
        self._refresh_models()

    def _reload_templates(self) -> None:
        ensure_default_templates()
        current_id = self._current_template().id if self._current_template() else ""
        self._templates = list_templates()
        self.cb_template.blockSignals(True)
        self.cb_template.clear()
        for template in self._templates:
            self.cb_template.addItem(template.name, template.id)
        self.cb_template.blockSignals(False)
        idx = next((i for i, t in enumerate(self._templates) if t.id == current_id), 0)
        if self._templates:
            self.cb_template.setCurrentIndex(idx)
            self._set_template_to_ui(self._templates[idx])

    def _current_template(self) -> ExplanationTemplate | None:
        idx = self.cb_template.currentIndex()
        if idx < 0 or idx >= len(self._templates):
            return None
        return self._templates[idx]

    def _load_selected_template(self) -> None:
        template = self._current_template()
        if template:
            self._set_template_to_ui(template)

    def _set_template_to_ui(self, template: ExplanationTemplate) -> None:
        self.ed_name.setText(template.name)
        self.ed_description.setText(template.description)
        self.editor.clear()
        cursor = self.editor.textCursor()
        for idx, block in enumerate(template.blocks):
            self._insert_block(cursor, block)
            next_block = template.blocks[idx + 1] if idx + 1 < len(template.blocks) else None
            if next_block and not block.text.endswith("\n") and not next_block.text.startswith("\n"):
                cursor.insertText("\n\n", QTextCharFormat())

    def _template_from_ui(self) -> ExplanationTemplate:
        name = self.ed_name.text().strip() or "Новый шаблон"
        current = self._current_template()
        template_id = current.id if current else slugify_template_id(name)
        blocks = self._blocks_from_editor()
        return ExplanationTemplate(
            id=template_id,
            name=name,
            description=self.ed_description.text().strip(),
            blocks=blocks,
        )

    def _new_template(self) -> None:
        self.cb_template.setCurrentIndex(-1)
        self.ed_name.setText("Новый шаблон")
        self.ed_description.clear()
        self.editor.setPlainText("1. Пояснительная записка\n\n[текст раздела]")

    def _duplicate_template(self) -> None:
        template = self._template_from_ui()
        template.id = slugify_template_id(template.name + "_copy")
        template.name = template.name + " (копия)"
        save_template(template)
        self._reload_templates()

    def _delete_template(self) -> None:
        template = self._current_template()
        if not template:
            return
        if QMessageBox.question(self, "Удалить шаблон", f"Удалить шаблон «{template.name}»?") != QMessageBox.StandardButton.Yes:
            return
        delete_template(template.id)
        self._reload_templates()

    def _save_current_template(self) -> None:
        template = self._template_from_ui()
        if not template.blocks:
            QMessageBox.warning(self, "Шаблон", "Добавьте хотя бы один блок.")
            return
        save_template(template)
        QMessageBox.information(self, "Шаблон", "Шаблон сохранён.")
        self._reload_templates()

    def _char_format_for_block(self, block: TemplateBlock) -> QTextCharFormat:
        fmt = QTextCharFormat()
        fmt.setProperty(TYPE_PROPERTY, block.type)
        if block.note:
            fmt.setProperty(NOTE_PROPERTY, block.note)
            fmt.setToolTip(block.note)
        if block.type == "fixed":
            fmt.setBackground(FIXED_COLOR)
        elif block.type == "generated":
            fmt.setBackground(GENERATED_COLOR)
        return fmt

    def _insert_block(self, cursor: QTextCursor, block: TemplateBlock) -> None:
        cursor.insertText(block.text, self._char_format_for_block(block))
        if block.note:
            cursor.insertText(COMMENT_MARKER, self._comment_marker_format(block.note))

    def _comment_marker_format(self, note: str) -> QTextCharFormat:
        fmt = QTextCharFormat()
        fmt.setForeground(QColor("#1f5fbf"))
        fmt.setProperty(COMMENT_MARKER_PROPERTY, True)
        fmt.setProperty(NOTE_PROPERTY, note)
        fmt.setToolTip(note)
        return fmt

    def _block_type_from_format(self, fmt: QTextCharFormat) -> BlockType:
        prop = fmt.property(TYPE_PROPERTY)
        if prop in {"fixed", "generated", "plain"}:
            return prop
        color = fmt.background().color()
        if color == FIXED_COLOR:
            return "fixed"
        if color == GENERATED_COLOR:
            return "generated"
        return "plain"

    def _blocks_from_editor(self) -> list[TemplateBlock]:
        blocks: list[TemplateBlock] = []
        text = self.editor.toPlainText()
        if not text:
            return []

        current_type: BlockType | None = None
        current_note = ""
        current_text: list[str] = []

        for pos, char in enumerate(text):
            cursor = QTextCursor(self.editor.document())
            cursor.setPosition(pos)
            cursor.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 1)
            fmt = cursor.charFormat()
            if char == COMMENT_MARKER:
                continue
            block_type = self._block_type_from_format(fmt)
            # Неразмеченный текст нужен пользователю как каркас шаблона, но не как блок результата.
            note = str(fmt.property(NOTE_PROPERTY) or "") if block_type != "plain" else ""
            if block_type != current_type or note != current_note:
                if current_text and current_type is not None:
                    blocks.append(TemplateBlock.create(type=current_type, text="".join(current_text), note=current_note))
                current_type = block_type
                current_note = note
                current_text = [char]
            else:
                current_text.append(char)

        if current_text and current_type is not None:
            blocks.append(TemplateBlock.create(type=current_type, text="".join(current_text), note=current_note))
        return self._merge_adjacent_blocks(blocks)

    def _merge_adjacent_blocks(self, blocks: list[TemplateBlock]) -> list[TemplateBlock]:
        merged: list[TemplateBlock] = []
        for block in blocks:
            if not block.text:
                continue
            if block.type == "plain":
                continue
            if merged and merged[-1].type == block.type and merged[-1].note == block.note:
                merged[-1].text += block.text
            else:
                merged.append(block)
        return merged

    def _selection_cursor(self) -> QTextCursor | None:
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            QMessageBox.warning(self, "Шаблон", "Сначала выделите часть текста.")
            return None
        return cursor

    def _mark_selection(self, block_type: BlockType) -> None:
        cursor = self._selection_cursor()
        if cursor is None:
            return
        fmt = QTextCharFormat()
        fmt.setProperty(TYPE_PROPERTY, block_type)
        fmt.setBackground(FIXED_COLOR if block_type == "fixed" else GENERATED_COLOR)
        cursor.mergeCharFormat(fmt)

    def _clear_selection_format(self) -> None:
        cursor = self._selection_cursor()
        if cursor is None:
            return
        cursor.setCharFormat(QTextCharFormat())
        self._remove_comment_marker_after_selection(cursor)

    def _set_selection_note(self) -> None:
        cursor = self._selection_cursor()
        if cursor is None:
            return
        current_note = str(cursor.charFormat().property(NOTE_PROPERTY) or "")
        self._edit_note_for_cursor(cursor, current_note=current_note)

    def _doc_pos_from_viewport(self, pos: QPointF) -> int | None:
        cursor = self.editor.cursorForPosition(pos.toPoint())
        return cursor.position()

    def _open_note_at_doc_pos(self, doc_pos: int) -> bool:
        text = self.editor.toPlainText()
        if not text or doc_pos < 0 or doc_pos >= len(text):
            return False

        if text[doc_pos] == COMMENT_MARKER:
            note = str(self._char_format_at(doc_pos).property(NOTE_PROPERTY) or "")
            colored_end = doc_pos
            left = colored_end
            while left > 0 and text[left - 1] != COMMENT_MARKER:
                if self._block_type_from_format(self._char_format_at(left - 1)) != self._block_type_from_format(
                    self._char_format_at(colored_end - 1)
                ):
                    break
                left -= 1
            target = QTextCursor(self.editor.document())
            target.setPosition(left)
            target.setPosition(colored_end, QTextCursor.MoveMode.KeepAnchor)
            self._edit_note_for_cursor(target, current_note=note)
            return True

        target = self._note_target_at_doc_pos(doc_pos)
        if target is None:
            return False
        note = str(target.charFormat().property(NOTE_PROPERTY) or "")
        if not note:
            return False
        self._edit_note_for_cursor(target, current_note=note)
        return True

    def _note_target_at_doc_pos(self, doc_pos: int) -> QTextCursor | None:
        text = self.editor.toPlainText()
        if not text:
            return None
        pos = min(max(doc_pos, 0), len(text) - 1)
        fmt = self._char_format_at(pos)
        if self._block_type_from_format(fmt) == "plain":
            return None

        left = pos
        while left > 0 and text[left - 1] != COMMENT_MARKER and self._block_type_from_format(self._char_format_at(left - 1)) == self._block_type_from_format(fmt):
            left -= 1

        right = pos
        while right + 1 < len(text) and text[right + 1] != COMMENT_MARKER and self._block_type_from_format(self._char_format_at(right + 1)) == self._block_type_from_format(fmt):
            right += 1

        target = QTextCursor(self.editor.document())
        target.setPosition(left)
        target.setPosition(right + 1, QTextCursor.MoveMode.KeepAnchor)
        return target

    def _char_format_at(self, pos: int) -> QTextCharFormat:
        cursor = QTextCursor(self.editor.document())
        cursor.setPosition(pos)
        cursor.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 1)
        return cursor.charFormat()

    def _note_target_cursor(self) -> QTextCursor | None:
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            return cursor

        text = self.editor.toPlainText()
        if not text:
            return None
        pos = min(cursor.position(), len(text) - 1)
        if text[pos] == COMMENT_MARKER:
            pos -= 1
        elif pos > 0 and text[pos - 1] == COMMENT_MARKER:
            pos -= 2
        if pos < 0:
            return None

        fmt = self._char_format_at(pos)
        block_type = self._block_type_from_format(fmt)
        if block_type == "plain":
            return None

        left = pos
        while left > 0 and text[left - 1] != COMMENT_MARKER and self._block_type_from_format(self._char_format_at(left - 1)) == block_type:
            left -= 1

        right = pos
        while right + 1 < len(text) and text[right + 1] != COMMENT_MARKER and self._block_type_from_format(self._char_format_at(right + 1)) == block_type:
            right += 1

        target = QTextCursor(self.editor.document())
        target.setPosition(left)
        target.setPosition(right + 1, QTextCursor.MoveMode.KeepAnchor)
        return target

    def _edit_note_for_cursor(self, cursor: QTextCursor, *, current_note: str) -> None:
        note, ok = QInputDialog.getMultiLineText(
            self,
            "Заметка к выделению",
            "Инструкция для LLM (откуда брать информацию, как формулировать):",
            current_note,
        )
        if not ok:
            return
        fmt = QTextCharFormat()
        clean_note = note.strip()
        fmt.setProperty(NOTE_PROPERTY, clean_note)
        fmt.setToolTip(clean_note)
        cursor.mergeCharFormat(fmt)
        self._remove_comment_marker_after_selection(cursor)
        if clean_note:
            marker_cursor = QTextCursor(cursor)
            marker_cursor.setPosition(max(cursor.selectionStart(), cursor.selectionEnd()))
            marker_cursor.insertText(COMMENT_MARKER, self._comment_marker_format(clean_note))

    def _remove_comment_marker_after_selection(self, cursor: QTextCursor) -> None:
        marker_cursor = QTextCursor(cursor)
        marker_cursor.setPosition(max(cursor.selectionStart(), cursor.selectionEnd()))
        marker_cursor.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 1)
        if marker_cursor.selectedText() == COMMENT_MARKER:
            marker_cursor.removeSelectedText()

    def _pick_context(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите сформированный документ",
            str(self._initial_folder),
            "Документы (*.docx *.pdf *.txt *.md);;Все файлы (*.*)",
        )
        if path:
            self.ed_context.setText(path)

    def _pick_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить пояснительную записку",
            self.ed_output.text() or str(self._initial_folder / "explanatory_note_generated.docx"),
            "DOCX (*.docx)",
        )
        if path:
            if not path.lower().endswith(".docx"):
                path += ".docx"
            self.ed_output.setText(path)

    def _generate(self) -> None:
        if self._thread is not None and self._thread.isRunning():
            return
        template = self._template_from_ui()
        context = Path(self.ed_context.text().strip())
        output = Path(self.ed_output.text().strip())
        if not context.is_file():
            QMessageBox.warning(self, "Пояснительная записка", "Выберите существующий документ-контекст.")
            return
        if not output.name:
            QMessageBox.warning(self, "Пояснительная записка", "Укажите путь для результата DOCX.")
            return
        self._save_current_template()
        self.btn_generate.setEnabled(False)
        self.preview.setPlainText("Генерация запущена. Это может занять несколько минут...")
        self._thread = _GenerateThread(
            template=template,
            context_path=context,
            output_path=output,
            model=self.cb_model.currentText().strip() or DEFAULT_MODEL,
            host=self.ed_host.text().strip() or DEFAULT_OLLAMA_HOST,
            low_memory=self.cb_low_memory.isChecked(),
            keep_highlight=self.cb_keep_highlight.isChecked(),
        )
        self._thread.done.connect(self._on_generated)
        self._thread.crashed.connect(self._on_crashed)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _on_generated(self, path: str, text: str) -> None:
        self.btn_generate.setEnabled(True)
        self.preview.setPlainText(text)
        QMessageBox.information(self, "Готово", f"Пояснительная записка сохранена:\n{path}")

    def _on_crashed(self, tb: str) -> None:
        self.btn_generate.setEnabled(True)
        QMessageBox.critical(self, "Ошибка генерации", tb)

    def _refresh_models(self) -> None:
        current = self.cb_model.currentText().strip() or DEFAULT_MODEL
        self.cb_model.clear()
        try:
            models = list_local_models(self.ed_host.text().strip() or DEFAULT_OLLAMA_HOST, timeout=5)
        except OllamaError:
            models = []
        if not models:
            models = [DEFAULT_MODEL]
        for model in models:
            self.cb_model.addItem(model)
        idx = self.cb_model.findText(current)
        if idx >= 0:
            self.cb_model.setCurrentIndex(idx)
