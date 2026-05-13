# GEO Documents

Десктопное приложение для **склейки отчётных материалов** (PDF, DOCX) в один документ DOCX с опциональным экспортом в PDF. Ориентировано на типовые наборы файлов геоотчётов и ИГИ: нумерованные разделы, приложения по буквам, графика `3.N`.

## Технологии

| Компонент | Назначение |
|-----------|------------|
| **Python 3.12** | язык и рантайм |
| **PyQt6** | графический интерфейс (окно, списки, диалоги, настройки) |
| **python-docx** | создание и изменение DOCX (страницы, заголовки, вставка растра) |
| **docxcompose** | объединение нескольких DOCX в один документ (`Composer`) |
| **PyMuPDF (fitz)** | растеризация страниц PDF в изображения для вставки в DOCX |
| **lxml** | XML/HTML-движок (транзитивная зависимость `python-docx` / docxcompose) |
| **Babel** | локализация/форматирование (зависимость **docxcompose**) |
| **LibreOffice** (внешняя установка) | экспорт DOCX → PDF через `soffice.exe` |
| **PyInstaller** | сборка одиночного исполняемого файла `GEO_Documents.exe` под Windows |

## Функциональность

- Выбор папки и автоматическое сканирование файлов с расширениями `.pdf`, `.docx` (временные файлы Office `~$*` игнорируются).
- Список документов с **перетаскиванием** для ручного порядка, кнопки «вверх / вниз», добавление отдельных файлов, удаление из списка.
- **Автосортировка** имён по правилам, близким к отчётной нумерации: сначала числовые разделы (`1.2`, `1.4.1`), затем приложения буквой (`А.`, `Г.3`), затем графика `3.N` и псевдонимы `э./ю./я.` — см. модуль `geo_documents.file_sorter`.
- Параметры склейки:
  - разрыв страницы между фрагментами;
  - опциональные заголовки с именем исходного файла;
  - DPI растра при включении PDF в DOCX (по умолчанию 120 dpi).
- Итоговые файлы: **`имя.docx`** и **`имя.pdf`** в выбранной папке (базовое имя задаётся в поле ввода).
- Путь к **soffice.exe** задаётся в интерфейсе и сохраняется через `QSettings`; дополнительно поддерживаются переменные окружения `LIBREOFFICE_EXECUTABLE` и `SOFFICE_PATH`.

**Ограничения:** файлы `.doc` не читаются и пропускаются. Без LibreOffice приложение создаст DOCX, но не сможет экспортировать финальный PDF.

## Архитектура

Приложение разбито на тонкий слой входа и модули предметной области.

```text
launcher.py              # точка входа для PyInstaller → вызывает main()
geo_documents/
  __main__.py            # python -m geo_documents
  main.py                # QApplication, создание MainWindow
  window.py              # MainWindow: UI, диалоги, вызов склейки
  merger.py              # merge_to_docx_and_pdf: оркестрация конвертаций и сборки
  file_sorter.py         # sort_key_from_filename, sorted_paths
  libreoffice.py         # find_soffice, docx_to_pdf
```

**Поток данных при склейке:**

1. `MainWindow` собирает упорядоченный список путей и параметры.
2. `merger.merge_to_docx_and_pdf` пропускает `.doc`; PDF-фрагменты превращаются в серию PNG через PyMuPDF и вставляются в общий документ; DOCX-фрагменты сшиваются через `docxcompose.Composer`.
3. Результат сохраняется в DOCX; при наличии `soffice` выполняется экспорт в PDF и при необходимости переименование/перемещение файла.

**Сборка exe:** файл `GEO_Documents.spec` задаёт one-file режим (`console=False`), подключает биндинги **PyMuPDF** через `collect_all`, скрытые импорты для **docxcompose** / **babel** / **lxml**, входной скрипт — `launcher.py`. Итог: `dist/GEO_Documents.exe`.

### Сборка из исходников (Windows)

```powershell
cd путь\к\GEO_DOCUMENTS
.\.venv\Scripts\pip install -r requirements.txt -r requirements-build.txt
.\.venv\Scripts\python.exe -m PyInstaller --clean --noconfirm GEO_Documents.spec
```

Готовый файл: **`dist\GEO_Documents.exe`**.

### Запуск без сборки

```powershell
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\python.exe -m geo_documents
```

При необходимости установите [LibreOffice](https://www.libreoffice.org/) и укажите `soffice.exe` в окне приложения.
