from __future__ import annotations

import re
from pathlib import Path

# Порядок приложений в типовом ИГИ: сначала раздел 1.x, затем А…, затем графика 3.x
_CYR_ORDER = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"


def _letter_index(ch: str) -> int | None:
    ch = ch.upper()
    if ch in _CYR_ORDER:
        return _CYR_ORDER.index(ch)
    if len(ch) == 1 and "A" <= ch <= "Z":
        return 100 + ord(ch)
    return None


def sort_key_from_filename(filename: str) -> tuple:
    """
    Возвращает ключ сортировки: чем меньше tuple, тем раньше в отчёте.
    Уровни: 0 — числовые разделы (1.1, 1.2), 1 — приложения буквой (А., Г.3),
    2 — графика 3.N или псевдонимы э./ю./я., 9 — не распознано.
    """
    stem = Path(filename).stem.strip()

    # 3.N в начале (графические приложения)
    m = re.match(r"^3\.(\d+)\b", stem, re.IGNORECASE)
    if m:
        return (2, 0, (3, int(m.group(1))), (), stem.casefold())

    # Частые «ломанные» обозначения графики → после текстовых приложений
    if stem.startswith("э."):
        return (2, 1, (3, 1), (), stem.casefold())
    if stem.startswith("ю."):
        return (2, 1, (3, 2), (), stem.casefold())
    if stem.startswith("я."):
        return (2, 1, (3, 3), (), stem.casefold())

    # 1.2 / 1.2. / 1.4 / 1.6 …
    m = re.match(r"^(\d+(?:\.\d+)+)\.?\s*", stem)
    if m:
        parts = tuple(int(x) for x in m.group(1).split("."))
        return (0, 0, parts, (), stem.casefold())

    m = re.match(r"^(\d+)\.?\s+", stem)
    if m:
        parts = (int(m.group(1)),)
        return (0, 0, parts, (), stem.casefold())

    # Приложение: «А. », «А.», «Г.3», «В.Ведомость»
    m = re.match(r"^([А-ЯЁA-Za-z])\.?\s*(\d+)?", stem)
    if m:
        letter = m.group(1)
        idx = _letter_index(letter)
        if idx is not None:
            sub = int(m.group(2)) if m.group(2) else 0
            return (1, 0, (idx, sub), (), stem.casefold())

    return (9, 9, (9999,), (), stem.casefold())


def sorted_paths(paths: list[str | Path]) -> list[Path]:
    items = [(sort_key_from_filename(Path(p).name), Path(p)) for p in paths]
    items.sort(key=lambda x: x[0])
    return [p for _, p in items]
