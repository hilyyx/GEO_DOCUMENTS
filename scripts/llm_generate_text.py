from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from geo_documents.llm_ollama import (
    DEFAULT_MODEL,
    DEFAULT_OLLAMA_HOST,
    OllamaError,
    build_context_prompt,
    generate_text,
    list_local_models,
    read_document_context,
)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Генерация текста через локальную Ollama/Gemma с учетом контекста документов."
    )
    parser.add_argument("task", help="Что нужно сгенерировать.")
    parser.add_argument("files", nargs="*", type=Path, help="DOCX/PDF/TXT/MD файлы для контекста.")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Модель Ollama, по умолчанию {DEFAULT_MODEL}.")
    parser.add_argument("--host", default=DEFAULT_OLLAMA_HOST, help=f"Адрес Ollama, по умолчанию {DEFAULT_OLLAMA_HOST}.")
    parser.add_argument("--output", type=Path, help="Куда сохранить ответ в UTF-8.")
    parser.add_argument("--temperature", type=float, default=0.2, help="Температура генерации.")
    parser.add_argument("--timeout", type=int, default=180, help="Таймаут запроса к Ollama в секундах.")
    parser.add_argument("--max-chars-per-file", type=int, default=12_000, help="Лимит контекста на файл.")
    parser.add_argument("--max-total-chars", type=int, default=30_000, help="Общий лимит контекста.")
    parser.add_argument("--list-models", action="store_true", help="Показать локальные модели Ollama и выйти.")
    parser.add_argument("--dry-run", action="store_true", help="Не вызывать Ollama, только показать собранный промпт.")
    return parser.parse_args()


def main() -> int:
    args = _parse_args()

    try:
        if args.list_models:
            for model in list_local_models(args.host, timeout=args.timeout):
                print(model)
            return 0

        if args.dry_run:
            contexts = read_document_context(args.files, max_chars_per_file=args.max_chars_per_file)
            print(build_context_prompt(args.task, contexts, max_total_chars=args.max_total_chars))
            return 0

        text = generate_text(
            task=args.task,
            context_paths=args.files,
            model=args.model,
            host=args.host,
            temperature=args.temperature,
            timeout=args.timeout,
            max_chars_per_file=args.max_chars_per_file,
            max_total_chars=args.max_total_chars,
        )
    except OllamaError as e:
        print(f"Ошибка Ollama: {e}")
        return 2
    except Exception as e:
        print(f"Ошибка: {e}")
        return 1

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
        print(f"Сохранено: {args.output}")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
