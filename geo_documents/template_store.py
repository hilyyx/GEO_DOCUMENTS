from __future__ import annotations

import json
import os
from pathlib import Path

from geo_documents.template_model import ExplanationTemplate, default_explanatory_template, slugify_template_id


def default_templates_dir() -> Path:
    base = Path(os.environ.get("APPDATA") or Path.home())
    return base / "GEO_DOCUMENTS" / "templates" / "explanatory_notes"


def ensure_templates_dir(path: Path | None = None) -> Path:
    target = path or default_templates_dir()
    target.mkdir(parents=True, exist_ok=True)
    return target


def template_path(template_id: str, *, directory: Path | None = None) -> Path:
    safe_id = slugify_template_id(template_id)
    return ensure_templates_dir(directory) / f"{safe_id}.json"


def save_template(template: ExplanationTemplate, *, directory: Path | None = None) -> Path:
    target = template_path(template.id, directory=directory)
    target.write_text(
        json.dumps(template.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return target


def load_template(path: Path) -> ExplanationTemplate:
    data = json.loads(path.read_text(encoding="utf-8"))
    return ExplanationTemplate.from_dict(data)


def list_templates(*, directory: Path | None = None) -> list[ExplanationTemplate]:
    target = ensure_templates_dir(directory)
    out: list[ExplanationTemplate] = []
    for path in sorted(target.glob("*.json")):
        try:
            out.append(load_template(path))
        except (OSError, json.JSONDecodeError, TypeError, ValueError):
            continue
    return out


def delete_template(template_id: str, *, directory: Path | None = None) -> None:
    path = template_path(template_id, directory=directory)
    path.unlink(missing_ok=True)


def ensure_default_templates(*, directory: Path | None = None) -> None:
    target = ensure_templates_dir(directory)
    if any(target.glob("*.json")):
        return
    save_template(default_explanatory_template(), directory=target)
