from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field
from typing import Any, Literal


BlockType = Literal["plain", "fixed", "generated"]


def new_block_id() -> str:
    return f"b_{uuid.uuid4().hex[:10]}"


def slugify_template_id(name: str) -> str:
    slug = re.sub(r"[^0-9A-Za-zА-Яа-яЁё_-]+", "_", name.strip()).strip("_")
    return slug[:80] or f"template_{uuid.uuid4().hex[:8]}"


@dataclass
class TemplateBlock:
    id: str
    type: BlockType
    text: str
    note: str = ""

    @classmethod
    def create(cls, *, type: BlockType = "plain", text: str = "", note: str = "") -> "TemplateBlock":
        return cls(id=new_block_id(), type=type, text=text, note=note)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TemplateBlock":
        raw_type = data.get("type", "plain")
        block_type: BlockType = raw_type if raw_type in {"plain", "fixed", "generated"} else "plain"
        return cls(
            id=str(data.get("id") or new_block_id()),
            type=block_type,
            text=str(data.get("text") or ""),
            note=str(data.get("note") or ""),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "type": self.type,
            "text": self.text,
            "note": self.note,
        }


@dataclass
class ExplanationTemplate:
    id: str
    name: str
    description: str = ""
    version: int = 1
    blocks: list[TemplateBlock] = field(default_factory=list)

    @classmethod
    def create(cls, name: str) -> "ExplanationTemplate":
        return cls(id=slugify_template_id(name), name=name, blocks=[])

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ExplanationTemplate":
        name = str(data.get("name") or "Новый шаблон")
        return cls(
            id=str(data.get("id") or slugify_template_id(name)),
            name=name,
            description=str(data.get("description") or ""),
            version=int(data.get("version") or 1),
            blocks=[TemplateBlock.from_dict(item) for item in data.get("blocks", [])],
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "version": self.version,
            "blocks": [block.to_dict() for block in self.blocks],
        }


def default_explanatory_template() -> ExplanationTemplate:
    """Базовый шаблон по структуре приложенного примера пояснительной записки."""
    return ExplanationTemplate(
        id="default_igi_explanatory_note",
        name="Пояснительная записка ИГИ",
        description="Базовый шаблон пояснительной записки по инженерно-геологическим изысканиям.",
        blocks=[
            TemplateBlock.create(
                type="fixed",
                text="1. Пояснительная записка по инженерно-геологическим изысканиям",
                note="Заголовок всегда оставлять без изменений.",
            ),
            TemplateBlock.create(type="fixed", text="1.1 Введение"),
            TemplateBlock.create(
                type="generated",
                text="[вводный абзац об объекте, договоре, задании и нормативных основаниях]",
                note=(
                    "Найти в разделе 1.1 Введение: название объекта, номер договора, дату договора, "
                    "техническое задание, программу работ и нормативные основания. Сформулировать 1-2 абзаца."
                ),
            ),
            TemplateBlock.create(
                type="generated",
                text="[исполнитель, заказчик, проектировщик]",
                note=(
                    "Найти исполнителя, заказчика и проектировщика. Если данных нет, написать [уточнить]. "
                    "Сохранить деловой стиль."
                ),
            ),
            TemplateBlock.create(
                type="generated",
                text="[идентификационные сведения об объекте]",
                note="Найти идентификационные сведения, категорию объекта, вид строительства и стадию проектирования.",
            ),
            TemplateBlock.create(
                type="generated",
                text="[цель и задачи инженерных изысканий]",
                note="Искать рядом с фразами 'Целью изысканий' и 'Задачи инженерных изысканий'.",
            ),
            TemplateBlock.create(type="fixed", text="1.2 Изученность инженерно-геологических условий"),
            TemplateBlock.create(
                type="generated",
                text="[описание изученности инженерно-геологических условий]",
                note="Искать в разделе 1.2. Не выдумывать ранее выполненные работы, если их нет в контексте.",
            ),
            TemplateBlock.create(type="fixed", text="1.3 Физико-географические и техногенные условия"),
            TemplateBlock.create(
                type="generated",
                text="[административное положение, рельеф, геоморфология и техногенные условия]",
                note="Искать в разделе 1.3. Сформулировать связным техническим текстом.",
            ),
        ],
    )
