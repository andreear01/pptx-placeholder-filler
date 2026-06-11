from __future__ import annotations

import base64
import io
import json
import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, Iterable, Tuple, Any

from pptx import Presentation


PLACEHOLDER_RE = re.compile(r"\{\{AI_[A-Z0-9_]+\}\}")


@dataclass
class FillReport:
    replaced_text_placeholders: int
    removed_image_placeholders: int
    remaining_placeholders: list[str]
    provided_but_not_found: list[str]
    image_prompts: dict[str, str]


def _normalize_mapping(mapping: dict[str, Any]) -> dict[str, str]:
    """
    Accepts either:
    {
      "{{AI_X}}": "text"
    }

    or a looser format where values can be objects:
    {
      "{{AI_X}}": {"text": "..."} / {"text_to_insert": "..."} / {"image_prompt": "..."}
    }
    """
    normalized: dict[str, str] = {}

    for key, value in mapping.items():
        if not isinstance(key, str):
            continue

        key = key.strip()

        if not PLACEHOLDER_RE.fullmatch(key):
            continue

        if isinstance(value, str):
            normalized[key] = value
        elif isinstance(value, dict):
            candidate = (
                value.get("text")
                or value.get("text_to_insert")
                or value.get("Text to insert")
                or value.get("image_prompt")
                or value.get("Image prompt")
                or ""
            )
            normalized[key] = str(candidate)
        else:
            normalized[key] = str(value)

    return normalized


def _iter_shapes(shapes):
    """
    Recursively iterates through normal shapes, grouped shapes and table cells.
    """
    for shape in shapes:
        yield shape

        if hasattr(shape, "shapes"):
            yield from _iter_shapes(shape.shapes)

        if getattr(shape, "has_table", False):
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    for subshape in _iter_shapes(cell.text_frame.paragraphs):
                        yield subshape


def _iter_text_frames(prs: Presentation):
    """
    Yields all text frames from slides and table cells.
    """
    for slide in prs.slides:
        for shape in _iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                yield shape.text_frame

            if getattr(shape, "has_table", False):
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            yield cell.text_frame


def _replace_in_paragraph(paragraph, replacements: dict[str, str], remove_image_placeholders: bool) -> tuple[bool, int, int, dict[str, str]]:
    """
    Replaces placeholders inside a paragraph.
    Preserves formatting as much as possible by writing the new content into the first run
    and clearing the remaining runs.
    """
    if not paragraph.runs:
        return False, 0, 0, {}

    original = "".join(run.text for run in paragraph.runs)
    updated = original
    replaced_text = 0
    removed_images = 0
    image_prompts: dict[str, str] = {}

    for placeholder, value in replacements.items():
        if placeholder not in updated:
            continue

        if placeholder.startswith("{{AI_IMAGE_"):
            image_prompts[placeholder] = value
            if remove_image_placeholders:
                updated = updated.replace(placeholder, "")
                removed_images += original.count(placeholder)
            else:
                updated = updated.replace(placeholder, value)
                replaced_text += original.count(placeholder)
        else:
            updated = updated.replace(placeholder, value)
            replaced_text += original.count(placeholder)

    if updated != original:
        paragraph.runs[0].text = updated
        for run in paragraph.runs[1:]:
            run.text = ""
        return True, replaced_text, removed_images, image_prompts

    return False, 0, 0, {}


def _find_placeholders(prs: Presentation) -> set[str]:
    found: set[str] = set()
    for tf in _iter_text_frames(prs):
        for p in tf.paragraphs:
            text = "".join(run.text for run in p.runs)
            found.update(PLACEHOLDER_RE.findall(text))
    return found


def fill_pptx_bytes(
    pptx_bytes: bytes,
    mapping: dict[str, Any],
    remove_image_placeholders: bool = True,
) -> tuple[bytes, FillReport]:
    replacements = _normalize_mapping(mapping)

    prs = Presentation(io.BytesIO(pptx_bytes))

    placeholders_before = _find_placeholders(prs)

    replaced_text_count = 0
    removed_image_count = 0
    collected_image_prompts: dict[str, str] = {}

    for tf in _iter_text_frames(prs):
        for paragraph in tf.paragraphs:
            _, replaced_text, removed_images, image_prompts = _replace_in_paragraph(
                paragraph, replacements, remove_image_placeholders
            )
            replaced_text_count += replaced_text
            removed_image_count += removed_images
            collected_image_prompts.update(image_prompts)

    placeholders_after = _find_placeholders(prs)

    provided_keys = set(replacements.keys())
    provided_but_not_found = sorted(provided_keys - placeholders_before)

    output = io.BytesIO()
    prs.save(output)

    report = FillReport(
        replaced_text_placeholders=replaced_text_count,
        removed_image_placeholders=removed_image_count,
        remaining_placeholders=sorted(placeholders_after),
        provided_but_not_found=provided_but_not_found,
        image_prompts=collected_image_prompts,
    )
    return output.getvalue(), report


def fill_pptx_file(
    input_pptx_path: str | Path,
    mapping_json_path: str | Path,
    output_pptx_path: str | Path,
    image_prompts_path: str | Path | None = None,
    report_path: str | Path | None = None,
) -> FillReport:
    pptx_bytes = Path(input_pptx_path).read_bytes()
    mapping = json.loads(Path(mapping_json_path).read_text(encoding="utf-8"))

    filled_bytes, report = fill_pptx_bytes(pptx_bytes, mapping)

    Path(output_pptx_path).write_bytes(filled_bytes)

    if image_prompts_path:
        lines = []
        for placeholder, prompt in report.image_prompts.items():
            lines.append(f"{placeholder}\n{prompt}\n")
        Path(image_prompts_path).write_text("\n".join(lines), encoding="utf-8")

    if report_path:
        Path(report_path).write_text(json.dumps(asdict(report), ensure_ascii=False, indent=2), encoding="utf-8")

    return report


def pptx_to_base64(pptx_bytes: bytes) -> str:
    return base64.b64encode(pptx_bytes).decode("utf-8")


def base64_to_pptx(b64: str) -> bytes:
    return base64.b64decode(b64)
