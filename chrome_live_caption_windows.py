from __future__ import annotations

from captioned_text import normalize_transcript_paragraph
from chrome_live_caption_common import split_live_caption_paragraphs


def find_live_caption_document(root_control, find_first_control):
    return find_first_control(
        root_control,
        lambda control: getattr(control, "ControlTypeName", "") == "DocumentControl"
        and getattr(control, "ClassName", "") == "CaptionBubbleLabel",
    )


def extract_live_caption_paragraphs(root_control, find_first_control) -> list[str]:
    document = find_live_caption_document(root_control, find_first_control)
    if document is None:
        return []

    try:
        name_text = (document.Name or "").strip()
    except Exception:
        name_text = ""

    if name_text:
        return split_live_caption_paragraphs(name_text)

    lines: list[str] = []
    try:
        for child in document.GetChildren():
            if child.ControlTypeName != "TextControl":
                continue
            if child.ClassName != "AXVirtualView":
                continue
            line = (child.Name or "").strip()
            if line:
                lines.append(line)
    except Exception:
        return []

    if not lines:
        return []
    paragraph = normalize_transcript_paragraph(" ".join(lines))
    return [paragraph] if paragraph else []
