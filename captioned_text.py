from __future__ import annotations

import re


PARAGRAPH_ID_PREFIX = "paragraph-"
TIMESTAMP_PREFIX_PATTERN = re.compile(
    r"^\s*(?:(?:\d+\s+(?:hours?|minutes?|seconds?))+\s*)+",
    re.IGNORECASE,
)
SPEAKER_BLOCK_PATTERN = re.compile(
    r"(Speaker\s+\d+)\s+((?:\d+\s+(?:hours?|minutes?|seconds?))(?:\s+\d+\s+(?:hours?|minutes?|seconds?))*)\s+",
    re.IGNORECASE,
)
SPEAKER_LINE_PATTERN = re.compile(r"^speaker\s+\d+\s*$", re.IGNORECASE)
TIMESTAMP_LINE_PATTERN = re.compile(
    r"^(?:\d+\s+(?:hours?|minutes?|seconds?))(?:\s+\d+\s+(?:hours?|minutes?|seconds?))*\s*$",
    re.IGNORECASE,
)


def normalize_transcript_paragraph(raw_paragraph: str) -> str:
    text = TIMESTAMP_PREFIX_PATTERN.sub("", raw_paragraph).strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_transcript_value(raw_text: str) -> list[str]:
    if not raw_text:
        return []

    lines = [line.strip() for line in raw_text.splitlines()]
    paragraphs: list[str] = []
    current_parts: list[str] = []
    skip_timestamp = False

    def flush_current() -> None:
        if not current_parts:
            return
        paragraph = normalize_transcript_paragraph(" ".join(current_parts))
        if paragraph:
            paragraphs.append(paragraph)
        current_parts.clear()

    for line in lines:
        if not line:
            continue
        if SPEAKER_LINE_PATTERN.match(line):
            flush_current()
            skip_timestamp = True
            continue
        if skip_timestamp and TIMESTAMP_LINE_PATTERN.match(line):
            skip_timestamp = False
            continue
        skip_timestamp = False
        current_parts.append(line)

    flush_current()
    return paragraphs


def extract_transcript_paragraphs_from_document_text(
    raw_text: str,
    paragraph_metadata: list[tuple[str, str]] | None = None,
) -> list[str]:
    if not raw_text:
        return []

    matches = list(SPEAKER_BLOCK_PATTERN.finditer(raw_text))
    if not matches:
        return []

    paragraphs: list[str] = []
    for index, match in enumerate(matches):
        content_start = match.end()
        content_end = matches[index + 1].start() if index + 1 < len(matches) else len(raw_text)
        paragraph = normalize_transcript_paragraph(raw_text[content_start:content_end])
        if paragraph:
            speaker = ""
            structural_text = ""
            if paragraph_metadata and index < len(paragraph_metadata):
                speaker, structural_text = paragraph_metadata[index]
                speaker = speaker.strip()
            if not speaker:
                speaker = match.group(1).strip()
            if structural_text:
                paragraph = choose_structural_paragraph_text(paragraph, structural_text)
            paragraphs.append(f"{speaker}\n{paragraph}")
    return paragraphs


def choose_structural_paragraph_text(document_text: str, structural_text: str) -> str:
    document_norm = normalized_compare_text(document_text)
    structural_norm = normalized_compare_text(structural_text)
    if not structural_norm:
        return document_text
    if document_norm.startswith(structural_norm):
        return structural_text
    if structural_norm in document_norm:
        return structural_text
    return document_text


def normalized_compare_text(text: str) -> str:
    return " ".join(text.casefold().split())
