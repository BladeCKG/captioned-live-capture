import difflib
import re
from dataclasses import dataclass
from functools import lru_cache

try:
    from rapidfuzz import fuzz
except ImportError:
    fuzz = None

try:
    from wordfreq import zipf_frequency
except ImportError:
    zipf_frequency = None

try:
    from lingua import Language, LanguageDetectorBuilder
except ImportError:
    Language = None
    LanguageDetectorBuilder = None


SPEAKER_LINE_PATTERN = re.compile(r"\bspe[a-z]*ker\s*\d*\b", re.IGNORECASE)
WORD_PATTERN = re.compile(r"[A-Za-z]+(?:['-][A-Za-z]+)?")
COMMON_TRANSCRIPT_WORDS = {
    "a", "about", "all", "am", "an", "and", "are", "as", "at", "be", "but", "can", "do",
    "for", "from", "go", "have", "he", "i", "if", "in", "is", "it", "me", "my", "no", "not",
    "of", "okay", "on", "or", "our", "right", "so", "that", "the", "then", "there", "this",
    "to", "uh", "um", "we", "well", "what", "with", "would", "yeah", "yes", "you", "your",
}
ALLOWED_UPPERCASE_WORDS = {"CEO", "SS"}
VOWELS = set("aeiou")


@dataclass(frozen=True)
class TranscriptProcessingConfig:
    min_line_letters: int = 4
    min_line_words: int = 2
    max_symbol_ratio: float = 0.45
    max_short_word_ratio: float = 0.60
    max_malformed_word_ratio: float = 0.25
    min_plausible_word_ratio: float = 0.55
    min_english_like_ratio: float = 0.35
    min_common_word_ratio: float = 0.08
    min_language_detection_chars: int = 25
    english_zipf_threshold: float = 2.4
    min_average_word_length: float = 2.2


@dataclass(frozen=True)
class MergerConfig:
    fragment_recent_paragraphs: int = 4
    duplicate_recent_paragraphs: int = 6
    min_containment_chars: int = 30
    min_extension_delta: int = 10
    fragment_partial_ratio: float = 0.90
    fragment_token_set_ratio: float = 0.78
    duplicate_token_set_ratio: float = 0.90
    duplicate_ratio: float = 0.82
    duplicate_token_overlap: float = 0.72


@dataclass(frozen=True)
class MatchScores:
    has_text: bool
    exact: bool
    ratio: float
    partial_ratio: float
    token_set_ratio: float
    token_overlap: float
    containment_ratio: float
    existing_len: int
    candidate_len: int


def build_english_detector():
    if LanguageDetectorBuilder is None or Language is None:
        return None
    try:
        return LanguageDetectorBuilder.from_languages(Language.ENGLISH).with_low_accuracy_mode().build()
    except Exception:
        return None


ENGLISH_DETECTOR = build_english_detector()


def normalized_line(line: str) -> str:
    return " ".join(WORD_PATTERN.findall(line.casefold()))


def split_paragraphs(text: str) -> list[str]:
    return [paragraph.strip() for paragraph in re.split(r"\n\s*\n+", text.strip()) if paragraph.strip()]


def is_ui_noise_line(line: str) -> bool:
    lowered = line.casefold()
    if "untitled recording" in lowered:
        return True
    if "feedback" in lowered and "share" in lowered:
        return True
    if "processing audio" in lowered:
        return True
    if re.search(r"\b\d{1,2}/\d{1,2}/\d{4}\b", lowered):
        return True
    if re.fullmatch(r"[\W_]*\d+x[\W_]*", lowered):
        return True
    if re.search(r"\b\d{1,2}:\d{2}\b", lowered) and not WORD_PATTERN.findall(lowered):
        return True
    return False


def looks_malformed_word(word: str) -> bool:
    if len(word) < 4:
        return False
    if word.upper() in ALLOWED_UPPERCASE_WORDS:
        return False
    if word.casefold() in COMMON_TRANSCRIPT_WORDS:
        return False
    uppercase_count = sum(ch.isupper() for ch in word)
    if word.isupper():
        return True
    return uppercase_count >= 2


def looks_plausible_word(word: str) -> bool:
    lowered = word.casefold()
    if len(lowered) < 3 or lowered in COMMON_TRANSCRIPT_WORDS:
        return True
    alpha_chars = [ch for ch in lowered if ch.isalpha()]
    if len(alpha_chars) < 3:
        return False
    vowel_ratio = sum(ch in VOWELS for ch in alpha_chars) / len(alpha_chars)
    if vowel_ratio < 0.18:
        return False
    repeated_runs = max((len(match.group(0)) for match in re.finditer(r"(.)\1+", lowered)), default=1)
    if repeated_runs >= 4:
        return False
    return not looks_malformed_word(word)


def looks_english_like_word(word: str, config: TranscriptProcessingConfig) -> bool:
    lowered = word.casefold()
    if len(lowered) < 3 or lowered in COMMON_TRANSCRIPT_WORDS:
        return True
    if zipf_frequency is not None:
        return zipf_frequency(lowered, "en") >= config.english_zipf_threshold
    return looks_plausible_word(word)


def line_is_probably_english(line: str, config: TranscriptProcessingConfig) -> bool:
    words = WORD_PATTERN.findall(line)
    alpha_only = " ".join(words)
    if len(alpha_only) < config.min_language_detection_chars or ENGLISH_DETECTOR is None:
        return True
    try:
        return ENGLISH_DETECTOR.detect_language_of(alpha_only) == Language.ENGLISH
    except Exception:
        return True


@lru_cache(maxsize=4096)
def normalized_words_for_text(text: str) -> tuple[str, ...]:
    return tuple(word.casefold().strip("'-") for word in WORD_PATTERN.findall(text))


def line_similarity(left: str, right: str) -> float:
    normalized_left = normalized_line(left)
    normalized_right = normalized_line(right)
    if not normalized_left or not normalized_right:
        return 1.0 if normalized_left == normalized_right else 0.0
    if fuzz is not None:
        return fuzz.ratio(normalized_left, normalized_right) / 100.0
    return difflib.SequenceMatcher(None, normalized_left, normalized_right).ratio()


class CapturedTextProcessor:
    def __init__(self, config: TranscriptProcessingConfig | None = None):
        self.config = config or TranscriptProcessingConfig()

    def clean(self, raw_text: str) -> str:
        kept_lines: list[str] = []
        pending_blank = False
        for raw_line in raw_text.splitlines():
            line = raw_line.rstrip()
            if SPEAKER_LINE_PATTERN.search(line):
                pending_blank = True
                continue
            if not line.strip():
                pending_blank = True
                continue
            if not self._keep_line(line):
                continue
            if pending_blank and kept_lines:
                kept_lines.append("")
            kept_lines.append(line)
            pending_blank = False
        return "\n".join(kept_lines).strip()

    def _keep_line(self, line: str) -> bool:
        stripped = line.strip()
        if not stripped or is_ui_noise_line(stripped):
            return False
        if not line_is_probably_english(stripped, self.config):
            return False

        words = WORD_PATTERN.findall(stripped)
        if len(words) < self.config.min_line_words:
            return False

        letters = sum(ch.isalpha() for ch in stripped)
        if letters < self.config.min_line_letters:
            return False

        alnum = sum(ch.isalnum() for ch in stripped)
        if alnum:
            symbol_ratio = 1 - (alnum / len(stripped))
            if symbol_ratio > self.config.max_symbol_ratio:
                return False

        normalized_words = normalized_words_for_text(stripped)
        common_word_count = sum(word in COMMON_TRANSCRIPT_WORDS for word in normalized_words)
        short_word_count = sum(len(word) <= 2 for word in normalized_words)
        malformed_word_count = sum(looks_malformed_word(word) for word in words)
        plausible_word_count = sum(looks_plausible_word(word) for word in words)
        english_like_count = sum(looks_english_like_word(word, self.config) for word in words)
        long_words = [word for word in words if len(word) >= 4]
        long_english_like_count = sum(looks_english_like_word(word, self.config) for word in long_words)
        average_word_length = sum(len(word) for word in words) / len(words)

        if plausible_word_count / len(words) < self.config.min_plausible_word_ratio:
            return False
        if english_like_count / len(words) < self.config.min_english_like_ratio:
            return False
        if (
            len(words) >= 10
            and common_word_count <= 3
            and malformed_word_count >= 2
            and long_words
            and (long_english_like_count / len(long_words)) < 0.65
        ):
            return False
        if len(words) <= 6 and short_word_count / len(words) > self.config.max_short_word_ratio:
            return False
        if len(words) >= 6 and malformed_word_count / len(words) > self.config.max_malformed_word_ratio:
            return False
        if len(words) >= 6 and common_word_count / len(words) < self.config.min_common_word_ratio:
            return False
        if len(words) <= 5 and common_word_count == 0 and not re.search(r"[.!?]", stripped):
            return False
        if average_word_length < self.config.min_average_word_length:
            return False
        return True


class TranscriptMerger:
    def __init__(self, config: MergerConfig | None = None):
        self.config = config or MergerConfig()

    def merge(self, existing: str, captured: str) -> tuple[str, bool]:
        captured = captured.strip()
        if not captured or captured == "(No text detected yet.)":
            return existing, False

        existing_paragraphs = split_paragraphs(existing)
        if not existing_paragraphs:
            return captured, True

        merged_paragraphs = list(existing_paragraphs)
        for candidate in split_paragraphs(captured):
            self._merge_candidate(merged_paragraphs, candidate)

        merged = "\n\n".join(merged_paragraphs).strip()
        return merged, merged != existing.strip()

    def _merge_candidate(self, paragraphs: list[str], candidate: str) -> None:
        if not candidate:
            return

        fragment_index = self._find_tail_fragment_index(paragraphs, candidate)
        if fragment_index is not None:
            paragraphs[fragment_index] = self._better_paragraph(paragraphs[fragment_index], candidate)
            return

        if self._find_recent_duplicate_index(paragraphs, candidate) is not None:
            return

        paragraphs.append(candidate)

    def _find_tail_fragment_index(self, paragraphs: list[str], candidate: str) -> int | None:
        if not paragraphs:
            return None

        last_index = len(paragraphs) - 1
        existing = paragraphs[last_index]
        if self._is_finished(existing):
            return None

        scores = self._compare(existing, candidate)
        if self._is_fragment_extension(scores):
            return last_index
        return None

    def _find_recent_duplicate_index(self, paragraphs: list[str], candidate: str) -> int | None:
        start = max(0, len(paragraphs) - self.config.duplicate_recent_paragraphs)
        for index in range(start, len(paragraphs)):
            existing = paragraphs[index]
            scores = self._compare(existing, candidate)
            if self._is_duplicate(scores):
                return index
        return None

    def _is_finished(self, paragraph: str) -> bool:
        return bool(re.search(r"[.!?][\"')\]]*$", paragraph.strip())) and len(WORD_PATTERN.findall(paragraph)) >= 8

    def _is_fragment_extension(self, scores: MatchScores) -> bool:
        if not scores.has_text:
            return False
        if scores.candidate_len <= scores.existing_len + self.config.min_extension_delta:
            return False
        if scores.containment_ratio >= 0.98:
            return True
        return (
            scores.partial_ratio >= self.config.fragment_partial_ratio
            and scores.token_set_ratio >= self.config.fragment_token_set_ratio
        )

    def _is_duplicate(self, scores: MatchScores) -> bool:
        if not scores.has_text:
            return False
        if scores.exact or scores.containment_ratio >= 0.98:
            return True
        if scores.token_set_ratio >= self.config.duplicate_token_set_ratio:
            return True
        return (
            scores.ratio >= self.config.duplicate_ratio
            and scores.token_overlap >= self.config.duplicate_token_overlap
        )

    def _better_paragraph(self, existing: str, candidate: str) -> str:
        existing_quality = self._quality(existing)
        candidate_quality = self._quality(candidate)
        if candidate_quality > existing_quality:
            return candidate
        if candidate_quality == existing_quality and len(candidate) > len(existing):
            return candidate
        return existing

    def _quality(self, paragraph: str) -> tuple[int, int, int]:
        words = WORD_PATTERN.findall(paragraph)
        malformed = sum(looks_malformed_word(word) for word in words)
        ending_punctuation = 1 if re.search(r"[.!?][\"')\]]*$", paragraph.strip()) else 0
        return (ending_punctuation, len(words) - malformed, len(paragraph))

    def _compare(self, existing: str, candidate: str) -> MatchScores:
        existing_norm = normalized_line(existing)
        candidate_norm = normalized_line(candidate)
        if not existing_norm or not candidate_norm:
            return MatchScores(
                has_text=False,
                exact=existing_norm == candidate_norm,
                ratio=0.0,
                partial_ratio=0.0,
                token_set_ratio=0.0,
                token_overlap=0.0,
                containment_ratio=0.0,
                existing_len=len(existing_norm),
                candidate_len=len(candidate_norm),
            )

        shorter, longer = sorted((existing_norm, candidate_norm), key=len)
        containment_ratio = (
            len(shorter) / len(longer)
            if len(shorter) >= self.config.min_containment_chars and shorter in longer
            else 0.0
        )
        token_overlap = self._token_overlap(existing_norm, candidate_norm)

        if fuzz is not None:
            ratio = fuzz.ratio(existing_norm, candidate_norm) / 100.0
            partial_ratio = fuzz.partial_ratio(existing_norm, candidate_norm) / 100.0
            token_set_ratio = fuzz.token_set_ratio(existing_norm, candidate_norm) / 100.0
        else:
            ratio = difflib.SequenceMatcher(None, existing_norm, candidate_norm).ratio()
            partial_ratio = ratio
            token_set_ratio = ratio

        return MatchScores(
            has_text=True,
            exact=existing_norm == candidate_norm,
            ratio=ratio,
            partial_ratio=partial_ratio,
            token_set_ratio=token_set_ratio,
            token_overlap=token_overlap,
            containment_ratio=containment_ratio,
            existing_len=len(existing_norm),
            candidate_len=len(candidate_norm),
        )

    def _token_overlap(self, existing_norm: str, candidate_norm: str) -> float:
        existing_words = {word for word in existing_norm.split() if len(word) > 2}
        candidate_words = {word for word in candidate_norm.split() if len(word) > 2}
        if not existing_words or not candidate_words:
            return 0.0
        return len(existing_words & candidate_words) / min(len(existing_words), len(candidate_words))


DEFAULT_TEXT_PROCESSOR = CapturedTextProcessor(TranscriptProcessingConfig())
