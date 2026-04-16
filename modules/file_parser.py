from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re


PREFIX_PATTERN = re.compile(r"^(\d+)(.*)$")
IGNORED_WAV_PREFIXES = (
    ".",
    "._",
    ".renaming_tmp_",
    ".undo_tmp_",
)


@dataclass(slots=True)
class ParsedAudioFile:
    path: Path
    original_filename: str
    original_index: int
    sort_suffix: str
    text_portion: str
    duplicate_index: bool = False


@dataclass(slots=True)
class ParseResult:
    files: list[ParsedAudioFile]
    excluded_files: list[str]
    detected_indices: list[int]
    duplicate_indices: list[int]


def extract_numeric_prefix(filename: str) -> tuple[int, str] | None:
    match = PREFIX_PATTERN.match(filename)
    if not match:
        return None
    number_text, remainder = match.groups()
    return int(number_text), remainder


def parse_audio_folder(folder: Path, selected_filenames: set[str] | None = None) -> ParseResult:
    parsed_files: list[ParsedAudioFile] = []
    excluded_files: list[str] = []

    for path in sorted(folder.iterdir(), key=lambda item: item.name.lower()):
        if not path.is_file():
            continue
        if path.suffix.lower() != ".wav":
            continue
        if path.name.startswith(IGNORED_WAV_PREFIXES):
            continue
        if selected_filenames is not None and path.name not in selected_filenames:
            continue

        parsed = extract_numeric_prefix(path.stem)
        if parsed is None:
            excluded_files.append(path.name)
            continue

        index, remainder = parsed
        parsed_files.append(
            ParsedAudioFile(
                path=path,
                original_filename=path.name,
                original_index=index,
                sort_suffix=path.name.lower(),
                text_portion=remainder,
            )
        )

    parsed_files.sort(key=lambda item: (item.original_index, item.sort_suffix))

    index_counts: dict[int, int] = {}
    for item in parsed_files:
        index_counts[item.original_index] = index_counts.get(item.original_index, 0) + 1

    duplicate_indices = sorted(index for index, count in index_counts.items() if count > 1)
    detected_indices = sorted(index_counts)

    for item in parsed_files:
        item.duplicate_index = item.original_index in duplicate_indices

    return ParseResult(
        files=parsed_files,
        excluded_files=excluded_files,
        detected_indices=detected_indices,
        duplicate_indices=duplicate_indices,
    )
