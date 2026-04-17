from __future__ import annotations

import csv
from pathlib import Path


BASE_FIELDNAMES = [
    "original_filename",
    "new_filename",
    "original_index",
    "new_index",
    "status",
]
OPTIONAL_FIELD_ORDER = [
    "folder_path",
    "processed_at",
    "note",
]


def write_rename_log(rows: list[dict[str, str]], destination: Path) -> None:
    fieldnames = list(BASE_FIELDNAMES)
    seen = set(fieldnames)

    for name in OPTIONAL_FIELD_ORDER:
        if any(name in row for row in rows):
            fieldnames.append(name)
            seen.add(name)

    for row in rows:
        for name in row:
            if name not in seen:
                fieldnames.append(name)
                seen.add(name)

    with destination.open("w", newline="", encoding="utf-8") as csv_file:
        csv_file.write("\ufeff")
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
