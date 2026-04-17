from __future__ import annotations

from collections.abc import Callable
from dataclasses import asdict, dataclass
from pathlib import Path
import json
import shutil

from modules.audio_editor import move_trim_backup_reference
from modules.file_parser import ParsedAudioFile


UNDO_MANIFEST_NAME = ".rename_undo.json"


@dataclass(slots=True)
class RenamePlanEntry:
    source_path: Path | None
    original_filename: str
    new_filename: str
    original_index: int
    new_index: int | None
    status: str


@dataclass(slots=True)
class RenameSettings:
    digits: int
    keep_text: bool
    move_ng_files: bool
    export_csv: bool
    ng_folder_name: str = "_NG"


def _build_filename(file_item: ParsedAudioFile, new_index: int, digits: int, keep_text: bool, text_value: str | None = None) -> str:
    number = f"{new_index:0{digits}d}"
    effective_text = file_item.text_portion if text_value is None else text_value
    if keep_text and effective_text:
        return f"{number}{effective_text}{file_item.path.suffix.lower()}"
    return f"{number}{file_item.path.suffix.lower()}"


def _next_available_index(candidate: int, missing_indices: set[int]) -> int:
    while candidate in missing_indices:
        candidate += 1
    return candidate


def build_rename_plan(
    files: list[ParsedAudioFile],
    ok_flags: dict[str, bool],
    missing_indices: set[int],
    settings: RenameSettings,
    text_overrides: dict[str, str] | None = None,
) -> list[RenamePlanEntry]:
    plan: list[RenamePlanEntry] = []
    next_index = 1

    for missing_index in sorted(missing_indices):
        plan.append(
            RenamePlanEntry(
                source_path=None,
                original_filename=f"{missing_index:0{settings.digits}d}",
                new_filename="",
                original_index=missing_index,
                new_index=None,
                status="MISSING",
            )
        )

    for file_item in files:
        is_ok = ok_flags.get(file_item.original_filename, True)
        if not is_ok:
            plan.append(
                RenamePlanEntry(
                    source_path=file_item.path,
                    original_filename=file_item.original_filename,
                    new_filename="",
                    original_index=file_item.original_index,
                    new_index=None,
                    status="NG",
                )
            )
            continue

        next_index = _next_available_index(next_index, missing_indices)
        text_value = None if text_overrides is None else text_overrides.get(file_item.original_filename, file_item.text_portion)
        new_filename = _build_filename(file_item, next_index, settings.digits, settings.keep_text, text_value=text_value)
        plan.append(
            RenamePlanEntry(
                source_path=file_item.path,
                original_filename=file_item.original_filename,
                new_filename=new_filename,
                original_index=file_item.original_index,
                new_index=next_index,
                status="OK",
            )
        )
        next_index += 1

    plan.sort(
        key=lambda item: (
            item.original_index,
            0 if item.status == "MISSING" else 1,
            item.original_filename.lower(),
        )
    )
    return plan


def write_undo_manifest(plan: list[RenamePlanEntry], folder: Path, settings: RenameSettings) -> Path:
    manifest_path = folder / UNDO_MANIFEST_NAME
    payload = {
        "settings": asdict(settings),
        "entries": [
            {
                "original_filename": entry.original_filename,
                "new_filename": entry.new_filename,
                "status": entry.status,
            }
            for entry in plan
        ],
    }
    manifest_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_path


def has_undo_manifest(folder: Path) -> bool:
    return (folder / UNDO_MANIFEST_NAME).exists()


def undo_last_rename(folder: Path, progress_callback: Callable[[float], None] | None = None) -> None:
    manifest_path = folder / UNDO_MANIFEST_NAME
    if not manifest_path.exists():
        raise FileNotFoundError("元に戻せる履歴がありません。")

    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    settings = RenameSettings(**payload["settings"])
    operations: list[tuple[Path, Path]] = []

    for entry in payload["entries"]:
        status = entry["status"]
        if status == "OK":
            source = folder / entry["new_filename"]
            target = folder / entry["original_filename"]
        elif status == "NG" and settings.move_ng_files:
            source = folder / settings.ng_folder_name / entry["original_filename"]
            target = folder / entry["original_filename"]
        else:
            continue
        operations.append((source, target))

    total_steps = max(len(operations), 1)
    temp_records: list[tuple[Path, Path, Path]] = []

    try:
        for step, (source, target) in enumerate(operations, start=1):
            if not source.exists():
                raise FileNotFoundError(f"元に戻す対象が見つかりません: {source.name}")
            temp_path = folder / f".undo_tmp_{step:04d}_{source.name}"
            source.rename(temp_path)
            temp_records.append((temp_path, source, target))
            if progress_callback:
                progress_callback(step / (total_steps * 2))

        for step, (temp_path, _source, target) in enumerate(temp_records, start=1):
            if target.exists():
                raise FileExistsError(f"元に戻し先が既に存在します: {target.name}")
            temp_path.rename(target)
            move_trim_backup_reference(_source, target)
            if progress_callback:
                progress_callback((total_steps + step) / (total_steps * 2))
    except Exception:
        for temp_path, source, target in reversed(temp_records):
            if temp_path.exists():
                temp_path.rename(source)
            elif target.exists() and not source.exists():
                target.rename(source)
                move_trim_backup_reference(target, source)
        raise

    manifest_path.unlink(missing_ok=True)


def execute_rename_plan(
    plan: list[RenamePlanEntry],
    folder: Path,
    settings: RenameSettings,
    progress_callback: Callable[[float], None] | None = None,
) -> None:
    operations = [entry for entry in plan if entry.source_path is not None]
    source_paths = {
        entry.source_path.resolve()
        for entry in operations
        if entry.source_path is not None
    }
    total_steps = max(len(operations), 1)
    temp_records: list[tuple[Path, Path, RenamePlanEntry]] = []

    ng_folder = folder / settings.ng_folder_name
    if settings.move_ng_files:
        ng_folder.mkdir(exist_ok=True)

    seen_destinations: set[Path] = set()
    for entry in operations:
        source_path = entry.source_path
        assert source_path is not None
        if entry.status == "NG":
            destination = ng_folder / source_path.name if settings.move_ng_files else folder / source_path.name
        else:
            destination = folder / entry.new_filename

        if destination in seen_destinations:
            raise FileExistsError(f"同じ新ファイル名が複数あります: {destination.name}")
        seen_destinations.add(destination)

        if destination.exists() and destination.resolve() not in source_paths:
            raise FileExistsError(
                "処理対象外の既存ファイルと名前が衝突しています: "
                f"{destination.name}\n\n"
                "同じフォルダ内に、今回の処理対象に含まれていない同名ファイルが残っています。"
            )

    try:
        for step, entry in enumerate(operations, start=1):
            source_path = entry.source_path
            assert source_path is not None
            temp_path = folder / f".renaming_tmp_{step:04d}_{source_path.name}"
            source_path.rename(temp_path)
            temp_records.append((temp_path, source_path, entry))
            if progress_callback:
                progress_callback(step / (total_steps * 2))

        for step, (temp_path, original_path, entry) in enumerate(temp_records, start=1):
            if entry.status == "NG":
                destination = ng_folder / original_path.name if settings.move_ng_files else folder / original_path.name
            else:
                destination = folder / entry.new_filename

            if destination.exists():
                raise FileExistsError(
                    "新しいファイル名の作成先がすでに埋まっています: "
                    f"{destination.name}\n\n"
                    "処理対象外のファイルが残っているか、前回の途中結果が残っている可能性があります。"
                )

            if destination.parent != folder and not destination.parent.exists():
                destination.parent.mkdir(parents=True, exist_ok=True)

            if entry.status == "NG" and settings.move_ng_files:
                shutil.move(str(temp_path), str(destination))
            else:
                temp_path.rename(destination)

            move_trim_backup_reference(original_path, destination)

            if progress_callback:
                progress_callback((total_steps + step) / (total_steps * 2))
    except Exception:
        for temp_path, original_path, entry in reversed(temp_records):
            if temp_path.exists():
                temp_path.rename(original_path)
                continue

            if entry.status == "NG" and settings.move_ng_files:
                moved_path = ng_folder / original_path.name
                if moved_path.exists():
                    shutil.move(str(moved_path), str(original_path))
                    move_trim_backup_reference(moved_path, original_path)
            elif entry.status == "OK":
                renamed_path = folder / entry.new_filename
                if renamed_path.exists():
                    renamed_path.rename(original_path)
                    move_trim_backup_reference(renamed_path, original_path)
        raise
