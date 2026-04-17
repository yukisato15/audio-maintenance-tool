from __future__ import annotations

from array import array
from dataclasses import dataclass
from pathlib import Path
import json
import math
import re
import tempfile
import uuid
import wave


TRIM_BACKUP_PREFIX = ".trim_backup_"
SPLIT_BACKUP_PREFIX = ".split_backup_"
TRIM_BACKUP_MANIFEST = ".trim_backup_manifest.json"
PREFIX_PATTERN = re.compile(r"^(\d+)")


@dataclass(slots=True)
class AudioMetadata:
    sample_rate: int
    channels: int
    sample_width: int
    frame_count: int

    @property
    def duration_ms(self) -> int:
        if self.sample_rate <= 0:
            return 0
        return int(self.frame_count * 1000 / self.sample_rate)


@dataclass(slots=True)
class AudioLevelStats:
    peak_db: float
    rms_db: float
    clipping_ratio: float
    silent_ratio: float


@dataclass(slots=True)
class AudioProcessOptions:
    smooth_edges: bool = True
    fade_ms: int = 400


def _backup_path(path: Path) -> Path:
    manifest = _load_trim_manifest(path.parent)
    backup_name = manifest.get(path.name)
    if not backup_name:
        return path.with_name(f"{TRIM_BACKUP_PREFIX}{path.name}")
    return path.parent / backup_name


def _legacy_backup_path(path: Path) -> Path:
    return path.with_name(f"{TRIM_BACKUP_PREFIX}{path.name}")


def _manifest_path(folder: Path) -> Path:
    return folder / TRIM_BACKUP_MANIFEST


def _load_trim_manifest(folder: Path) -> dict[str, str]:
    manifest_path = _manifest_path(folder)
    if not manifest_path.exists():
        return {}
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    files = payload.get("files", {})
    if not isinstance(files, dict):
        return {}
    return {str(key): str(value) for key, value in files.items()}


def _save_trim_manifest(folder: Path, mapping: dict[str, str]) -> None:
    manifest_path = _manifest_path(folder)
    cleaned = {key: value for key, value in mapping.items() if value}
    if not cleaned:
        manifest_path.unlink(missing_ok=True)
        return
    payload = {"files": cleaned}
    manifest_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _generate_trim_backup_name(path: Path) -> str:
    return f"{TRIM_BACKUP_PREFIX}{uuid.uuid4().hex}_{path.name}"


def _find_trim_backup(path: Path) -> tuple[Path | None, dict[str, str]]:
    manifest = _load_trim_manifest(path.parent)
    backup_name = manifest.get(path.name)
    if backup_name:
        backup_path = path.parent / backup_name
        if backup_path.exists():
            return backup_path, manifest
        manifest.pop(path.name, None)

    legacy_backup = _legacy_backup_path(path)
    if legacy_backup.exists():
        return legacy_backup, manifest
    return None, manifest


def _ensure_trim_backup(path: Path) -> Path:
    backup_path, manifest = _find_trim_backup(path)
    if backup_path is not None:
        if manifest.get(path.name) != backup_path.name:
            manifest[path.name] = backup_path.name
            _save_trim_manifest(path.parent, manifest)
        return backup_path

    backup_name = _generate_trim_backup_name(path)
    backup_path = path.parent / backup_name
    path.rename(backup_path)
    manifest[path.name] = backup_name
    _save_trim_manifest(path.parent, manifest)
    return backup_path


def _remove_trim_backup(path: Path) -> None:
    manifest = _load_trim_manifest(path.parent)
    backup_name = manifest.pop(path.name, None)
    if backup_name:
        (path.parent / backup_name).unlink(missing_ok=True)
    _legacy_backup_path(path).unlink(missing_ok=True)
    _save_trim_manifest(path.parent, manifest)


def move_trim_backup_reference(source: Path, destination: Path) -> None:
    source_manifest = _load_trim_manifest(source.parent)
    backup_name = source_manifest.pop(source.name, None)
    if not backup_name:
        legacy_backup = _legacy_backup_path(source)
        if not legacy_backup.exists():
            _save_trim_manifest(source.parent, source_manifest)
            return
        backup_name = _generate_trim_backup_name(source)
        backup_source = legacy_backup
    else:
        backup_source = source.parent / backup_name

    destination.parent.mkdir(parents=True, exist_ok=True)
    backup_destination = destination.parent / backup_name
    if backup_source.exists() and backup_source != backup_destination:
        if backup_destination.exists():
            raise FileExistsError(f"Trim backup destination already exists: {backup_destination.name}")
        backup_source.rename(backup_destination)

    _save_trim_manifest(source.parent, source_manifest)
    destination_manifest = _load_trim_manifest(destination.parent)
    destination_manifest[destination.name] = backup_name
    _save_trim_manifest(destination.parent, destination_manifest)


def _split_backup_path(path: Path) -> Path:
    return path.with_name(f"{SPLIT_BACKUP_PREFIX}{path.name}")


def _split_output_path(path: Path, segment_number: int) -> Path:
    match = PREFIX_PATTERN.match(path.stem)
    prefix = match.group(1) if match else path.stem
    return path.with_name(f"{prefix}__split{segment_number:02d}{path.suffix.lower()}")


def split_audio_in_place(
    path: Path,
    split_points_ms: list[int],
    options: AudioProcessOptions | None = None,
) -> list[Path]:
    metadata = get_audio_metadata(path)
    if metadata.duration_ms <= 1:
        raise ValueError("分割できる長さがありません。")

    clean_points = sorted({point for point in split_points_ms if 0 < point < metadata.duration_ms})
    if not clean_points:
        raise ValueError("分割位置がありません。")

    _remove_trim_backup(path)

    split_backup = _split_backup_path(path)
    if split_backup.exists():
        raise FileExistsError("このファイルには未確定の分割バックアップがあります。")

    path.rename(split_backup)
    created_paths: list[Path] = []
    try:
        boundaries = [0, *clean_points, metadata.duration_ms]
        for segment_number, (start_ms, end_ms) in enumerate(zip(boundaries, boundaries[1:]), start=1):
            if end_ms <= start_ms:
                continue
            destination = _split_output_path(path, segment_number)
            if destination.exists():
                raise FileExistsError(f"分割後ファイルが既に存在します: {destination.name}")
            _write_trimmed_audio(split_backup, destination, start_ms, metadata.duration_ms - end_ms, options=options)
            created_paths.append(destination)
        if len(created_paths) < 2:
            raise ValueError("2つ以上のセグメントに分割できませんでした。")
    except Exception:
        for created_path in created_paths:
            created_path.unlink(missing_ok=True)
        if split_backup.exists() and not path.exists():
            split_backup.rename(path)
        raise
    return created_paths


def has_split_backup(path: Path) -> bool:
    return _split_backup_path(path).exists()

def has_trim_backup(path: Path) -> bool:
    backup_path, _manifest = _find_trim_backup(path)
    return backup_path is not None and backup_path.exists()


def get_audio_metadata(path: Path) -> AudioMetadata:
    with wave.open(str(path), "rb") as wav_file:
        return AudioMetadata(
            sample_rate=wav_file.getframerate(),
            channels=wav_file.getnchannels(),
            sample_width=wav_file.getsampwidth(),
            frame_count=wav_file.getnframes(),
        )


def _resolve_source(path: Path) -> Path:
    backup = _backup_path(path)
    return backup if backup.exists() else path


def _read_normalized_samples(path: Path, use_backup_if_available: bool = False) -> tuple[list[float], int, int]:
    source = _resolve_source(path) if use_backup_if_available else path
    with wave.open(str(source), "rb") as wav_file:
        channels = wav_file.getnchannels()
        sample_width = wav_file.getsampwidth()
        frame_count = wav_file.getnframes()
        raw_frames = wav_file.readframes(frame_count)

    if frame_count <= 0 or not raw_frames:
        return [], channels, frame_count

    if sample_width == 1:
        samples = array("B", raw_frames)
        values = [(sample - 128) / 128 for sample in samples]
    elif sample_width == 2:
        samples = array("h")
        samples.frombytes(raw_frames)
        values = [sample / 32768 for sample in samples]
    elif sample_width == 4:
        samples = array("i")
        samples.frombytes(raw_frames)
        values = [sample / 2147483648 for sample in samples]
    else:
        return [], channels, frame_count

    return values, channels, frame_count


def get_waveform_minmax(path: Path, bucket_count: int = 240) -> list[tuple[float, float]]:
    if bucket_count <= 0:
        return []

    values, channels, frame_count = _read_normalized_samples(path, use_backup_if_available=False)
    if frame_count <= 0 or not values:
        return [(0.0, 0.0)] * bucket_count

    minmax: list[tuple[float, float] | None] = [None] * bucket_count
    for frame_index in range(frame_count):
        bucket_index = min(int(frame_index * bucket_count / frame_count), bucket_count - 1)
        frame_min = 1.0
        frame_max = -1.0
        base = frame_index * channels
        for channel in range(channels):
            sample_index = base + channel
            if sample_index >= len(values):
                break
            amplitude = max(-1.0, min(1.0, values[sample_index]))
            frame_min = min(frame_min, amplitude)
            frame_max = max(frame_max, amplitude)
        current = minmax[bucket_index]
        if current is None:
            minmax[bucket_index] = (frame_min, frame_max)
        else:
            minmax[bucket_index] = (min(current[0], frame_min), max(current[1], frame_max))

    return [pair if pair is not None else (0.0, 0.0) for pair in minmax]


def analyze_audio_levels(path: Path) -> AudioLevelStats:
    values, channels, frame_count = _read_normalized_samples(path, use_backup_if_available=False)
    if frame_count <= 0 or not values:
        return AudioLevelStats(peak_db=-120.0, rms_db=-120.0, clipping_ratio=0.0, silent_ratio=1.0)

    frame_peaks: list[float] = []
    squared_sum = 0.0
    clip_count = 0
    silent_count = 0

    for frame_index in range(frame_count):
        base = frame_index * channels
        frame_peak = 0.0
        for channel in range(channels):
            sample_index = base + channel
            if sample_index >= len(values):
                break
            amplitude = abs(values[sample_index])
            squared_sum += values[sample_index] * values[sample_index]
            frame_peak = max(frame_peak, amplitude)
        frame_peaks.append(frame_peak)
        if frame_peak >= 0.995:
            clip_count += 1
        if frame_peak <= 0.01:
            silent_count += 1

    peak = max(frame_peaks, default=0.0)
    rms = (squared_sum / max(len(values), 1)) ** 0.5
    return AudioLevelStats(
        peak_db=20 * math.log10(max(peak, 1e-6)),
        rms_db=20 * math.log10(max(rms, 1e-6)),
        clipping_ratio=clip_count / max(frame_count, 1),
        silent_ratio=silent_count / max(frame_count, 1),
    )


def _trim_frame_range(frame_count: int, sample_rate: int, trim_start_ms: int, trim_end_ms: int) -> tuple[int, int]:
    start_frame = max(0, int(sample_rate * trim_start_ms / 1000))
    end_frame = max(0, int(sample_rate * trim_end_ms / 1000))
    start_frame = min(start_frame, frame_count)
    end_frame = min(end_frame, max(frame_count - start_frame - 1, 0))
    remaining = max(frame_count - start_frame - end_frame, 1)
    return start_frame, remaining


def _decode_pcm(raw_frames: bytes, sample_width: int) -> list[int]:
    if sample_width == 1:
        return [value - 128 for value in raw_frames]
    if sample_width == 2:
        values = array("h")
        values.frombytes(raw_frames)
        return list(values)
    if sample_width == 4:
        values = array("i")
        values.frombytes(raw_frames)
        return list(values)
    raise ValueError("この WAV のビット深度にはまだ対応していません。")


def _encode_pcm(samples: list[int], sample_width: int) -> bytes:
    if sample_width == 1:
        return bytes(max(0, min(255, value + 128)) for value in samples)
    if sample_width == 2:
        clipped = array("h", (max(-32768, min(32767, value)) for value in samples))
        return clipped.tobytes()
    if sample_width == 4:
        clipped = array("i", (max(-2147483648, min(2147483647, value)) for value in samples))
        return clipped.tobytes()
    raise ValueError("この WAV のビット深度にはまだ対応していません。")


def _frame_signal(samples: list[int], channels: int, frame_index: int) -> float:
    base = frame_index * channels
    frame = samples[base:base + channels]
    if not frame:
        return 0.0
    return sum(frame) / len(frame)


def _find_zero_cross_frame(
    samples: list[int],
    channels: int,
    target_frame: int,
    min_frame: int,
    max_frame: int,
    search_window_frames: int,
    direction: str = "nearest",
) -> int:
    if direction == "forward":
        low = max(min_frame + 1, target_frame)
        high = min(max_frame - 1, target_frame + search_window_frames)
    elif direction == "backward":
        low = max(min_frame + 1, target_frame - search_window_frames)
        high = min(max_frame - 1, target_frame)
    else:
        low = max(min_frame + 1, target_frame - search_window_frames)
        high = min(max_frame - 1, target_frame + search_window_frames)
    if high < low:
        return max(min_frame, min(target_frame, max_frame))

    best_frame = target_frame
    best_score: tuple[int, float] | None = None
    for frame_index in range(low, high + 1):
        previous_signal = _frame_signal(samples, channels, frame_index - 1)
        current_signal = _frame_signal(samples, channels, frame_index)
        crosses_zero = (previous_signal <= 0 <= current_signal) or (previous_signal >= 0 >= current_signal)
        if not crosses_zero and abs(current_signal) > 64:
            continue
        score = (abs(frame_index - target_frame), abs(current_signal))
        if best_score is None or score < best_score:
            best_score = score
            best_frame = frame_index
    return max(min_frame, min(best_frame, max_frame))


def _apply_fade(samples: list[int], channels: int, sample_rate: int, fade_ms: int, sample_width: int) -> None:
    if fade_ms <= 0 or not samples:
        return
    frame_count = len(samples) // max(channels, 1)
    if frame_count <= 1:
        return
    fade_frames = max(1, int(sample_rate * fade_ms / 1000))
    fade_frames = min(fade_frames, frame_count // 2)
    if fade_frames <= 0:
        return
    for frame_index in range(fade_frames):
        if fade_frames == 1:
            in_gain = 0.0
            out_gain = 0.0
        else:
            in_gain = frame_index / (fade_frames - 1)
            out_gain = (fade_frames - 1 - frame_index) / (fade_frames - 1)
        start_base = frame_index * channels
        end_base = (frame_count - fade_frames + frame_index) * channels
        for channel in range(channels):
            start_index = start_base + channel
            end_index = end_base + channel
            samples[start_index] = int(samples[start_index] * in_gain)
            samples[end_index] = int(samples[end_index] * out_gain)


def _write_trimmed_audio(
    source: Path,
    destination: Path,
    trim_start_ms: int,
    trim_end_ms: int,
    options: AudioProcessOptions | None = None,
) -> None:
    options = options or AudioProcessOptions()
    with wave.open(str(source), "rb") as src:
        params = src.getparams()
        sample_rate = src.getframerate()
        channels = src.getnchannels()
        sample_width = src.getsampwidth()
        total_frames = src.getnframes()
        raw_frames = src.readframes(total_frames)

    start_frame, remaining_frames = _trim_frame_range(total_frames, sample_rate, trim_start_ms, trim_end_ms)
    end_frame = min(start_frame + remaining_frames, total_frames)
    samples = _decode_pcm(raw_frames, sample_width)

    if options.smooth_edges:
        search_window_frames = max(1, int(sample_rate * 12 / 1000))
        start_frame = _find_zero_cross_frame(
            samples,
            channels,
            start_frame,
            0,
            max(end_frame - 1, 1),
            search_window_frames,
            direction="forward",
        )
        end_frame = _find_zero_cross_frame(
            samples,
            channels,
            end_frame,
            min(start_frame + 1, total_frames - 1),
            total_frames - 1,
            search_window_frames,
            direction="backward",
        )
        end_frame = max(start_frame + 1, min(end_frame, total_frames))

    trimmed_samples = samples[start_frame * channels:end_frame * channels]
    if options.smooth_edges:
        _apply_fade(trimmed_samples, channels, sample_rate, max(0, options.fade_ms), sample_width)
    frames = _encode_pcm(trimmed_samples, sample_width)

    with wave.open(str(destination), "wb") as dst:
        dst.setparams(params)
        dst.writeframes(frames)


def create_trim_preview(
    path: Path,
    trim_start_ms: int,
    trim_end_ms: int,
    options: AudioProcessOptions | None = None,
) -> Path:
    # 視聴は常に「いま画面に出ている現在ファイル」を基準にする。
    # バックアップ元を使うと、適用後の再視聴で内容が古くなり波形とずれる。
    source = path
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".wav")
    temp_file.close()
    destination = Path(temp_file.name)
    _write_trimmed_audio(source, destination, trim_start_ms, trim_end_ms, options=options)
    return destination


def attenuate_audio_in_place(path: Path, gain_db: float = -3.0) -> None:
    factor = 10 ** (gain_db / 20)
    source = path
    temp_destination = path.with_name(f".trim_tmp_{path.name}")

    with wave.open(str(source), "rb") as src:
        params = src.getparams()
        channels = src.getnchannels()
        sample_width = src.getsampwidth()
        frame_count = src.getnframes()
        raw_frames = src.readframes(frame_count)

    if sample_width == 1:
        samples = array("B", raw_frames)
        adjusted = bytearray()
        for sample in samples:
            centered = sample - 128
            value = int(centered * factor)
            value = max(-128, min(127, value))
            adjusted.append(value + 128)
        output_bytes = bytes(adjusted)
    elif sample_width == 2:
        samples = array("h")
        samples.frombytes(raw_frames)
        adjusted = array("h", (max(-32768, min(32767, int(sample * factor))) for sample in samples))
        output_bytes = adjusted.tobytes()
    elif sample_width == 4:
        samples = array("i")
        samples.frombytes(raw_frames)
        adjusted = array("i", (max(-2147483648, min(2147483647, int(sample * factor))) for sample in samples))
        output_bytes = adjusted.tobytes()
    else:
        raise ValueError("この WAV のビット深度にはまだ対応していません。")

    with wave.open(str(temp_destination), "wb") as dst:
        dst.setparams(params)
        dst.writeframes(output_bytes)

    if path.exists():
        path.unlink()
    temp_destination.rename(path)


def apply_trim_in_place(
    path: Path,
    trim_start_ms: int,
    trim_end_ms: int,
    options: AudioProcessOptions | None = None,
) -> Path:
    backup = _ensure_trim_backup(path)
    source = backup
    temp_destination = path.with_name(f".trim_tmp_{path.name}")
    _write_trimmed_audio(source, temp_destination, trim_start_ms, trim_end_ms, options=options)
    if path.exists():
        path.unlink()
    temp_destination.rename(path)
    return backup


def restore_trim_backup(path: Path) -> None:
    manifest = _load_trim_manifest(path.parent)
    backup, _ = _find_trim_backup(path)
    if backup is None:
        backup = _backup_path(path)
    if not backup.exists():
        raise FileNotFoundError("元に戻せる余白修正バックアップがありません。")
    if path.exists():
        path.unlink()
    backup.rename(path)
    manifest.pop(path.name, None)
    _save_trim_manifest(path.parent, manifest)
